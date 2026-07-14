"""
Telegram bot: Excel arama kayıtlarını personel bazlı tekrar analizi.

Sadece ALLOWED_CHAT_IDS ile izin verilen grup(lar)da çalışır.
Özel sohbet ve diğer gruplarda sessiz kalır.

Kullanım:
  1. .env / Railway: TELEGRAM_BOT_TOKEN + ALLOWED_CHAT_IDS
  2. pip install -r requirements.txt
  3. python bot.py
  4. Botu izinli gruba ekleyin; Group Privacy kapalı veya bot admin olsun
"""

from __future__ import annotations

import logging
import os
import re
import tempfile
from pathlib import Path

from dotenv import load_dotenv
from telegram import Document, Update
from telegram.constants import ChatAction, ChatType
from telegram.error import BadRequest
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from analyzer import analyze_workbook, build_report_workbook, format_text_summary

# .env her zaman bot.py yanından yüklensin (çalışma dizininden bağımsız)
_BASE_DIR = Path(__file__).resolve().parent
load_dotenv(_BASE_DIR / ".env")

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("tekrar-bot")

# Varsayılan: aynı personel aynı numarayı en az 2 kez aramışsa "tekrar"
DEFAULT_MIN_REPEAT = int(os.getenv("MIN_REPEAT", "2"))
MAX_FILE_MB = float(os.getenv("MAX_FILE_MB", "20"))

# Yalnızca grup + süpergrup (ek olarak chat id allowlist kontrolü var)
GROUP_FILTER = filters.ChatType.GROUPS


def parse_allowed_chat_ids(*raw_values: str | None) -> frozenset[int]:
    """
    ' -100123, -100456 ' veya tek id → frozenset[int].
    Virgül / noktalı virgül / boşluk ayraçlarını kabul eder.
    """
    ids: set[int] = set()
    for raw in raw_values:
        if not raw:
            continue
        for part in re.split(r"[,;\s]+", str(raw).strip()):
            if not part:
                continue
            try:
                ids.add(int(part))
            except ValueError as exc:
                raise ValueError(
                    f"Geçersiz chat id: {part!r}. "
                    "Örnek: ALLOWED_CHAT_IDS=-1001234567890"
                ) from exc
    return frozenset(ids)


def load_allowed_chat_ids_from_env() -> frozenset[int]:
    """ALLOWED_CHAT_IDS ve/veya ALLOWED_CHAT_ID env değerlerini okur."""
    return parse_allowed_chat_ids(
        os.getenv("ALLOWED_CHAT_IDS"),
        os.getenv("ALLOWED_CHAT_ID"),
    )


# Başlangıçta yüklenir; testler ALLOWED_CHAT_IDS'i yeniden atayabilir
ALLOWED_CHAT_IDS: frozenset[int] = load_allowed_chat_ids_from_env()


HELP_TEXT = """\
Yardım — Komut Listesi

Bu bot yalnızca izin verilen grupta çalışır (ALLOWED_CHAT_IDS).
Özel sohbet ve diğer gruplarda yanıt vermez.

/start — Botu başlatır, kısa tanıtım ve kullanım bilgisini gösterir.

/yardim — Tüm komutları açıklamalarıyla listeler (bu mesaj).

/help — /yardim ile aynıdır.

/chatid — Bu grubun Chat ID bilgisini gösterir.

/esik — Bu gruptaki tekrar eşiğini gösterir veya ayarlar.
  • Kullanım: /esik → mevcut eşiği gösterir
  • Kullanım: /esik 3 → aynı personelin aynı numarayı en az 3 kez araması “tekrar” sayılır
  • Minimum değer: 2
  • Eşik gruba özeldir (tüm üyeler aynı eşiği kullanır)

Dosya gönderme
Excel (.xlsx / .xlsm) dosyasını gruba gönderin.
Bot personel bazlı tekrarlı arama özeti + detaylı Excel rapor döner.

Not: Bot dosyaları görebilsin diye BotFather → /setprivacy → Disable
veya botu gruba yönetici olarak ekleyin.

Excel sütunları
• A — Telefon numarası
• B — Arama tarihi
• C — Arama saati
• E — Konuşma süresi
• F — Çaldırma süresi
• G — Personel adı (Seda -O vb.; K / - gibi kayıtlar elenir)

Kurallar
• Tekrarlar yalnızca aynı personel içinde sayılır
• Farklı personellerin aynı numarayı araması birbirini etkilemez
• Türkçe / İngilizce karakterler (İ, ı, Ş, ğ …) doğru işlenir
"""


def is_group_chat(update: Update) -> bool:
    """Grup veya süpergrup mu? (özel sohbet = False)."""
    chat = update.effective_chat
    if not chat:
        return False
    return chat.type in (ChatType.GROUP, ChatType.SUPERGROUP)


def is_allowed_chat(update: Update) -> bool:
    """
    İzinli grup mu?
    - Özel sohbet: hayır
    - Grup ama chat id listede yok: hayır (sessiz; log yazar)
    """
    if not is_group_chat(update):
        return False
    chat = update.effective_chat
    if not chat:
        return False
    if not ALLOWED_CHAT_IDS:
        return False
    if chat.id not in ALLOWED_CHAT_IDS:
        logger.info(
            "Yetkisiz grup yok sayıldı | chat_id=%s | title=%r",
            chat.id,
            getattr(chat, "title", None),
        )
        return False
    return True


def get_min_repeat(context: ContextTypes.DEFAULT_TYPE) -> int:
    """Grup bazlı eşik (chat_data)."""
    return int(context.chat_data.get("min_repeat", DEFAULT_MIN_REPEAT))


def safe_filename(name: str | None, default: str = "dosya.xlsx") -> str:
    """Telegram dosya adını güvenli hale getirir (path traversal engeli)."""
    raw = (name or default).replace("\\", "/").split("/")[-1].strip()
    raw = re.sub(r"[^\w.\- ()\[\]]+", "_", raw, flags=re.UNICODE)
    raw = raw.strip(" .") or default
    if not raw.lower().endswith((".xlsx", ".xlsm")):
        raw = raw + ".xlsx"
    return raw[:180]


async def reply_plain(message, text: str) -> None:
    await message.reply_text(text)


async def edit_plain(message, text: str) -> None:
    try:
        await message.edit_text(text)
    except BadRequest as exc:
        if "not modified" not in str(exc).lower():
            logger.warning("edit_text başarısız: %s", exc)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed_chat(update) or not update.message:
        return
    min_r = get_min_repeat(context)
    await reply_plain(
        update.message,
        "Merhaba\n\n"
        "Bu bot yalnızca izin verilen bu grupta çalışır "
        "(özel sohbet ve diğer gruplarda yanıt vermez).\n\n"
        "Arama Excel dosyanızı personel bazlı inceler:\n"
        "• Hangi personel\n"
        "• Hangi numarayı\n"
        "• Hangi tarih/saatlerde\n"
        "• Kaç kez aramış\n"
        "• Konuşma / çaldırma süreleri\n\n"
        f"Bu grubun tekrar eşiği: en az {min_r} arama\n"
        "(/esik 3 ile değiştirebilirsiniz)\n\n"
        "Analiz için .xlsx dosyasını gruba gönderin.\n\n"
        "Tüm komutlar için: /yardim\n\n"
        "Kurallar:\n"
        "• G sütunu personel adı (Seda -O vb.); K / - gibi kayıtlar elenir\n"
        "• Tekrarlar aynı personel içinde sayılır\n"
        "• A=telefon, B=tarih, C=saat, E=konuşma, F=çaldırma, G=personel\n"
        "• Türkçe/İngilizce karakterler desteklenir",
    )


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed_chat(update) or not update.message:
        return
    min_r = get_min_repeat(context)
    text = HELP_TEXT + f"\nBu grubun mevcut eşiği: {min_r}"
    await reply_plain(update.message, text)


async def set_threshold(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed_chat(update) or not update.message:
        return
    if not context.args:
        current = get_min_repeat(context)
        await reply_plain(
            update.message,
            f"Bu grubun mevcut eşiği: {current}\nKullanım: /esik 2",
        )
        return
    try:
        value = int(context.args[0])
        if value < 2:
            raise ValueError
    except ValueError:
        await reply_plain(update.message, "Eşik en az 2 olmalı. Örnek: /esik 2")
        return
    context.chat_data["min_repeat"] = value
    await reply_plain(
        update.message,
        f"Bu grup için tekrar eşiği {value} olarak ayarlandı. "
        "Yeni Excel gönderildiğinde bu eşik kullanılır.",
    )


async def chat_id_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    İzinli grupta tam bilgi verir.
    Yetkisiz grupta da sadece Chat ID döner (Railway'e yazmak için kurulum kolaylığı).
    Özel sohbette sessiz.
    """
    if not is_group_chat(update) or not update.message:
        return
    chat = update.effective_chat
    user = update.effective_user
    if not chat:
        return

    allowed = bool(ALLOWED_CHAT_IDS and chat.id in ALLOWED_CHAT_IDS)

    lines = [
        "Chat ID Bilgisi",
        "",
        f"Chat ID: {chat.id}",
        f"Sohbet türü: {chat.type}",
    ]
    if chat.title:
        lines.append(f"Sohbet adı: {chat.title}")
    if allowed and user:
        lines.append(f"User ID: {user.id}")
        if user.username:
            lines.append(f"Kullanıcı: @{user.username}")
        name = " ".join(p for p in [user.first_name, user.last_name] if p)
        if name:
            lines.append(f"İsim: {name}")
    if allowed:
        lines.append("Durum: bu grup izin listesinde (bot burada çalışır)")
    else:
        lines.append("Durum: bu grup henüz izinli değil")
        lines.append(
            "Railway / .env içine yazın:\n"
            f"ALLOWED_CHAT_IDS={chat.id}"
        )

    await reply_plain(update.message, "\n".join(lines))


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed_chat(update) or not update.message:
        return
    doc: Document | None = update.message.document
    if not doc:
        return

    file_name = safe_filename(doc.file_name)
    lower = file_name.lower()
    if not (lower.endswith(".xlsx") or lower.endswith(".xlsm")):
        await reply_plain(
            update.message,
            "Lütfen .xlsx formatında Excel gönderin (eski .xls desteklenmez).",
        )
        return

    if doc.file_size and doc.file_size > MAX_FILE_MB * 1024 * 1024:
        await reply_plain(
            update.message,
            f"Dosya çok büyük (max {MAX_FILE_MB:.0f} MB).",
        )
        return

    min_repeat = get_min_repeat(context)
    status = await update.message.reply_text(
        f"{file_name} alınıyor ve analiz ediliyor…\n"
        f"(tekrar eşiği: ≥{min_repeat})"
    )

    try:
        await context.bot.send_chat_action(
            chat_id=update.effective_chat.id,
            action=ChatAction.TYPING,
        )
        tg_file = await doc.get_file()

        with tempfile.TemporaryDirectory(prefix="tekrar_bot_") as tmp:
            src_path = Path(tmp) / file_name
            await tg_file.download_to_drive(custom_path=str(src_path))

            reports = analyze_workbook(src_path, min_repeat=min_repeat)
            summary = format_text_summary(reports)
            report_buf = build_report_workbook(reports, min_repeat=min_repeat)

            out_name = f"tekrar_rapor_{Path(file_name).stem[:80]}.xlsx"
            out_name = safe_filename(out_name)
            out_path = Path(tmp) / out_name
            out_path.write_bytes(report_buf.getvalue())

            await edit_plain(status, summary)

            await context.bot.send_chat_action(
                chat_id=update.effective_chat.id,
                action=ChatAction.UPLOAD_DOCUMENT,
            )
            with out_path.open("rb") as f:
                await update.message.reply_document(
                    document=f,
                    filename=out_name,
                    caption=(
                        "Detaylı personel bazlı rapor\n"
                        "Sayfalar: Bilgi | Ozet | TekrarliAramalar | personel adları"
                    ),
                )
    except Exception:
        logger.exception("Analiz hatası")
        await edit_plain(
            status,
            "Dosya işlenirken hata oluştu.\n"
            "Sütunların A/B/C/E/F/G düzeninde ve .xlsx olduğundan emin olun.",
        )


async def handle_non_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed_chat(update) or not update.message:
        return
    if update.message.text and update.message.text.startswith("/"):
        return
    await reply_plain(
        update.message,
        "Analiz için gruba bir Excel (.xlsx) dosyası gönderin.\n"
        "Komut listesi: /yardim",
    )


async def post_init(app: Application) -> None:
    from telegram import BotCommand

    await app.bot.set_my_commands(
        [
            BotCommand("start", "Botu başlat / kısa tanıtım"),
            BotCommand("yardim", "Tüm komutlar ve açıklamaları"),
            BotCommand("help", "Yardım (yardim ile aynı)"),
            BotCommand("chatid", "Bu grubun Chat ID bilgisini göster"),
            BotCommand("esik", "Grup tekrar eşiğini gör / ayarla"),
        ]
    )


def main() -> None:
    global ALLOWED_CHAT_IDS

    token = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
    if not token:
        raise SystemExit(
            "TELEGRAM_BOT_TOKEN bulunamadı.\n"
            "Railway Variables / .env:\n"
            "  TELEGRAM_BOT_TOKEN=123456:ABC-DEF...\n"
        )

    try:
        ALLOWED_CHAT_IDS = load_allowed_chat_ids_from_env()
    except ValueError as exc:
        raise SystemExit(str(exc)) from exc

    if not ALLOWED_CHAT_IDS:
        raise SystemExit(
            "ALLOWED_CHAT_IDS zorunludur. Bot yalnızca izin verdiğiniz grupta çalışır.\n"
            "Railway Variables / .env örneği:\n"
            "  ALLOWED_CHAT_IDS=-1001234567890\n"
            "Birden fazla grup: ALLOWED_CHAT_IDS=-100111,-100222\n"
            "Chat ID öğrenmek: botu gruba ekleyip /chatid yazın "
            "(bu komut kurulum için her grupta çalışır)."
        )

    app = (
        Application.builder()
        .token(token)
        .concurrent_updates(True)
        .post_init(post_init)
        .build()
    )

    # Özel sohbet: handler bağlanmaz → hiç yanıt yok
    # Yetkisiz gruplar: is_allowed_chat içinde sessizce elenir (/chatid hariç)
    app.add_handler(CommandHandler("start", start, filters=GROUP_FILTER))
    app.add_handler(CommandHandler("help", help_cmd, filters=GROUP_FILTER))
    app.add_handler(CommandHandler("yardim", help_cmd, filters=GROUP_FILTER))
    app.add_handler(CommandHandler("chatid", chat_id_cmd, filters=GROUP_FILTER))
    app.add_handler(CommandHandler("esik", set_threshold, filters=GROUP_FILTER))
    app.add_handler(
        MessageHandler(filters.Document.ALL & GROUP_FILTER, handle_document)
    )
    app.add_handler(
        MessageHandler(
            filters.TEXT & ~filters.COMMAND & GROUP_FILTER,
            handle_non_document,
        )
    )

    logger.info(
        "Bot başlatılıyor | izinli chat_id sayısı=%s | ids=%s",
        len(ALLOWED_CHAT_IDS),
        sorted(ALLOWED_CHAT_IDS),
    )
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
