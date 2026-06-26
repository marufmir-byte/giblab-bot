import os
import json
import base64
import logging
from io import BytesIO

import anthropic
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")

# Модель Claude. Можно поменять в Railway через переменную ANTHROPIC_MODEL.
# Если переменной нет, бот использует эту модель.
ANTHROPIC_MODEL = os.environ.get(
    "ANTHROPIC_MODEL",
    "claude-3-5-sonnet-20241022"
)

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

SYSTEM_PROMPT = """Ты помощник для анализа рукописных списков деталей мебели для программы раскроя GibLab.
Анализируй фото и возвращай ТОЛЬКО JSON без пояснений, без markdown, без ```json.

ПРАВИЛА ПРЕОБРАЗОВАНИЯ РАЗМЕРОВ:
- Если число написано без точки и похоже на мебельный размер, оставляй как есть.
- Если число написано с точкой, убирай точку:
39.8 → 398
200.8 → 2008
150.6 → 1506
77.5 → 775
98.7 → 987

ВАЖНО:
- 160 означает 160, если на фото явно написано 160.
- 1600 означает 1600, если на фото явно написано 1600.
- Не придумывай лишний ноль без причины.

ПРАВИЛА КРОМКИ — определяй ТОЛЬКО по подчёркиваниям под числами:
- Нет подчёркиваний → e="", f="", g="", h=""
- Одна черта под длиной → e="Кромка", остальные ""
- Две черты под длиной → e="Кромка", f="Кромка", g="", h=""
- Одна черта под шириной → g="Кромка", остальные ""
- Две черты под шириной → g="Кромка", h="Кромка", e="", f=""
- Двойное подчёркивание под обоими → e="Кромка", f="Кромка", g="Кромка", h="Кромка"

Если рядом с размером написано количество, например:
77 x 75 = 2
то qty=2.

Если количество не указано, ставь qty=1.

Возвращай ТОЛЬКО JSON:
{"parts":[{"length":770,"width":750,"qty":2,"e":"Кромка","f":"","g":"","h":""}]}"""

user_photos = {}


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 *Привет! Я бот для раскроя GibLab.*\n\n"
        "Как работать:\n"
        "1️⃣ Отправь фото рукописного списка деталей\n"
        "2️⃣ Если листов несколько — отправляй по одному\n"
        "3️⃣ Нажми кнопку *«Создать XLSX»*\n"
        "4️⃣ Получи файл готовый для GibLab\n\n"
        "📸 Отправляй первое фото!",
        parse_mode="Markdown"
    )


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id

    if user_id not in user_photos:
        user_photos[user_id] = []

    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)

    photo_bytes = BytesIO()
    await file.download_to_memory(photo_bytes)
    photo_bytes.seek(0)

    user_photos[user_id].append(photo_bytes.read())

    count = len(user_photos[user_id])

    keyboard = [[InlineKeyboardButton(
        f"✅ Создать XLSX ({count} фото)",
        callback_data="process"
    )]]

    await update.message.reply_text(
        f"📸 Фото {count} добавлено.\nМожешь добавить ещё или нажми кнопку:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )


async def process_photos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = update.effective_user.id

    if user_id not in user_photos or not user_photos[user_id]:
        await query.edit_message_text("❌ Нет фото. Сначала отправь фото.")
        return

    photos = user_photos[user_id]

    await query.edit_message_text(f"⏳ Анализирую {len(photos)} фото, подожди...")

    try:
        content = []

        for photo_data in photos:
            b64 = base64.standard_b64encode(photo_data).decode("utf-8")

            content.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/jpeg",
                    "data": b64
                }
            })

        content.append({
            "type": "text",
            "text": "Проанализируй все фото и верни JSON со ВСЕМИ деталями из всех фото одним списком."
        })

        response = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=4000,
            system=SYSTEM_PROMPT,
            messages=[
                {
                    "role": "user",
                    "content": content
                }
            ]
        )

        raw = response.content[0].text.strip()

        # Чистим случайный markdown, потому что даже модели иногда ведут себя как стажёр после обеда.
        if "```" in raw:
            raw = raw.replace("```json", "").replace("```", "").strip()

        data = json.loads(raw)
        parts = data.get("parts", [])

        if not parts:
            await context.bot.send_message(
                chat_id=update.effective_chat.id,
                text="❌ Не нашёл деталей на фото. Попробуй отправить фото чётче."
            )
            user_photos[user_id] = []
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "GibLab"

        for p in parts:
            ws.append([
                p.get("length", ""),
                p.get("width", ""),
                p.get("qty", ""),
                "",
                p.get("e", ""),
                p.get("f", ""),
                p.get("g", ""),
                p.get("h", ""),
            ])

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=output,
            filename="giblab_import.xlsx",
            caption=f"✅ Готово! {len(parts)} деталей → GibLab"
        )

        user_photos[user_id] = []

    except json.JSONDecodeError:
        logger.error("JSON decode error. Raw answer: %s", raw)

        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text="❌ Claude ответил не JSON. Попробуй ещё раз или пришли более чёткое фото."
        )

        user_photos[user_id] = []

    except Exception as e:
        logger.error("Error: %s", e)

        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text=f"❌ Ошибка: {str(e)}\nПопробуй снова."
        )

        user_photos[user_id] = []


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_photos[user_id] = []

    await update.message.reply_text("🗑 Фото очищены. Начни заново — отправь фото.")


def main():
    if not TELEGRAM_BOT_TOKEN:
        raise ValueError("TELEGRAM_BOT_TOKEN не найден в Railway Variables.")

    if not ANTHROPIC_API_KEY:
        raise ValueError("ANTHROPIC_API_KEY не найден в Railway Variables.")

    app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("cancel", cancel))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(CallbackQueryHandler(process_photos, pattern="process"))

    logger.info("Bot started with Claude model: %s", ANTHROPIC_MODEL)

    app.run_polling()


if __name__ == "__main__":
    main()
