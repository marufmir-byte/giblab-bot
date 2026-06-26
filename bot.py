import os
import json
import base64
import logging
from io import BytesIO

import anthropic
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)

# Основные логи
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Убираем подробные httpx-логи, чтобы Telegram-токен не светился в Railway Logs
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")

# Модель можно менять в Railway через переменную ANTHROPIC_MODEL
ANTHROPIC_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-haiku-4-5")

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

SYSTEM_PROMPT = """Ты помощник для анализа списков деталей мебели для программы раскроя GibLab.
Анализируй фото, скриншот таблицы или рукописный список и возвращай ТОЛЬКО JSON.
Без пояснений, без markdown, без ```json.

ЗАДАЧА:
Нужно распознать размеры мебельных деталей и подготовить JSON для Excel/GibLab.

ФОРМАТ СТРОКИ:
Строка может выглядеть так:
77 x 75 = 2
77 x 39 = 1
98 x 75 = 2

Или как таблица:
2255 550 2
2031 150 1
1705 550 2

Где:
- первое число = length
- второе число = width
- третье число или число после "=" = qty
- если количество не указано, qty=1

ПРАВИЛА РАЗМЕРОВ:
1. Если число уже написано в миллиметрах, оставляй как есть:
2255 → 2255
2031 → 2031
1705 → 1705
550 → 550
150 → 150
303 → 303
133 → 133
70 → 70

2. Если число написано с точкой, убирай точку:
39.8 → 398
200.8 → 2008
150.6 → 1506
77.5 → 775
98.7 → 987

3. Если на рукописном фото короткие размеры написаны как 77 x 75, переводи в миллиметры:
77 → 770
75 → 750
39 → 390
98 → 980
37 → 370
84 → 840
65 → 650
60 → 600

4. Если это уже таблица с размерами 2255, 550, 150, 303, 70, 133, 320, 350, ничего не умножай.

ПРАВИЛА КРОМКИ — САМОЕ ВАЖНОЕ:
На фото или скриншоте под некоторыми числами есть чёрные горизонтальные подчёркивания.
Эти подчёркивания означают кромку.

Игнорируй:
- линии таблицы;
- вертикальные линии;
- рамки ячеек;
- границы строк;
- линии между строками.

Учитывай только короткую чёрную линию прямо под самим числом.

Поля кромки:
- e и f относятся к первому размеру length
- g и h относятся ко второму размеру width

Правила:
1. Если подчёркнут первый размер одной линией:
   e="Кромка", f="", g="", h=""

2. Если подчёркнут первый размер двумя линиями:
   e="Кромка", f="Кромка", g="", h=""

3. Если подчёркнут второй размер одной линией:
   e="", f="", g="Кромка", h=""

4. Если подчёркнут второй размер двумя линиями:
   e="", f="", g="Кромка", h="Кромка"

5. Если подчёркнуты оба размера:
   ставь кромку и для length, и для width по количеству линий.

6. Если подчёркивание видно даже слабо, но оно находится прямо под числом, считай это кромкой.

7. Если кромка есть, обязательно заполни соответствующие поля словом "Кромка".

ПРИМЕРЫ:
Если строка:
2255 подчёркнуто, 550, qty 2
то:
{"length":2255,"width":550,"qty":2,"e":"Кромка","f":"","g":"","h":""}

Если строка:
850 подчёркнуто, 350 подчёркнуто, qty 6
то:
{"length":850,"width":350,"qty":6,"e":"Кромка","f":"","g":"Кромка","h":""}

Если строка:
626 без подчёркивания, 577 подчёркнуто, qty 3
то:
{"length":626,"width":577,"qty":3,"e":"","f":"","g":"Кромка","h":""}

ВАЖНО:
- Не придумывай детали, которых нет на фото.
- Не пропускай подчёркивания.
- Если строка плохо читается, всё равно постарайся распознать.
- Для каждой детали обязательно верни поля: length, width, qty, e, f, g, h.
- Возвращай только валидный JSON.
- Никакого текста до JSON и после JSON.
- Никакого markdown.
- Никаких ```json.

Возвращай строго в таком формате:
{"parts":[{"length":2255,"width":550,"qty":2,"e":"Кромка","f":"","g":"","h":""}]}
"""

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
        parse_mode="Markdown",
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

    keyboard = [
        [
            InlineKeyboardButton(
                f"✅ Создать XLSX ({count} фото)",
                callback_data="process",
            )
        ]
    ]

    await update.message.reply_text(
        f"📸 Фото {count} добавлено.\nМожешь добавить ещё или нажми кнопку:",
        reply_markup=InlineKeyboardMarkup(keyboard),
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

    raw = ""

    try:
        content = []

        for photo_data in photos:
            b64 = base64.standard_b64encode(photo_data).decode("utf-8")

            content.append(
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/jpeg",
                        "data": b64,
                    },
                }
            )

        content.append(
            {
                "type": "text",
                "text": (
                    "Проанализируй все фото. "
                    "Особое внимание удели коротким чёрным подчёркиваниям прямо под числами: они означают кромку. "
                    "Верни JSON со ВСЕМИ деталями из всех фото одним списком. "
                    "Для каждой строки обязательно заполни length, width, qty, e, f, g, h."
                ),
            }
        )

        response = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=4000,
            system=SYSTEM_PROMPT,
            messages=[
                {
                    "role": "user",
                    "content": content,
                }
            ],
        )

        raw = response.content[0].text.strip()

        # Сырой ответ Claude в Railway Logs.
        # Здесь можно проверить, увидел ли Claude кромку.
        logger.info("Claude raw answer: %s", raw)

        if "```json" in raw:
            raw = raw.replace("```json", "").replace("```", "").strip()
        elif "```" in raw:
            raw = raw.replace("```", "").strip()

        data = json.loads(raw)
        parts = data.get("parts", [])

        if not parts:
            await context.bot.send_message(
                chat_id=update.effective_chat.id,
                text="❌ Не нашёл деталей на фото. Попробуй отправить фото чётче.",
            )
            user_photos[user_id] = []
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "GibLab"

        for p in parts:
            length = p.get("length", "")
            width = p.get("width", "")
            qty = p.get("qty", "")

            e = p.get("e", "")
            f = p.get("f", "")
            g = p.get("g", "")
            h = p.get("h", "")

            ws.append(
                [
                    length,
                    width,
                    qty,
                    "",
                    e,
                    f,
                    g,
                    h,
                ]
            )

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=output,
            filename="giblab_import.xlsx",
            caption=f"✅ Готово! {len(parts)} деталей → GibLab",
        )

        user_photos[user_id] = []

    except json.JSONDecodeError:
        logger.error("JSON decode error. Raw answer: %s", raw)

        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text="❌ Claude ответил не JSON. Попробуй ещё раз или пришли более чёткое фото.",
        )

        user_photos[user_id] = []

    except Exception as e:
        logger.error("Error: %s", e)

        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text=f"❌ Ошибка: {str(e)}\nПопробуй снова.",
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

    logger.info("Bot started with Claude model: %s", ANTHROPIC_MODEL)

    app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("cancel", cancel))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(CallbackQueryHandler(process_photos, pattern="process"))

    app.run_polling()


if __name__ == "__main__":
    main()
