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

# Основная модель.
# Если в Railway есть ANTHROPIC_MODEL, бот сначала попробует её.
# Если модели нет или будет 404, бот попробует запасные варианты.
ANTHROPIC_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-6")

MODEL_CANDIDATES = []
if ANTHROPIC_MODEL:
    MODEL_CANDIDATES.append(ANTHROPIC_MODEL)

# Запасные модели. Если первая не работает, бот попробует следующую.
for model_name in ["claude-sonnet-4-6", "claude-haiku-4-5"]:
    if model_name not in MODEL_CANDIDATES:
        MODEL_CANDIDATES.append(model_name)

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

SYSTEM_PROMPT = """Ты помощник для анализа списков деталей мебели для программы раскроя GibLab.
Анализируй фото, скриншот таблицы или рукописный список и возвращай ТОЛЬКО JSON.
Без пояснений, без markdown, без ```json.

ЗАДАЧА:
Нужно распознать размеры мебельных деталей и подготовить JSON для Excel/GibLab.

СТРУКТУРА GIBLAB:
В Excel должны быть такие колонки:

A = Д-на = length
B = Ш-на = width
C = Кол = qty
D = Ткт
E = ОВ = оклейка кромки сверху, по длине
F = ОН = оклейка кромки снизу, по длине
G = ОЛ = оклейка кромки слева, по ширине
H = ОП = оклейка кромки справа, по ширине

В JSON используй поля:
length, width, qty, e, f, g, h

Соответствие:
e = ОВ
f = ОН
g = ОЛ
h = ОП

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
320 → 320
350 → 350
577 → 577

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

4. Если это уже таблица с размерами 2255, 550, 150, 303, 70, 133, 320, 350, 577, ничего не умножай.

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

КРИТИЧЕСКОЕ ПРАВИЛО:
Если подчёркнуто число в ПЕРВОМ столбце, то это кромка по длине:
e="Кромка"

Если подчёркнуто число во ВТОРОМ столбце, то это кромка по ширине:
g="Кромка"

Если под числом две отдельные линии:
для первого столбца e="Кромка", f="Кромка"
для второго столбца g="Кромка", h="Кромка"

ОБЫЧНЫЕ ПРАВИЛА:
1. Если подчёркнут первый размер length одной линией:
   e="Кромка", f="", g="", h=""

2. Если подчёркнут первый размер length двумя линиями:
   e="Кромка", f="Кромка", g="", h=""

3. Если подчёркнут второй размер width одной линией:
   e="", f="", g="Кромка", h=""

4. Если подчёркнут второй размер width двумя линиями:
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
1134, 303 подчёркнуто, qty 1
то:
{"length":1134,"width":303,"qty":1,"e":"","f":"","g":"Кромка","h":""}

Если строка:
850 подчёркнуто, 350 подчёркнуто, qty 6
то:
{"length":850,"width":350,"qty":6,"e":"Кромка","f":"","g":"Кромка","h":""}

Если строка:
626 подчёркнуто, 577, qty 3
то:
{"length":626,"width":577,"qty":3,"e":"Кромка","f":"","g":"","h":""}

ВАЖНО:
- Не придумывай детали, которых нет на фото.
- Не пропускай подчёркивания.
- Не путай подчёркивание числа с линией таблицы.
- Если короткая линия прямо под цифрами, это кромка.
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
        "1️⃣ Отправь фото или скриншот списка деталей\n"
        "2️⃣ Лучше отправлять как *Файл*, а не как обычное фото — так качество выше\n"
        "3️⃣ Если листов несколько — отправляй по одному\n"
        "4️⃣ Нажми кнопку *«Создать XLSX»*\n"
        "5️⃣ Получи файл готовый для GibLab\n\n"
        "📸 Отправляй первое фото или файл!",
        parse_mode="Markdown",
    )


def get_keyboard(count: int):
    return InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(
                    f"✅ Создать XLSX ({count} фото)",
                    callback_data="process",
                )
            ]
        ]
    )


async def save_user_image(update: Update, context: ContextTypes.DEFAULT_TYPE, image_bytes: bytes):
    user_id = update.effective_user.id

    if user_id not in user_photos:
        user_photos[user_id] = []

    user_photos[user_id].append(image_bytes)

    count = len(user_photos[user_id])

    await update.message.reply_text(
        f"📸 Фото {count} добавлено.\nМожешь добавить ещё или нажми кнопку:",
        reply_markup=get_keyboard(count),
    )


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)

    photo_bytes = BytesIO()
    await file.download_to_memory(photo_bytes)
    photo_bytes.seek(0)

    await save_user_image(update, context, photo_bytes.read())


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document

    if not document:
        await update.message.reply_text("❌ Файл не найден.")
        return

    mime_type = document.mime_type or ""

    allowed_mime_types = [
        "image/jpeg",
        "image/png",
        "image/webp",
    ]

    if mime_type not in allowed_mime_types:
        await update.message.reply_text(
            "❌ Сейчас я принимаю только изображения: JPG, PNG или WEBP.\n"
            "Отправь скриншот или фото как файл."
        )
        return

    file = await context.bot.get_file(document.file_id)

    photo_bytes = BytesIO()
    await file.download_to_memory(photo_bytes)
    photo_bytes.seek(0)

    await save_user_image(update, context, photo_bytes.read())


async def call_claude_with_fallback(content):
    last_error = None

    for model_name in MODEL_CANDIDATES:
        try:
            logger.info("Trying Claude model: %s", model_name)

            response = client.messages.create(
                model=model_name,
                max_tokens=4000,
                system=SYSTEM_PROMPT,
                messages=[
                    {
                        "role": "user",
                        "content": content,
                    }
                ],
            )

            logger.info("Claude model used: %s", model_name)
            return response

        except Exception as e:
            last_error = e
            error_text = str(e)

            logger.error("Claude model failed: %s | Error: %s", model_name, error_text)

            if "not_found_error" in error_text or "404" in error_text or "model:" in error_text:
                continue

            raise e

    raise last_error


async def process_photos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = update.effective_user.id

    if user_id not in user_photos or not user_photos[user_id]:
        await query.edit_message_text("❌ Нет фото. Сначала отправь фото или файл.")
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
                    "Если подчёркнут первый столбец, заполни e='Кромка'. "
                    "Если подчёркнут второй столбец, заполни g='Кромка'. "
                    "Верни JSON со ВСЕМИ деталями из всех фото одним списком. "
                    "Для каждой строки обязательно заполни length, width, qty, e, f, g, h."
                ),
            }
        )

        response = await call_claude_with_fallback(content)

        raw = response.content[0].text.strip()

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
                text="❌ Не нашёл деталей на фото. Попробуй отправить фото чётче или как файл.",
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
            text="❌ Claude ответил не JSON. Попробуй ещё раз или отправь фото как файл.",
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

    await update.message.reply_text("🗑 Фото очищены. Начни заново — отправь фото или файл.")


def main():
    if not TELEGRAM_BOT_TOKEN:
        raise ValueError("TELEGRAM_BOT_TOKEN не найден в Railway Variables.")

    if not ANTHROPIC_API_KEY:
        raise ValueError("ANTHROPIC_API_KEY не найден в Railway Variables.")

    logger.info("Bot started. Model candidates: %s", MODEL_CANDIDATES)

    app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("cancel", cancel))

    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.Document.IMAGE, handle_document))

    app.add_handler(CallbackQueryHandler(process_photos, pattern="process"))

    app.run_polling()


if __name__ == "__main__":
    main()
