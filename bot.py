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

# -----------------------------
# LOGGING
# -----------------------------

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Убираем подробные httpx-логи, чтобы Telegram-токен не светился в Railway Logs
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)

# -----------------------------
# ENV
# -----------------------------

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")

# Можно менять в Railway через переменную ANTHROPIC_MODEL
ANTHROPIC_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-6")

MODEL_CANDIDATES = []

if ANTHROPIC_MODEL:
    MODEL_CANDIDATES.append(ANTHROPIC_MODEL)

for model_name in ["claude-sonnet-4-6", "claude-haiku-4-5"]:
    if model_name not in MODEL_CANDIDATES:
        MODEL_CANDIDATES.append(model_name)

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

# -----------------------------
# PROMPTS
# -----------------------------

SYSTEM_PROMPT_EXTRACT = """Ты помощник для анализа списков деталей мебели для программы раскроя GibLab.

ТВОЯ ЗАДАЧА В ЭТОМ ЗАПРОСЕ:
Распознать ТОЛЬКО размеры и количество деталей.

НЕ анализируй кромку.
НЕ объясняй.
НЕ пиши markdown.
НЕ пиши ```json.
Верни ТОЛЬКО JSON.

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

2. Если число написано с точкой или запятой, убирай точку/запятую:
39.8 → 398
39,8 → 398
200.8 → 2008
150.6 → 1506
77.5 → 775
98.7 → 987
61.8 → 618
60.8 → 608
37.6 → 376

3. Если на рукописном фото короткие размеры написаны как 77 x 75, переводи в миллиметры:
77 → 770
75 → 750
39 → 390
98 → 980
37 → 370
84 → 840
65 → 650
60 → 600

4. Если это уже печатная таблица с размерами 2255, 550, 150, 303, 70, 133, 320, 350, 577, ничего не умножай.

ВАЖНО:
- Не анализируй кромку.
- Не придумывай строки, которых нет.
- Сохраняй порядок строк сверху вниз.
- Если строка плохо читается, всё равно постарайся распознать.
- Верни только валидный JSON.

Формат ответа строго такой:
{"parts":[{"length":2255,"width":550,"qty":2}]}
"""

SYSTEM_PROMPT_EDGES = """Ты помощник для анализа кромки деталей мебели для программы раскроя GibLab.

ТВОЯ ЗАДАЧА В ЭТОМ ЗАПРОСЕ:
Определить ТОЛЬКО кромку по уже известным строкам.

НЕ объясняй.
НЕ пиши markdown.
НЕ пиши ```json.
Верни ТОЛЬКО JSON.

ВАЖНО:
На изображении могут быть:
1) рукописные строки
2) печатная таблица с сеткой

Тебя интересуют ТОЛЬКО короткие подчёркивания непосредственно под числами.

НЕ СЧИТАЙ КРОМКОЙ:
- обычные линии таблицы;
- длинные горизонтальные линии от одной границы ячейки до другой;
- нижние границы ячеек;
- сетку таблицы;
- линии между строками.

СЧИТАЙ КРОМКОЙ:
- короткую отдельную линию прямо под числом;
- линию, которая заметно короче ширины ячейки;
- линию, которая визуально относится именно к числу, а не к таблице.

СТОЛБЦЫ:
- первое число в строке = length
- второе число в строке = width

ПОЛЯ КРОМКИ:
- e = ОВ = кромка сверху по длине
- f = ОН = кромка снизу по длине
- g = ОЛ = кромка слева по ширине
- h = ОП = кромка справа по ширине

ПРАВИЛА:
1. Если подчёркнуто первое число одной линией:
   e="Кромка", f="", g="", h=""

2. Если подчёркнуто первое число двумя линиями:
   e="Кромка", f="Кромка", g="", h=""

3. Если подчёркнуто второе число одной линией:
   e="", f="", g="Кромка", h=""

4. Если подчёркнуто второе число двумя линиями:
   e="", f="", g="Кромка", h="Кромка"

5. Если подчёркнуты оба числа:
   заполни и length, и width по количеству линий.

6. Если кромки нет:
   e="", f="", g="", h=""

КРИТИЧНО:
- В печатной таблице линия сетки под числом НЕ является кромкой.
- Если сомневаешься между "линия таблицы" и "подчёркивание", выбирай "линия таблицы".
- Кромку ставь только когда отдельное подчёркивание действительно видно.
- Для каждой входной строки обязательно верни один объект.

Ты получишь список строк в правильном порядке.
Для каждой строки верни один объект с полями:
row, e, f, g, h

Формат ответа строго такой:
{"rows":[{"row":1,"e":"","f":"","g":"Кромка","h":""}]}
"""

SYSTEM_PROMPT_REPAIR_JSON = """Ты исправляешь ответ в валидный JSON.
Верни ТОЛЬКО JSON.
Без пояснений.
Без markdown.
Без ```json.
Нельзя добавлять новые данные, которых нет в исходном ответе.
"""

# -----------------------------
# STORAGE
# -----------------------------

user_photos = {}

# -----------------------------
# HELPERS
# -----------------------------

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


def image_to_content_block(image_bytes: bytes, mime_type: str):
    b64 = base64.standard_b64encode(image_bytes).decode("utf-8")

    return {
        "type": "image",
        "source": {
            "type": "base64",
            "media_type": mime_type,
            "data": b64,
        },
    }


def clean_json_text(raw: str) -> str:
    if raw is None:
        return ""

    text = raw.strip()

    text = text.replace("```json", "").replace("```JSON", "").replace("```", "").strip()

    start = text.find("{")
    end = text.rfind("}")

    if start != -1 and end != -1 and end > start:
        text = text[start:end + 1]

    return text.strip()


def parse_json_safely(raw: str):
    cleaned = clean_json_text(raw)

    if not cleaned:
        raise json.JSONDecodeError("Empty JSON", raw or "", 0)

    return json.loads(cleaned)


async def repair_json_with_claude(raw: str):
    content = [
        {
            "type": "text",
            "text": (
                "Исправь этот ответ в валидный JSON. "
                "Верни только JSON, без пояснений.\n\n"
                f"Исходный ответ:\n{raw}"
            ),
        }
    ]

    last_error = None

    for model_name in MODEL_CANDIDATES:
        try:
            logger.info("Trying JSON repair with model: %s", model_name)

            response = client.messages.create(
                model=model_name,
                max_tokens=4000,
                system=SYSTEM_PROMPT_REPAIR_JSON,
                messages=[
                    {
                        "role": "user",
                        "content": content,
                    }
                ],
            )

            repaired_raw = response.content[0].text.strip()
            logger.info("JSON repair raw answer: %s", repaired_raw)

            return parse_json_safely(repaired_raw)

        except Exception as e:
            last_error = e
            logger.error("JSON repair failed with %s | Error: %s", model_name, str(e))
            continue

    raise last_error


async def call_claude_json(system_prompt: str, content: list):
    last_error = None

    for model_name in MODEL_CANDIDATES:
        try:
            logger.info("Trying Claude model: %s", model_name)

            response = client.messages.create(
                model=model_name,
                max_tokens=4000,
                system=system_prompt,
                messages=[
                    {
                        "role": "user",
                        "content": content,
                    }
                ],
            )

            raw = response.content[0].text.strip()

            logger.info("Claude model used: %s", model_name)
            logger.info("Claude raw answer: %s", raw)

            try:
                return parse_json_safely(raw)
            except Exception as json_error:
                logger.error("First JSON parse failed: %s", str(json_error))
                return await repair_json_with_claude(raw)

        except Exception as e:
            last_error = e
            error_text = str(e)

            logger.error("Claude model failed: %s | Error: %s", model_name, error_text)

            if "not_found_error" in error_text or "404" in error_text or "model:" in error_text:
                continue

            raise e

    raise last_error


async def save_user_image(update: Update, image_bytes: bytes, mime_type: str):
    user_id = update.effective_user.id

    if user_id not in user_photos:
        user_photos[user_id] = []

    user_photos[user_id].append(
        {
            "bytes": image_bytes,
            "mime_type": mime_type,
        }
    )

    count = len(user_photos[user_id])

    await update.message.reply_text(
        f"📸 Фото {count} добавлено.\nМожешь добавить ещё или нажми кнопку:",
        reply_markup=get_keyboard(count),
    )


async def extract_parts_from_image(image_bytes: bytes, mime_type: str):
    content = [
        image_to_content_block(image_bytes, mime_type),
        {
            "type": "text",
            "text": (
                "Распознай все строки деталей на этом изображении. "
                "Верни только размеры и количество. "
                "Не анализируй кромку. "
                "Сохрани порядок строк сверху вниз. "
                "Ответ строго JSON."
            ),
        },
    ]

    data = await call_claude_json(SYSTEM_PROMPT_EXTRACT, content)
    return data.get("parts", [])


async def extract_edges_from_image(image_bytes: bytes, mime_type: str, parts: list):
    rows_for_prompt = []

    for idx, p in enumerate(parts, start=1):
        rows_for_prompt.append(
            {
                "row": idx,
                "length": p.get("length", ""),
                "width": p.get("width", ""),
                "qty": p.get("qty", ""),
            }
        )

    content = [
        image_to_content_block(image_bytes, mime_type),
        {
            "type": "text",
            "text": (
                "Ниже уже распознанные строки таблицы/списка. "
                "Для каждой строки определи только кромку. "
                "Не путай нижнюю границу ячейки таблицы с подчёркиванием. "
                "Если под числом только обычная линия сетки таблицы, кромку не ставь. "
                "Если под числом есть короткое отдельное подчёркивание, ставь кромку. "
                "Ответ строго JSON.\n\n"
                f"Строки:\n{json.dumps(rows_for_prompt, ensure_ascii=False)}"
            ),
        },
    ]

    data = await call_claude_json(SYSTEM_PROMPT_EDGES, content)
    return data.get("rows", [])


def merge_parts_and_edges(parts: list, edge_rows: list):
    edges_map = {}

    for row in edge_rows:
        row_num = row.get("row")

        try:
            row_num = int(row_num)
        except Exception:
            continue

        edges_map[row_num] = row

    merged = []

    for idx, part in enumerate(parts, start=1):
        edge = edges_map.get(idx, {})

        merged.append(
            {
                "length": part.get("length", ""),
                "width": part.get("width", ""),
                "qty": part.get("qty", 1),
                "e": edge.get("e", ""),
                "f": edge.get("f", ""),
                "g": edge.get("g", ""),
                "h": edge.get("h", ""),
            }
        )

    return merged


def build_excel(parts: list):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "GibLab"

    for p in parts:
        ws.append(
            [
                p.get("length", ""),
                p.get("width", ""),
                p.get("qty", ""),
                "",
                p.get("e", ""),
                p.get("f", ""),
                p.get("g", ""),
                p.get("h", ""),
            ]
        )

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output


# -----------------------------
# HANDLERS
# -----------------------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 *Привет! Я бот для раскроя GibLab.*\n\n"
        "Как работать:\n"
        "1️⃣ Отправь фото, скриншот или изображение таблицы\n"
        "2️⃣ Лучше отправлять как *Файл*, если таблица цифровая\n"
        "3️⃣ Если листов несколько — отправляй по одному\n"
        "4️⃣ Нажми кнопку *«Создать XLSX»*\n"
        "5️⃣ Получи файл готовый для GibLab\n\n"
        "📸 Отправляй первое фото или файл!",
        parse_mode="Markdown",
    )


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)

    photo_bytes = BytesIO()
    await file.download_to_memory(photo_bytes)
    photo_bytes.seek(0)

    await save_user_image(update, photo_bytes.read(), "image/jpeg")


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
            "Отправь изображение как файл."
        )
        return

    file = await context.bot.get_file(document.file_id)

    photo_bytes = BytesIO()
    await file.download_to_memory(photo_bytes)
    photo_bytes.seek(0)

    await save_user_image(update, photo_bytes.read(), mime_type)


async def process_photos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = update.effective_user.id

    if user_id not in user_photos or not user_photos[user_id]:
        await query.edit_message_text("❌ Нет фото. Сначала отправь фото или файл.")
        return

    images = user_photos[user_id]

    await query.edit_message_text(f"⏳ Анализирую {len(images)} фото, подожди...")

    try:
        all_parts = []

        for image_index, image_item in enumerate(images, start=1):
            image_bytes = image_item["bytes"]
            mime_type = image_item["mime_type"]

            logger.info("Processing image %s", image_index)

            parts = await extract_parts_from_image(image_bytes, mime_type)
            logger.info("Extracted parts from image %s: %s", image_index, parts)

            if not parts:
                logger.warning("No parts found on image %s", image_index)
                continue

            edge_rows = await extract_edges_from_image(image_bytes, mime_type, parts)
            logger.info("Extracted edges from image %s: %s", image_index, edge_rows)

            merged = merge_parts_and_edges(parts, edge_rows)
            all_parts.extend(merged)

        if not all_parts:
            await context.bot.send_message(
                chat_id=update.effective_chat.id,
                text="❌ Не удалось распознать детали. Попробуй отправить изображение чётче или как файл.",
            )
            user_photos[user_id] = []
            return

        output = build_excel(all_parts)

        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=output,
            filename="giblab_import.xlsx",
            caption=f"✅ Готово! {len(all_parts)} деталей → GibLab",
        )

        user_photos[user_id] = []

    except Exception as e:
        logger.error("Error: %s", e)

        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text=(
                "❌ Ошибка анализа.\n"
                "Попробуй отправить фото чётче или как файл.\n\n"
                f"Технически: {str(e)}"
            ),
        )

        user_photos[user_id] = []


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_photos[user_id] = []

    await update.message.reply_text("🗑 Фото очищены. Начни заново — отправь фото или файл.")


# -----------------------------
# MAIN
# -----------------------------

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
