# server.py - X Backend v11.0 (исправлены баги заполнения таблицы)
# Установка: pip install aiohttp telethon openpyxl
# Запуск: python server.py
# Порт: 8765

import asyncio, json, os, re, time, traceback
from datetime import datetime
from aiohttp import web
import aiohttp
from telethon import TelegramClient
from telethon.errors import SessionPasswordNeededError, FloodWaitError
from telethon.sessions import StringSession
from telethon.tl.types import DocumentAttributeFilename
from openpyxl import load_workbook, Workbook

API_ID = 2985935
API_HASH = "a436d51ced3ec96a65d8414eb8e0a92d"
sessions = {}
user_clients = {}
TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)
pending_confirms = {}
active_probevs = set()  # защита от двойного запуска

# ====================== MIDDLEWARE ======================
@web.middleware
async def log_and_cors(request, handler):
    """Логирует ВСЕ запросы + CORS."""
    ts = datetime.now().strftime('%H:%M:%S')
    print(f"[REQ] {ts}  {request.method} {request.path}  from={request.remote}", flush=True)
    if request.method == "OPTIONS":
        response = web.Response(status=204)
    else:
        try:
            response = await handler(request)
        except Exception as e:
            print(f"[REQ] {ts}  ERROR {request.method} {request.path}: {e}", flush=True)
            traceback.print_exc()
            response = web.json_response({"ok": False, "error": str(e)}, status=500)
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    print(f"[REQ] {ts}  -> {response.status}", flush=True)
    return response

# ====================== УТИЛИТЫ ======================
def normalize_fio(raw):
    """Приводит ФИО к единому формату: ИВАНОВ ИВАН ИВАНОВИЧ"""
    if not raw:
        return ""
    words = [re.sub(r'[^A-Za-zА-ЯЁа-яё\-]', '', w) for w in str(raw).strip().split() if w]
    # Фильтруем пустые строки (могут появиться после очистки от цифр/спецсимволов)
    words = [w for w in words if w]
    if not words:
        return ""
    return ' '.join([w[0].upper() + w[1:].lower() for w in words]).upper().replace('Ё', 'Е')


# Словарь исправлений опечаток в отчествах
PATRONYMIC_FIXES = {
    'ВЛАДИМИРОЕВИЧ': 'ВЛАДИМИРОВИЧ', 'ВЛАДИМИРОЕВНА': 'ВЛАДИМИРОВНА',
    'АЛЕКСАНДРОЕВИЧ': 'АЛЕКСАНДРОВИЧ', 'АЛЕКСАНДРОЕВНА': 'АЛЕКСАНДРОВНА',
    'АЛЕКСЕЕВИЧ': 'АЛЕКСЕЕВИЧ',  # правильно
    'НИКОЛАЕВИЧ': 'НИКОЛАЕВИЧ',  # правильно
    'СЕРГЕЕВИЧ': 'СЕРГЕЕВИЧ',    # правильно
    'МИХАЙЛОЕВИЧ': 'МИХАЙЛОВИЧ', 'МИХАЙЛОЕВНА': 'МИХАЙЛОВНА',
    'АНДРЕЕВИЧ': 'АНДРЕЕВИЧ',    # правильно
    'ДМИТРИЕВИЧ': 'ДМИТРИЕВИЧ',  # правильно
    'ВАЛЕРЬЕВИЧ': 'ВАЛЕРЬЕВИЧ',  # правильно
    'АЛЕКСЕЕВНА': 'АЛЕКСЕЕВНА',  # правильно
    'АНДРЕЕВНА': 'АНДРЕЕВНА',    # правильно
    'СЕРГЕЕВНА': 'СЕРГЕЕВНА',    # правильно
}

# Явно мусорные значения (placeholder'ы, а не имена)
GARBAGE_NAMES = {'NULL', 'NONE', 'NAN', 'XXX', 'ХХХ', '-', '--', '...', '•••',
                 'НЕТ', 'НЕТУ', 'НЕИЗВЕСТНО', 'НЕИЗВЕСТЕН', 'НЕИЗВЕСТНА', 'UNKNOWN'}

# Слова-маркеры для удаления (оглы/кызы и др.)
REMOVE_MARKERS = {'ОГЛЫ', 'ОГЛИ', 'КЫЗЫ', 'КИЗИ', 'УГЛИ', 'УЛЫ', 'ГЫЗЫ'}


def validate_and_clean_fio(raw):
    """
    Полная нормализация и валидация ФИО.
    Возвращает (нормализованное_ФИО, причина_удаления).
    Если ФИО валидно — причина_удаления = None.
    """
    if not raw or not str(raw).strip():
        return "", "пустое ФИО"

    # Нормализуем
    fio = normalize_fio(raw)
    if not fio:
        return "", "пустое после нормализации"

    parts = fio.split()
    # Не требуем 2+ слов — «Анна» или «Анна Васильевна» тоже валидны

    # Проверка на мусор
    for part in parts:
        if part in GARBAGE_NAMES:
            return "", f"мусорное имя: {fio}"
        # Одна буква — подозрительно (но 'Я' или 'И' могут быть реальными именами)
        if len(part) == 1 and part not in ('Я', 'И', 'А', 'У', 'О', 'Е'):
            return "", f"одна буква в ФИО: {fio}"

    # Проверка на оглы/кызы — только это удаляем безусловно
    for part in parts:
        if part in REMOVE_MARKERS:
            return "", f"содержит {part}: {fio}"

    # Исправление перемешанных букв ФИО: ИФО/ФОИ/ОИФ/ИОФ → ФИО
    fixed_parts = []
    for part in parts:
        if set(part) == {'Ф', 'И', 'О'} and len(part) == 3:
            fixed_parts.append('ФИО')
        elif set(part) == {'Ф', 'И'} and len(part) == 2:
            fixed_parts.append('ФИ')
        else:
            fixed_parts.append(part)

    # Исправление опечаток в отчествах
    final_parts = []
    for part in fixed_parts:
        if part in PATRONYMIC_FIXES:
            final_parts.append(PATRONYMIC_FIXES[part])
        # Общий паттерн: ОЕВИЧ → ОВИЧ, ОЕВНА → ОВНА
        elif 'ОЕВИЧ' in part:
            final_parts.append(part.replace('ОЕВИЧ', 'ОВИЧ'))
        elif 'ОЕВНА' in part:
            final_parts.append(part.replace('ОЕВНА', 'ОВНА'))
        else:
            final_parts.append(part)

    fio = ' '.join(final_parts)
    return fio, None


def clean_phone(text):
    """Очищает телефон до формата +7XXXXXXXXXX (10 цифр после +7)"""
    if not text:
        return ""
    d = re.sub(r'[^0-9]', '', str(text))
    if len(d) >= 10:
        return '+7' + d[-10:]
    return d


def clean_snils(text):
    """Оставляет только цифры СНИЛС (до 11 знаков)"""
    if not text:
        return ""
    d = re.sub(r'[^0-9]', '', str(text))
    return d[:11] if len(d) >= 11 else d


def parse_date(val):
    """Парсит дату в формат ДД.ММ.ГГГГ"""
    if not val:
        return ""
    s = str(val).strip()
    # Excel может вернуть datetime-объект
    if hasattr(val, 'strftime'):
        return val.strftime('%d.%m.%Y')
    m = re.search(r'(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})', s)
    if m:
        d, mo, y = m.group(1), m.group(2), m.group(3)
        if len(y) == 2:
            y = '20' + y if int(y) < 30 else '19' + y
        return f"{d.zfill(2)}.{mo.zfill(2)}.{y}"
    return ""


def phones_match(p1, p2):
    """Сравнивает два телефона по последним 10 цифрам. Возвращает True если совпадают."""
    d1 = re.sub(r'[^0-9]', '', str(p1 or ''))
    d2 = re.sub(r'[^0-9]', '', str(p2 or ''))
    if not d1 or not d2:
        return False
    # Сравниваем последние 10 цифр
    return d1[-10:] == d2[-10:]


# ====================== TELEGRAM КЛИЕНТ ======================
async def get_client(ss):
    if ss in user_clients:
        c = user_clients[ss]
        if not c.is_connected():
            await c.connect()
        if await c.is_user_authorized():
            return c
    c = TelegramClient(StringSession(ss), API_ID, API_HASH)
    await c.connect()
    if await c.is_user_authorized():
        user_clients[ss] = c
        return c
    await c.disconnect()
    raise Exception("Session invalid")


# ====================== ПОДТВЕРЖДЕНИЯ ЧЕРЕЗ БОТА ======================
async def send_confirm(bot_token, chat_id, stage_name, count, confirm_id):
    text = f"\u26a0\ufe0f ПОДТВЕРДИТЕ ПРОБИВ\n\nЭтап: {stage_name}\nСтрок: {count}"
    kb = {
        "inline_keyboard": [[
            {"text": "\u2705 ПОДТВЕРДИТЬ", "callback_data": f"confirm_{confirm_id}"},
            {"text": "\u274c ОТМЕНА", "callback_data": f"cancel_{confirm_id}"}
        ]]
    }
    try:
        async with aiohttp.ClientSession() as s:
            await s.post(
                f"https://api.telegram.org/bot{bot_token}/sendMessage",
                json={"chat_id": int(chat_id), "text": text, "reply_markup": kb}
            )
        asyncio.create_task(poll_updates(bot_token, chat_id, confirm_id))
        return True
    except Exception as e:
        print(f"[CONFIRM] Ошибка отправки подтверждения: {e}")
        return False


async def send_file_to_bot(bot_token, chat_id, filepath, caption=""):
    """Отправляет файл в бота-уведомитель"""
    try:
        async with aiohttp.ClientSession() as s:
            data = aiohttp.FormData()
            data.add_field('chat_id', str(chat_id))
            data.add_field('caption', caption)
            data.add_field('document', open(filepath, 'rb'))
            await s.post(f"https://api.telegram.org/bot{bot_token}/sendDocument", data=data)
            print(f"[BOT] Файл отправлен: {caption}")
    except Exception as e:
        print(f"[BOT] Ошибка отправки файла: {e}")


async def poll_updates(bot_token, chat_id, confirm_id):
    """Ждет нажатия кнопки подтверждения/отмены в боте"""
    offset = 0
    start = time.time()
    print(f"[POLL] Начинаю опрос для confirm_id={confirm_id}, chat={chat_id}")
    while time.time() - start < 600:
        try:
            async with aiohttp.ClientSession() as s:
                async with s.get(
                    f"https://api.telegram.org/bot{bot_token}/getUpdates",
                    params={"offset": offset, "timeout": 30}
                ) as resp:
                    data = await resp.json()
                    if data.get("ok") and data.get("result"):
                        for upd in data["result"]:
                            offset = upd["update_id"] + 1
                            cb = upd.get("callback_query")
                            if cb and str(cb.get("message", {}).get("chat", {}).get("id")) == str(chat_id):
                                cb_data = cb.get("data", "")
                                await s.post(
                                    f"https://api.telegram.org/bot{bot_token}/answerCallbackQuery",
                                    json={"callback_query_id": cb["id"]}
                                )
                                msg_id = cb.get("message", {}).get("message_id")
                                if cb_data == f"confirm_{confirm_id}":
                                    await s.post(
                                        f"https://api.telegram.org/bot{bot_token}/editMessageText",
                                        json={"chat_id": int(chat_id), "message_id": msg_id,
                                              "text": "\u2705 ПОДТВЕРЖДЕНО! Выполняю..."}
                                    )
                                    pending_confirms[confirm_id] = True
                                    print(f"[POLL] ПОДТВЕРЖДЕНО: {confirm_id}")
                                    return
                                elif cb_data == f"cancel_{confirm_id}":
                                    await s.post(
                                        f"https://api.telegram.org/bot{bot_token}/editMessageText",
                                        json={"chat_id": int(chat_id), "message_id": msg_id,
                                              "text": "\u274c ОТМЕНЕНО"}
                                    )
                                    pending_confirms[confirm_id] = False
                                    print(f"[POLL] ОТМЕНЕНО: {confirm_id}")
                                    return
        except Exception as e:
            print(f"[POLL] Ошибка опроса: {e}")
        await asyncio.sleep(1)
    # Таймаут
    pending_confirms[confirm_id] = False
    print(f"[POLL] ТАЙМАУТ: {confirm_id}")


async def safe_confirm(bot_token, chat_id, stage_name, count, confirm_id, add_log):
    """Безопасное подтверждение: если бот недоступен — авто-продолжаем, не виснем."""
    if not bot_token or not chat_id:
        add_log("[v] Бот не настроен — авто-продолжаю")
        return True

    sent = await send_confirm(bot_token, chat_id, stage_name, count, confirm_id)
    if not sent:
        add_log("[!] Не удалось отправить запрос в бот — авто-продолжаю")
        return True

    add_log(f"[ОЖИДАНИЕ] Откройте бот и нажмите «Подтвердить» для: {stage_name}")
    ok = await wait_confirm(confirm_id, timeout=600)
    if ok is None or ok is False:
        add_log(f"[x] {stage_name} — отменён или таймаут")
        return False
    add_log(f"[v] ПОДТВЕРЖДЕНО: {stage_name}")
    return True


async def wait_confirm(confirm_id, timeout=600):
    """Ожидает результат подтверждения"""
    start = time.time()
    while time.time() - start < timeout:
        if confirm_id in pending_confirms:
            r = pending_confirms.pop(confirm_id)
            return r
        await asyncio.sleep(0.5)
    return False


# ====================== РАБОТА С БОТАМИ ПРОБИВА ======================
async def clear_bot(client, bot):
    """Сбрасывает состояние бота командой /start"""
    try:
        e = await client.get_entity(bot)
        await client.send_message(e, "/start")
        await asyncio.sleep(2)
        print(f"[BOT] /start отправлен в {bot}")
    except Exception as ex:
        print(f"[BOT] Ошибка /start для {bot}: {ex}")


async def click_btn(client, bot, text, retries=3):
    """Ищет и нажимает кнопку с указанным текстом. Делает до retries попыток."""
    e = await client.get_entity(bot)
    for attempt in range(retries):
        if attempt > 0:
            await asyncio.sleep(3)  # ждём перед повтором
        try:
            async for msg in client.iter_messages(e, limit=30):
                if msg.buttons and time.time() - msg.date.timestamp() < 300:
                    for row in msg.buttons:
                        for btn in row:
                            if btn.text and text.lower() in btn.text.lower():
                                await btn.click()
                                await asyncio.sleep(2)
                                print(f"[BOT] Нажата кнопка '{btn.text}' в {bot}")
                                return True
            # Логируем кнопки для отладки
            async for msg in client.iter_messages(e, limit=5):
                if msg.buttons:
                    btns = [b.text for row in msg.buttons for b in row if b.text]
                    print(f"[BOT] Доступные кнопки в {bot}: {btns}")
                    break
            print(f"[BOT] Кнопка '{text}' не найдена в {bot} (попытка {attempt+1}/{retries})")
        except Exception as ex:
            print(f"[BOT] Ошибка поиска кнопки: {ex}")
    return False


async def wait_xlsx(client, bot, timeout=180, since_msg_id=None):
    """Ждет XLSX-файл от бота. since_msg_id — искать только сообщения с ID > этого."""
    e = await client.get_entity(bot)
    start = time.time()
    print(f"[BOT] Ожидаю XLSX от {bot} (таймаут {timeout}с, since_msg_id={since_msg_id})...")
    while time.time() - start < timeout:
        msgs = await client.get_messages(e, limit=5)
        for msg in msgs:
            if not msg or not msg.document:
                continue
            if since_msg_id is not None and msg.id <= since_msg_id:
                continue  # старое сообщение — пропускаем
            for a in msg.document.attributes:
                if isinstance(a, DocumentAttributeFilename) and a.file_name.endswith('.xlsx'):
                    print(f"[BOT] Получен XLSX: {a.file_name} (msg_id={msg.id})")
                    return msg
        await asyncio.sleep(3)
    print(f"[BOT] XLSX не получен от {bot} за {timeout}с")
    return None


# ====================== ПАРСИНГ XLSX ОТ БОТОВ ======================
def parse_xlsx(path):
    """
    Парсит XLSX-ответ от бота.
    Ищет колонки по заголовкам: ФИО, Дата, Телефон, СНИЛС.
    Возвращает список dict с ключами fio, date, phone, snils.
    """
    res = []
    try:
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        h = {}  # ключ -> номер колонки

        # Ищем заголовки в первой строке (пропускаем ИНН/Паспорт)
        for col in range(1, ws.max_column + 1):
            v = str(ws.cell(row=1, column=col).value or "").upper().strip()
            print(f"[PARSE] Колонка {col}: '{v}'")
            # Пропускаем колонки ИНН и Паспорт
            if any(k in v for k in ['ИНН', 'INN', 'ПАСПОРТ', 'PASSPORT']):
                continue
            if any(k in v for k in ['FIO', 'ФИО', 'ИМЯ', 'ФАМИЛИЯ', 'NAME']):
                h['fio'] = col
            if any(k in v for k in ['ДАТА', 'BIRTH', 'РОЖД', 'DATE']):
                h['date'] = col
            # НОМЕР — только если это НЕ паспорт и НЕ ИНН
            if any(k in v for k in ['ТЕЛЕФОН', 'PHONE', 'ТЕЛ']):
                h['phone'] = col
            elif 'НОМЕР' in v and 'ПАСПОРТ' not in v and 'ИНН' not in v:
                h['phone'] = col
            if any(k in v for k in ['СНИЛС', 'SNILS']):
                h['snils'] = col

        print(f"[PARSE] Найдены заголовки: {h}")

        # Читаем строки — каждая в своём try/except, одна битая не ломает всё
        for row in range(2, ws.max_row + 1):
            try:
                r = {}
                if 'fio' in h:
                    r['fio'] = normalize_fio(ws.cell(row=row, column=h['fio']).value)
                if 'date' in h:
                    r['date'] = parse_date(ws.cell(row=row, column=h['date']).value)
                if 'phone' in h:
                    r['phone'] = clean_phone(ws.cell(row=row, column=h['phone']).value)
                if 'snils' in h:
                    r['snils'] = clean_snils(ws.cell(row=row, column=h['snils']).value)

                # Добавляем только если есть хоть что-то
                if any(v for v in r.values() if v):
                    res.append(r)
            except Exception as row_err:
                print(f"[PARSE] Ошибка в строке {row}: {row_err}")

        print(f"[PARSE] Всего записей: {len(res)}")
        if res:
            print(f"[PARSE] Пример: {res[0]}")
        return res
    except Exception as e:
        print(f"[PARSE] Ошибка: {e}")
        traceback.print_exc()
        return []


# ====================== ЗАПОЛНЕНИЕ ТАБЛИЦЫ ======================
# СТРОГИЕ КОЛОНКИ в итоговой таблице:
#   A=1: № таблицы
#   B=2: ФИО
#   C=3: Дата
#   D=4: Номер
#   E=5: СНИЛС
#   F=6: Адресс

COL_NO = 1
COL_FIO = 2
COL_DATE = 3
COL_PHONE = 4
COL_SNILS = 5
COL_ADDR = 6


def fill_dates_from_response(ws, response_records):
    """
    Заполняет ДАТЫ (колонка C) в итоговой таблице.
    Ищет совпадение по НОМЕРУ (колонка D).
    """
    filled = 0
    print(f"[FILL-DATES] Начинаю заполнение дат. Ответов: {len(response_records)}")
    for rec in response_records:
        rec_phone = rec.get('phone', '')
        rec_date = rec.get('date', '')
        if not rec_date or not rec_phone:
            print(f"[FILL-DATES]   Пропуск (нет даты или телефона): phone={rec_phone}, date={rec_date}")
            continue
        found = False
        for row in range(2, ws.max_row + 1):
            table_phone = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
            existing_date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()

            if phones_match(table_phone, rec_phone):
                if not existing_date or existing_date == 'None':
                    ws.cell(row=row, column=COL_DATE).value = rec_date
                    filled += 1
                    print(f"[FILL-DATES]   Строка {row}: ЗАПОЛНЕНО! phone={table_phone} -> date={rec_date}")
                else:
                    print(f"[FILL-DATES]   Строка {row}: дата уже есть '{existing_date}', пропуск")
                found = True
                break

        if not found:
            print(f"[FILL-DATES]   НЕ НАЙДЕНО: phone={rec_phone}, date={rec_date}")
            sample = []
            for row in range(2, min(ws.max_row + 1, 7)):
                sample.append(str(ws.cell(row=row, column=COL_PHONE).value or "").strip())
            print(f"[FILL-DATES]   Телефоны в таблице (первые 5): {sample}")

    print(f"[FILL-DATES] ИТОГО заполнено дат: {filled}")
    return filled


def fill_phones_from_response(ws, response_records):
    """
    Заполняет НОМЕРА (колонка D) в итоговой таблице.
    Ищет совпадение по ФИО (колонка B) и ДАТЕ (колонка C).
    """
    filled = 0
    print(f"[FILL-PHONES] Начинаю заполнение номеров. Ответов: {len(response_records)}")
    for rec in response_records:
        rec_fio = normalize_fio(rec.get('fio', ''))
        rec_phone = rec.get('phone', '')
        rec_date = rec.get('date', '')
        if not rec_phone or not rec_fio:
            print(f"[FILL-PHONES]   Пропуск (нет телефона или ФИО): fio={rec_fio}, phone={rec_phone}")
            continue

        found = False
        for row in range(2, ws.max_row + 1):
            table_fio = normalize_fio(str(ws.cell(row=row, column=COL_FIO).value or ""))
            existing_phone = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
            table_date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()

            # Совпадение по ФИО
            if table_fio == rec_fio:
                # Если в ответе есть дата — сверяем и её
                if rec_date and table_date and table_date != rec_date:
                    print(f"[FILL-PHONES]   Строка {row}: ФИО совпало, но даты разные: '{table_date}' vs '{rec_date}', пропуск")
                    continue

                if not existing_phone or existing_phone == 'None':
                    ws.cell(row=row, column=COL_PHONE).value = rec_phone
                    filled += 1
                    print(f"[FILL-PHONES]   Строка {row}: ЗАПОЛНЕНО! fio={table_fio} -> phone={rec_phone}")
                else:
                    print(f"[FILL-PHONES]   Строка {row}: номер уже есть '{existing_phone}', пропуск")
                found = True
                break

        if not found:
            print(f"[FILL-PHONES]   НЕ НАЙДЕНО: fio={rec_fio}, phone={rec_phone}, date={rec_date}")
            sample = []
            for row in range(2, min(ws.max_row + 1, 7)):
                sample.append(normalize_fio(str(ws.cell(row=row, column=COL_FIO).value or "")))
            print(f"[FILL-PHONES]   ФИО в таблице (первые 5): {sample}")

    print(f"[FILL-PHONES] ИТОГО заполнено номеров: {filled}")
    return filled


def fill_snils_dates(ws, response_records):
    """
    Заполняет ДАТЫ (колонка C). Сначала по СНИЛС, если нет — по ФИО.
    """
    filled = 0
    print(f"[FILL-SNILS] Начинаю заполнение дат. Ответов: {len(response_records)}")
    for rec in response_records:
        rec_snils = clean_snils(rec.get('snils', ''))
        rec_fio = normalize_fio(rec.get('fio', ''))
        rec_date = rec.get('date', '')
        if not rec_date:
            continue

        found = False
        # Попытка 1: сопоставление по СНИЛС
        if rec_snils and len(rec_snils) >= 11:
            for row in range(2, ws.max_row + 1):
                table_snils = clean_snils(str(ws.cell(row=row, column=COL_SNILS).value or ""))
                if table_snils == rec_snils:
                    existing = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
                    if not existing or existing == 'None':
                        ws.cell(row=row, column=COL_DATE).value = rec_date
                        filled += 1
                        print(f"[FILL-SNILS]   Строка {row}: СНИЛС={table_snils} -> дата={rec_date}")
                    found = True
                    break

        # Попытка 2: fallback по ФИО (бот2 не возвращает СНИЛС)
        if not found and rec_fio:
            for row in range(2, ws.max_row + 1):
                table_fio = normalize_fio(str(ws.cell(row=row, column=COL_FIO).value or ""))
                if table_fio == rec_fio:
                    existing = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
                    if not existing or existing == 'None':
                        ws.cell(row=row, column=COL_DATE).value = rec_date
                        filled += 1
                        print(f"[FILL-SNILS]   Строка {row}: ФИО={table_fio} -> дата={rec_date}")
                    found = True
                    break

        if not found:
            print(f"[FILL-SNILS]   НЕ НАЙДЕНО: fio={rec_fio}, snils={rec_snils}, date={rec_date}")

    print(f"[FILL-SNILS] ИТОГО заполнено дат: {filled}")
    return filled


# ====================== ПОЛНЫЙ ЦИКЛ ПРОБИВА ======================
async def run_full_cycle(ss, bot1, bot2, bot_token, chat_id,
                         items_no_date, items_no_phone, items_snils,
                         year_range, original_rows):
    """
    Основной цикл пробива (6 этапов).
    items_no_date  — строки где нет даты, но есть ФИО+номер
    items_no_phone — строки где нет номера, но есть ФИО+дата
    items_snils    — список СНИЛС (только для строк с пустой датой!)
    original_rows  — [tableNum, fio, date, phone, snils, addr]
    """
    client = await get_client(ss)
    log = []

    def add(msg):
        ts = datetime.now().strftime('%H:%M:%S')
        log.append(f"[{ts}] {msg}")
        print(f"[LOG] {msg}")

    async def send_status(stage_label):
        """Отправляет текущее состояние таблицы в бот-уведомитель."""
        if bot_token and chat_id:
            wb.save(result_file)
            await send_file_to_bot(bot_token, chat_id, result_file,
                                   f"Таблица после: {stage_label}")

    # === ПРЕД-ФИЛЬТР ПО ГОДАМ (дублирует фронтенд для надёжности) ===
    if year_range:
        try:
            parts = year_range.split('-')
            yf, yt = int(parts[0]), int(parts[1])
            before = len(original_rows)
            filtered = []
            removed_out = 0
            for row in original_rows:
                date_str = str(row[2] if len(row) > 2 else "").strip()  # row[2] = дата (новый формат)
                if not date_str or date_str == 'None' or date_str == '0':
                    filtered.append(row)
                    continue
                parsed = parse_date(date_str)
                if not parsed:
                    filtered.append(row)
                    continue
                try:
                    year = int(parsed.split('.')[2])
                    if yf <= year <= yt:
                        filtered.append(row)
                    else:
                        removed_out += 1
                except ValueError:
                    filtered.append(row)
            original_rows = filtered
            add(f"[ПРЕД-ФИЛЬТР] {yf}-{yt}: удалено {removed_out} строк (дата вне диапазона), оставлено {len(original_rows)}")
        except Exception as e:
            add(f"[ПРЕД-ФИЛЬТР] Ошибка: {e}")

    # === НОРМАЛИЗАЦИЯ ФИО: удаление оглы/кызы, мусора, исправление опечаток ===
    removed_fio = 0
    fixed_typos = 0
    cleaned_rows = []
    for row in original_rows:
        fio_raw = str(row[1] if len(row) > 1 else "")  # row[1] = ФИО
        clean_fio, reason = validate_and_clean_fio(fio_raw)
        if reason:
            removed_fio += 1
            add(f"[ФИО] Удалена строка №{row[0]}: {reason}")
            continue
        if clean_fio != normalize_fio(fio_raw):
            fixed_typos += 1
            # Обновляем ФИО в строке
            row_list = list(row)
            row_list[1] = clean_fio
            cleaned_rows.append(tuple(row_list))
        else:
            cleaned_rows.append(row)
    original_rows = cleaned_rows
    add(f"[ФИО] Нормализация: удалено {removed_fio} строк (оглы/кызы/мусор), исправлено опечаток: {fixed_typos}")

    # === СОЗДАЕМ ИТОГОВУЮ ТАБЛИЦУ ===
    result_file = os.path.join(TEMP_DIR, f"result_{int(time.time())}.xlsx")
    wb = Workbook()
    ws = wb.active
    # СТРОГИЙ порядок колонок:
    ws.append(["N таблицы", "ФИО", "Дата", "Номер", "СНИЛС", "Адресс"])

    for i, row in enumerate(original_rows):
        ws.append([
            str(row[0] if len(row) > 0 else "").strip(),  # tableNum — ОРИГИНАЛЬНЫЙ, не трогаем!
            str(row[1] if len(row) > 1 else "").strip(),  # ФИО
            str(row[2] if len(row) > 2 else "").strip(),  # Дата
            str(row[3] if len(row) > 3 else "").strip(),  # Номер
            str(row[4] if len(row) > 4 else "").strip(),  # СНИЛС
            str(row[5] if len(row) > 5 else "").strip()   # Адресс
        ])
    wb.save(result_file)
    add(f"Таблица создана: {ws.max_row - 1} строк")
    await send_status("создание таблицы")

    # === АНАЛИЗИРУЕМ ЧТО РЕАЛЬНО НУЖНО ПРОБИТЬ ===
    # Фильтруем СНИЛС: только для строк где ПУСТАЯ дата в колонке C
    real_snils = []
    for row in range(2, ws.max_row + 1):
        existing_date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        snils_val = str(ws.cell(row=row, column=COL_SNILS).value or "").strip()
        if (not existing_date or existing_date == 'None') and snils_val and len(clean_snils(snils_val)) >= 11:
            real_snils.append(clean_snils(snils_val))

    # Убираем дубликаты
    real_snils = list(set(real_snils))

    add(f"К пробиву: дат={len(items_no_date)} номеров={len(items_no_phone)} снилс={len(real_snils)}")

    # ============ ЭТАП 1: ФИО+НОМЕР -> БОТ1 (заполняем даты) ============
    if items_no_date:
        cid = f"s1_{int(time.time())}"
        add(f"ЭТАП 1: ФИО+НОМЕР -> бот1 ({len(items_no_date)} строк)")

        if bot_token and chat_id:
            ok = await safe_confirm(bot_token, chat_id, "ЭТАП 1: ФИО+НОМЕР (пустые даты)", len(items_no_date), cid, add)
            if ok is False:
                add("[x] ЭТАП 1 ОТМЕНЁН! Отправляю текущий файл...")
                wb.save(result_file)
                await send_file_to_bot(bot_token, chat_id, result_file, "Файл после отмены этапа 1")
                return {"ok": True, "log": log, "cancelled": True}

        add("[v] ПОДТВЕРЖДЕНО! Начинаю этап 1...")

        # Формируем TXT: ФИО\tНОМЕР
        txt = "\n".join([f"{it.get('fio','')}\t{it.get('phone','')}" for it in items_no_date])
        tpath = os.path.join(TEMP_DIR, f"t1_{int(time.time())}.txt")
        with open(tpath, 'w', encoding='utf-8') as f:
            f.write(txt)
        add(f"TXT создан: {tpath}")

        await clear_bot(client, bot1)
        e = await client.get_entity(bot1)
        await client.send_message(e, "Пробивы")
        await asyncio.sleep(2)
        await click_btn(client, bot1, "ФИО+номер")
        await asyncio.sleep(2)
        # Запоминаем ID последнего сообщения перед отправкой
        last_msgs = await client.get_messages(e, limit=1)
        last_msg_id = last_msgs[0].id if last_msgs else 0

        await client.send_file(e, tpath)
        add("Файл отправлен в бот1, жду ответ...")

        msg = await wait_xlsx(client, bot1, 300, since_msg_id=last_msg_id)
        if msg:
            rpath = os.path.join(TEMP_DIR, f"r1_{int(time.time())}.xlsx")
            await client.download_media(msg, file=rpath)
            recs = parse_xlsx(rpath)
            add(f"Получено ответов от бот1: {len(recs)}")
            fill_dates_from_response(ws, recs)
            wb.save(result_file)
            await send_status("этап 1 (ФИО+номер → бот1)")
        else:
            add("[!] Бот1 не ответил на этапе 1")
    else:
        add("ЭТАП 1 пропущен (нет строк без даты)")

    # === ПЕРЕСЧЁТ СНИЛС после этапа 1 — часть дат уже заполнена! ===
    real_snils = []
    for row in range(2, ws.max_row + 1):
        existing_date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        snils_val = str(ws.cell(row=row, column=COL_SNILS).value or "").strip()
        if (not existing_date or existing_date == 'None') and snils_val and len(clean_snils(snils_val)) >= 11:
            real_snils.append(clean_snils(snils_val))
    real_snils = list(set(real_snils))
    add(f"[ПЕРЕСЧЁТ] СНИЛС с пустой датой после этапа 1: {len(real_snils)}")

    # ============ ЭТАП 2: СНИЛС -> БОТ1 ============
    if real_snils:
        cid = f"s2_{int(time.time())}"
        add(f"ЭТАП 2: СНИЛС -> бот1 ({len(real_snils)} снилс)")

        if bot_token and chat_id:
            ok = await safe_confirm(bot_token, chat_id, "ЭТАП 2: СНИЛС (бот1)", len(real_snils), cid, add)
            if ok is False:
                add("[x] ЭТАП 2 ОТМЕНЁН! Отправляю текущий файл...")
                wb.save(result_file)
                await send_file_to_bot(bot_token, chat_id, result_file, "Файл после отмены этапа 2")
                return {"ok": True, "log": log, "cancelled": True}

        add("[v] ПОДТВЕРЖДЕНО! Начинаю этап 2...")

        txt = "\n".join(real_snils)
        tpath = os.path.join(TEMP_DIR, f"t2_{int(time.time())}.txt")
        with open(tpath, 'w', encoding='utf-8') as f:
            f.write(txt)

        await clear_bot(client, bot1)
        e = await client.get_entity(bot1)
        await client.send_message(e, "Пробивы")
        await asyncio.sleep(4)  # даём боту время показать меню
        await click_btn(client, bot1, "СНИЛС")
        await asyncio.sleep(3)
        # Запоминаем ID последнего сообщения перед отправкой
        last_msgs = await client.get_messages(e, limit=1)
        last_msg_id = last_msgs[0].id if last_msgs else 0

        await client.send_file(e, tpath)
        add("СНИЛС отправлены в бот1, жду ответ...")

        msg = await wait_xlsx(client, bot1, 300, since_msg_id=last_msg_id)
        if msg:
            rpath = os.path.join(TEMP_DIR, f"r2_{int(time.time())}.xlsx")
            await client.download_media(msg, file=rpath)
            recs = parse_xlsx(rpath)
            add(f"Получено через СНИЛС бот1: {len(recs)}")
            fill_snils_dates(ws, recs)
            wb.save(result_file)
            await send_status("этап 2 (СНИЛС → бот1)")
        else:
            add("[!] Бот1 не ответил на этапе 2")
    else:
        add("ЭТАП 2 пропущен (нет СНИЛС с пустой датой)")

    # ============ ЭТАП 3: СНИЛС -> БОТ2 (только оставшиеся с пустой датой) ============
    # ПЕРЕПРОВЕРЯЕМ после этапа 2 — часть дат уже могла заполниться
    snils_still_empty = []
    for row in range(2, ws.max_row + 1):
        existing_date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        snils_val = str(ws.cell(row=row, column=COL_SNILS).value or "").strip()
        if (not existing_date or existing_date == 'None') and snils_val and len(clean_snils(snils_val)) >= 11:
            snils_still_empty.append(clean_snils(snils_val))
    snils_still_empty = list(set(snils_still_empty))

    if snils_still_empty:
        cid = f"s3_{int(time.time())}"
        add(f"ЭТАП 3: СНИЛС -> бот2 ({len(snils_still_empty)} снилс, после этапа 2)")

        if bot_token and chat_id:
            ok = await safe_confirm(bot_token, chat_id, "ЭТАП 3: СНИЛС (бот2)", len(snils_still_empty), cid, add)
            if ok is False:
                add("[x] ЭТАП 3 ОТМЕНЁН! Отправляю текущий файл...")
                wb.save(result_file)
                await send_file_to_bot(bot_token, chat_id, result_file, "Файл после отмены этапа 3")
                return {"ok": True, "log": log, "cancelled": True}

        add("[v] ПОДТВЕРЖДЕНО! Начинаю этап 3...")

        txt = "\n".join(snils_still_empty)
        tpath = os.path.join(TEMP_DIR, f"t3_{int(time.time())}.txt")
        with open(tpath, 'w', encoding='utf-8') as f:
            f.write(txt)

        e = await client.get_entity(bot2)
        # Запоминаем ID последнего сообщения перед отправкой
        last_msgs = await client.get_messages(e, limit=1)
        last_msg_id = last_msgs[0].id if last_msgs else 0

        await client.send_file(e, tpath)
        add("СНИЛС отправлены в бот2, жду ответ...")

        msg = await wait_xlsx(client, bot2, 300, since_msg_id=last_msg_id)
        if msg:
            rpath = os.path.join(TEMP_DIR, f"r3_{int(time.time())}.xlsx")
            await client.download_media(msg, file=rpath)
            recs = parse_xlsx(rpath)
            add(f"Получено через СНИЛС бот2: {len(recs)}")
            fill_snils_dates(ws, recs)
            wb.save(result_file)
            await send_status("этап 3 (СНИЛС → бот2)")
        else:
            add("[!] Бот2 не ответил на этапе 3")
    else:
        add("ЭТАП 3 пропущен (все СНИЛС уже с датами после этапа 2)")

    # ============ ЭТАП 4: ФИЛЬТР ГОДОВ ============
    if year_range:
        try:
            parts = year_range.split('-')
            yf, yt = int(parts[0]), int(parts[1])
            add(f"ЭТАП 4: Фильтр годов {yf}-{yt}")

            rows_to_delete = []
            for row in range(2, ws.max_row + 1):
                date_val = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
                parts_d = date_val.split('.')
                if len(parts_d) == 3:
                    try:
                        year = int(parts_d[2])
                        if year < yf or year > yt:
                            rows_to_delete.append(row)
                            add(f"  Строка {row}: год {year} вне диапазона — удалена")
                    except ValueError:
                        rows_to_delete.append(row)
                        add(f"  Строка {row}: некорректная дата '{date_val}' — удалена")
                else:
                    rows_to_delete.append(row)
                    add(f"  Строка {row}: нет даты — удалена")

            for row in reversed(rows_to_delete):
                ws.delete_rows(row)
            wb.save(result_file)
            await send_status(f"этап 4 (фильтр {yf}-{yt})")
            add(f"Удалено строк: {len(rows_to_delete)}, осталось: {ws.max_row - 1}")
        except Exception as e:
            add(f"Ошибка фильтра годов: {e}")
    else:
        add("ЭТАП 4 пропущен (не задан диапазон годов)")

    # ============ ЭТАП 5: ФИО+ДАТА -> БОТ1 (заполняем номера) ============
    # ПЕРЕСЧИТЫВАЕМ строки без номера — таблица изменилась после этапов 1-4!
    items_no_phone_fresh = []
    for row in range(2, ws.max_row + 1):
        fio_val = str(ws.cell(row=row, column=COL_FIO).value or "").strip()
        date_val = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        phone_val = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
        if (fio_val and fio_val != 'None'
                and date_val and date_val != 'None' and date_val != '0'
                and (not phone_val or phone_val == 'None' or phone_val == '0')):
            items_no_phone_fresh.append({
                'fio': normalize_fio(fio_val),
                'date': date_val
            })

    if items_no_phone_fresh:
        cid = f"s5_{int(time.time())}"
        add(f"ЭТАП 5: ФИО+ДАТА -> бот1 ({len(items_no_phone_fresh)} строк, было {len(items_no_phone)} до обработки)")

        if bot_token and chat_id:
            ok = await safe_confirm(bot_token, chat_id, "ЭТАП 5: ФИО+ДАТА (пустые номера)", len(items_no_phone_fresh), cid, add)
            if ok is False:
                add("[x] ЭТАП 5 ОТМЕНЁН! Отправляю текущий файл...")
                wb.save(result_file)
                await send_file_to_bot(bot_token, chat_id, result_file, "Файл после отмены этапа 5")
                return {"ok": True, "log": log, "cancelled": True}

        add("[v] ПОДТВЕРЖДЕНО! Начинаю этап 5...")

        txt = "\n".join([f"{it.get('fio','')}\t{it.get('date','')}" for it in items_no_phone_fresh])
        tpath = os.path.join(TEMP_DIR, f"t5_{int(time.time())}.txt")
        with open(tpath, 'w', encoding='utf-8') as f:
            f.write(txt)

        await clear_bot(client, bot1)
        e = await client.get_entity(bot1)
        await client.send_message(e, "Пробивы")
        await asyncio.sleep(2)
        await click_btn(client, bot1, "ФИО+дата")
        await asyncio.sleep(2)
        # Запоминаем ID последнего сообщения перед отправкой
        last_msgs = await client.get_messages(e, limit=1)
        last_msg_id = last_msgs[0].id if last_msgs else 0

        await client.send_file(e, tpath)
        add("Файл отправлен в бот1, жду ответ...")

        msg = await wait_xlsx(client, bot1, 300, since_msg_id=last_msg_id)
        if msg:
            rpath = os.path.join(TEMP_DIR, f"r5_{int(time.time())}.xlsx")
            await client.download_media(msg, file=rpath)
            recs = parse_xlsx(rpath)
            add(f"Получено ответов: {len(recs)}")
            fill_phones_from_response(ws, recs)
            wb.save(result_file)
            await send_status("этап 5 (ФИО+дата → бот1)")
        else:
            add("[!] Бот1 не ответил на этапе 5")
    else:
        add("ЭТАП 5 пропущен (нет строк без номера)")

    # ============ ЭТАП 6: ДОБИВ -> БОТ2 ============
    # ПЕРЕСЧИТЫВАЕМ после этапа 5 — часть номеров уже заполнена!
    items_no_phone_final = []
    for row in range(2, ws.max_row + 1):
        fio_val = str(ws.cell(row=row, column=COL_FIO).value or "").strip()
        date_val = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        phone_val = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
        if (fio_val and fio_val != 'None'
                and date_val and date_val != 'None' and date_val != '0'
                and (not phone_val or phone_val == 'None' or phone_val == '0')):
            items_no_phone_final.append({
                'fio': normalize_fio(fio_val),
                'date': date_val
            })

    if items_no_phone_final:
        cid = f"s6_{int(time.time())}"
        add(f"ЭТАП 6: ДОБИВ -> бот2 ({len(items_no_phone_final)} строк, было {len(items_no_phone_fresh)} до этапа 5)")

        if bot_token and chat_id:
            ok = await safe_confirm(bot_token, chat_id, "ЭТАП 6: ДОБИВ (бот2)", len(items_no_phone_final), cid, add)
            if ok is False:
                add("[x] ЭТАП 6 ОТМЕНЁН! Отправляю текущий файл...")
                wb.save(result_file)
                await send_file_to_bot(bot_token, chat_id, result_file, "Файл после отмены этапа 6")
                return {"ok": True, "log": log, "cancelled": True}

        add("[v] ПОДТВЕРЖДЕНО! Добиваю номера через бот2...")

        e = await client.get_entity(bot2)
        for i, it in enumerate(items_no_phone_final):
            try:
                fio = it.get('fio', '')
                date = it.get('date', '')
                await client.send_message(e, f"{fio} {date}")
                await asyncio.sleep(5)

                # Ищем ответ с телефоном
                async for msg in client.iter_messages(e, limit=5):
                    if msg.text and time.time() - msg.date.timestamp() < 60:
                        phones = re.findall(r'(?:\+?79\d{9})', msg.text)
                        if phones:
                            p = clean_phone(phones[0])
                            add(f"  Найден телефон: {p} для {fio}")

                            # Заполняем в таблице
                            for row in range(2, ws.max_row + 1):
                                table_fio = normalize_fio(str(ws.cell(row=row, column=COL_FIO).value or ""))
                                table_date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
                                existing_phone = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()

                                if table_fio == normalize_fio(fio) and table_date == date:
                                    if not existing_phone or existing_phone == 'None':
                                        ws.cell(row=row, column=COL_PHONE).value = p
                                        add(f"  Добив: {fio} -> {p} (строка {row})")
                                    break
                            break

                if (i + 1) % 10 == 0:
                    wb.save(result_file)
                    add(f"  Сохранено после {i+1} строк")
                    await asyncio.sleep(3)

            except FloodWaitError as e:
                add(f"  FloodWait: {e.seconds}с")
                await asyncio.sleep(e.seconds)
            except Exception as e:
                add(f"  Ошибка добива для {fio}: {e}")

        wb.save(result_file)
        await send_status("этап 6 (добив → бот2)")
        add("Добив завершён")
    else:
        add("ЭТАП 6 пропущен (нет строк без номера)")

    # === ФИНАЛ ===
    add("=== ВСЕ ЭТАПЫ ЗАВЕРШЕНЫ ===")
    # Подсчитаем статистику
    total_dates = 0
    total_phones = 0
    for row in range(2, ws.max_row + 1):
        if str(ws.cell(row=row, column=COL_DATE).value or "").strip():
            total_dates += 1
        if str(ws.cell(row=row, column=COL_PHONE).value or "").strip():
            total_phones += 1
    add(f"Итого: строк={ws.max_row-1}, с датами={total_dates}, с номерами={total_phones}")

    await send_status("ФИНАЛ (все этапы)")

    return {"ok": True, "log": log}


# ====================== ЭНДПОИНТЫ ======================
async def handle_health(request):
    return web.json_response({"ok": True, "message": "X Backend v11.0"})


async def handle_root(request):
    """Корневой эндпоинт — всегда отдаёт HTML (health-check всё равно нужен только 200)."""
    try:
        html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "сайт.html")
        with open(html_path, "r", encoding="utf-8") as f:
            html = f.read()
        # aiohttp сам добавляет charset при text= — нельзя дублировать в content_type
        return web.Response(text=html, content_type="text/html")
    except Exception as e:
        print(f"[ROOT] Ошибка чтения HTML: {e}", flush=True)
        return web.Response(text="OK", content_type="text/plain")


async def handle_ping(request):
    """Проверка живости — максимально быстрый ответ."""
    return web.Response(text="pong", content_type="text/plain")


async def handle_favicon(request):
    """Отдаём пустой favicon, чтобы браузер не спамил 500."""
    return web.Response(status=204)  # No Content


async def handle_send_code(request):
    """Отправка кода подтверждения Telegram"""
    try:
        d = await request.json()
        phone = d.get("phone", "").strip()
        if not phone:
            return web.json_response({"ok": False, "error": "Введите номер"}, status=400)

        c = TelegramClient(StringSession(), API_ID, API_HASH)
        await c.connect()
        r = await c.send_code_request(phone)
        sessions[phone] = {"client": c, "hash": r.phone_code_hash}
        print(f"[AUTH] Код отправлен на {phone}")
        return web.json_response({"ok": True})
    except Exception as e:
        print(f"[AUTH] Ошибка send_code: {e}")
        return web.json_response({"ok": False, "error": str(e)}, status=400)


async def handle_verify_code(request):
    """Проверка кода + 2FA"""
    try:
        d = await request.json()
        phone = d.get("phone", "").strip()
        code = d.get("code", "").strip()
        password = d.get("password", "").strip()

        if phone not in sessions:
            return web.json_response({"ok": False, "error": "Сначала отправьте код"}, status=400)

        s = sessions[phone]
        c = s["client"]

        try:
            await c.sign_in(phone, code)
        except SessionPasswordNeededError:
            if not password:
                return web.json_response({"ok": True, "need_2fa": True})
            await c.sign_in(password=password)

        ss = c.session.save()
        me = await c.get_me()
        del sessions[phone]
        user_clients[ss] = c
        print(f"[AUTH] Успешный вход: {me.first_name} ({phone})")
        return web.json_response({
            "ok": True,
            "session": ss,
            "phone": phone,
            "name": me.first_name or ""
        })
    except Exception as e:
        print(f"[AUTH] Ошибка verify_code: {e}")
        return web.json_response({"ok": False, "error": str(e)}, status=400)


async def handle_full_probev(request):
    """Запуск полного цикла пробива"""
    try:
        d = await request.json()
        ss = d.get("session", "")
        bot1 = d.get("bot1", "@osint_pam_pam_bot")
        bot2 = d.get("bot2", "@proverim123_bot")
        bot_token = d.get("bot_token", "")
        chat_id = d.get("chat_id", "")
        year_range = d.get("year_range", "") or "1945-1975"
        items_no_date = d.get("items_no_date", [])
        items_no_phone = d.get("items_no_phone", [])
        items_snils = d.get("items_snils", [])
        original_rows = d.get("original_rows", [])

        if not ss:
            return web.json_response({"ok": False, "error": "Нет сессии"}, status=400)

        # Защита от двойного запуска
        if ss in active_probevs:
            return web.json_response({"ok": False, "error": "Пробив уже выполняется для этой сессии"}, status=409)
        active_probevs.add(ss)

        try:
            print(f"\n{'='*60}")
            print(f"[PROBEV] ЗАПУСК ПОЛНОГО ЦИКЛА")
            print(f"[PROBEV] Бот1: {bot1}, Бот2: {bot2}")
            print(f"[PROBEV] Строк без даты: {len(items_no_date)}")
            print(f"[PROBEV] Строк без номера: {len(items_no_phone)}")
            print(f"[PROBEV] СНИЛС: {len(items_snils)}")
            print(f"[PROBEV] Всего строк: {len(original_rows)}")
            print(f"{'='*60}\n")

            result = await run_full_cycle(
                ss, bot1, bot2, bot_token, chat_id,
                items_no_date, items_no_phone, items_snils,
                year_range, original_rows
            )
            return web.json_response(result)
        finally:
            active_probevs.discard(ss)
    except Exception as e:
        traceback.print_exc()
        return web.json_response({"ok": False, "error": str(e)}, status=400)


# ====================== ЗАПУСК ======================
# client_max_size=50MB — чтобы фронтенд мог слать большие таблицы (129+ строк)
app = web.Application(middlewares=[log_and_cors], client_max_size=50 * 1024 * 1024)
app.router.add_get("/", handle_root)
app.router.add_get("/health", handle_health)
app.router.add_get("/ping", handle_ping)
app.router.add_get("/favicon.ico", handle_favicon)
app.router.add_post("/send-code", handle_send_code)
app.router.add_post("/verify-code", handle_verify_code)
app.router.add_post("/full-probev", handle_full_probev)


async def on_startup(app):
    """Вызывается при старте приложения."""
    port = app["port"]
    msg = (
        "=" * 60 + "\n"
        f"X Backend v11.0 ЗАПУЩЕН\n"
        f"Host: 0.0.0.0  |  Port: {port}\n"
        f"Environment PORT: {os.environ.get('PORT', 'не задан')}\n"
        + "=" * 60
    )
    print(msg, flush=True)


async def on_shutdown(app):
    """Вызывается при остановке (SIGTERM от Railway)."""
    print("[SHUTDOWN] Получен сигнал остановки, закрываю соединения...", flush=True)
    for ss, client in list(user_clients.items()):
        try:
            if client.is_connected():
                await client.disconnect()
                print(f"[SHUTDOWN] Клиент отключён: {ss[:10]}...", flush=True)
        except Exception:
            pass
    user_clients.clear()
    sessions.clear()
    pending_confirms.clear()
    print("[SHUTDOWN] Завершено.", flush=True)


if __name__ == "__main__":
    import sys, signal as _signal
    print("[START] Инициализация...", flush=True)
    try:
        port = int(os.environ.get("PORT", 4545))
        app["port"] = port

        app.on_startup.append(on_startup)
        app.on_shutdown.append(on_shutdown)

        print(f"[START] Запуск на 0.0.0.0:{port}...", flush=True)

        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

        runner = web.AppRunner(app)
        loop.run_until_complete(runner.setup())

        # Вручную создаём TCP-сокет — полный контроль над binding
        import socket as _socket
        sock = _socket.socket(_socket.AF_INET, _socket.SOCK_STREAM)
        sock.setsockopt(_socket.SOL_SOCKET, _socket.SO_REUSEADDR, 1)
        # SO_REUSEPORT — может помочь на Railway
        try:
            sock.setsockopt(_socket.SOL_SOCKET, _socket.SO_REUSEPORT, 1)
        except (AttributeError, OSError):
            pass
        sock.bind(("0.0.0.0", port))
        sock.listen(128)
        sock.setblocking(False)

        site = web.SockSite(runner, sock)
        loop.run_until_complete(site.start())

        print(f"[START] СЕРВЕР ГОТОВ — порт {port} открыт", flush=True)
        print(f"[START] Ожидаю запросы...", flush=True)

        # Явная обработка SIGTERM (Railway) и SIGINT
        def shutdown():
            print("[SIGNAL] Получен сигнал, останавливаю...", flush=True)
            loop.create_task(_do_shutdown(runner, sock, loop))

        async def _do_shutdown(runner, sock, loop):
            await runner.cleanup()
            sock.close()
            loop.stop()

        for sig in (_signal.SIGTERM, _signal.SIGINT):
            try:
                loop.add_signal_handler(sig, shutdown)
            except NotImplementedError:
                pass  # Windows

        loop.run_forever()
        loop.close()
        print("[START] Сервер остановлен.", flush=True)

    except Exception as e:
        print(f"[FATAL] Ошибка запуска: {e}", flush=True)
        traceback.print_exc()
        sys.stderr.flush()
        sys.exit(1)