# server.py - X Backend v12.0 (FULL INTEGRATION: DeepSeek + 7 этапов + ZIP + управление)
# Установка: pip install aiohttp telethon openpyxl
# Запуск: python server.py
# Порт: 8765

import asyncio, json, os, re, time, traceback, io, zipfile
from datetime import datetime
from aiohttp import web
import aiohttp
from telethon import TelegramClient
from telethon.errors import SessionPasswordNeededError, FloodWaitError
from telethon.sessions import StringSession
from telethon.tl.types import DocumentAttributeFilename
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill

API_ID = 2985935
API_HASH = "a436d51ced3ec96a65d8414eb8e0a92d"
DEEPSEEK_API_KEY = "sk-ca6b6569a9b64d0a908eb16ec3b69ce5"

sessions = {}
user_clients = {}
TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)
pending_confirms = {}
active_probevs = set()
stop_requested = False  # Глобальный флаг остановки

# ====================== MIDDLEWARE ======================
@web.middleware
async def log_and_cors(request, handler):
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
    if not raw:
        return ""
    words = [re.sub(r'[^A-Za-zА-ЯЁа-яё\-]', '', w) for w in str(raw).strip().split() if w]
    words = [w for w in words if w]
    if not words:
        return ""
    return ' '.join([w[0].upper() + w[1:].lower() for w in words]).upper().replace('Ё', 'Е')


async def normalize_fio_deepseek(fio_text):
    """Нормализация ФИО через DeepSeek API"""
    if not fio_text or len(str(fio_text).strip()) < 2:
        return normalize_fio(fio_text)
    
    try:
        async with aiohttp.ClientSession() as session:
            url = "https://api.deepseek.com/v1/chat/completions"
            headers = {
                "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
                "Content-Type": "application/json"
            }
            payload = {
                "model": "deepseek-chat",
                "messages": [
                    {"role": "system", "content": "Ты нормализатор ФИО. Приведи ФИО к формату: ФАМИЛИЯ ИМЯ ОТЧЕСТВО. Только ФИО, ничего лишнего. Используй ЗАГЛАВНЫЕ буквы."},
                    {"role": "user", "content": f"Нормализуй: {fio_text}"}
                ],
                "temperature": 0.1,
                "max_tokens": 50
            }
            async with session.post(url, headers=headers, json=payload, timeout=10) as resp:
                data = await resp.json()
                if data.get("choices"):
                    result = data["choices"][0]["message"]["content"].strip().upper()
                    print(f"[DEEPSEEK] Нормализовано: {fio_text} -> {result}")
                    return result
    except Exception as e:
        print(f"[DEEPSEEK] Ошибка: {e}")
    return normalize_fio(fio_text)


def clean_phone(text):
    if not text:
        return ""
    d = re.sub(r'[^0-9]', '', str(text))
    if len(d) >= 10:
        return '+7' + d[-10:]
    return d


def clean_snils(text):
    if not text:
        return ""
    d = re.sub(r'[^0-9]', '', str(text))
    return d[:11] if len(d) >= 11 else d


def parse_date(val):
    if not val:
        return ""
    s = str(val).strip()
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
    d1 = re.sub(r'[^0-9]', '', str(p1 or ''))
    d2 = re.sub(r'[^0-9]', '', str(p2 or ''))
    if not d1 or not d2:
        return False
    return d1[-10:] == d2[-10:]


def parse_bot_response(text):
    """Парсит ответ бота: телефоны, СНИЛС, ИНН, паспорт"""
    result = {'phones': [], 'snils': [], 'inn': [], 'passport': []}
    phone_pattern = r'(?:\+?79\d{9})'
    for phone in re.findall(phone_pattern, text):
        clean = re.sub(r'[^0-9]', '', phone)
        if len(clean) == 11 and clean.startswith('79') and clean not in result['phones']:
            result['phones'].append(clean)
    for snil in re.findall(r'\b\d{11}\b', text):
        if snil not in result['snils']:
            result['snils'].append(snil)
    for inn in re.findall(r'\b\d{12}\b', text):
        if inn not in result['inn']:
            result['inn'].append(inn)
    for passport in re.findall(r'\b\d{10}\b', text):
        if passport not in result['passport']:
            result['passport'].append(passport)
    return result


def extract_phones_from_text(text):
    """Извлекает все телефоны из текста"""
    phones = []
    for phone in re.findall(r'(?:\+?79\d{9})', text):
        clean = re.sub(r'[^0-9]', '', phone)
        if len(clean) == 11 and clean.startswith('79'):
            phones.append(clean)
    return list(set(phones))


# ====================== MERGE / SPLIT ======================
def merge_tables(tables_data):
    """
    Объединяет несколько таблиц в одну.
    tables_data: список {headers: [...], rows: [[...]]}
    Возвращает: {headers: [...], rows: [[...]]}
    """
    if not tables_data:
        return None
    
    # Используем заголовки первой таблицы
    base_headers = list(tables_data[0]['headers'])
    addr_idx = -1
    for i, h in enumerate(base_headers):
        if h.lower() in ['адресс', 'адрес', 'address']:
            addr_idx = i
            break
    
    # Добавляем колонку "№ таблицы" перед адресом
    new_headers = []
    table_num_col = '№ таблицы'
    if addr_idx >= 0:
        new_headers = base_headers[:addr_idx] + [table_num_col] + base_headers[addr_idx:]
    else:
        new_headers = [table_num_col] + base_headers
    
    all_rows = []
    for table_idx, table in enumerate(tables_data, 1):
        for row in table['rows']:
            # Создаём словарь для быстрого доступа
            row_map = {}
            for i, h in enumerate(table['headers']):
                if i < len(row):
                    row_map[h.lower()] = row[i]
            
            # Формируем новую строку
            new_row = []
            for h in new_headers:
                if h == table_num_col:
                    new_row.append(str(table_idx))
                else:
                    new_row.append(row_map.get(h.lower(), ''))
            all_rows.append(new_row)
    
    return {'headers': new_headers, 'rows': all_rows}


def split_by_table_num(headers, rows):
    """
    Разбивает таблицу по колонке "№ таблицы"
    Возвращает: {table_num: {headers: [...], rows: [[...]]}}
    """
    table_num_idx = -1
    for i, h in enumerate(headers):
        if h == '№ таблицы':
            table_num_idx = i
            break
    
    if table_num_idx == -1:
        return None
    
    # Убираем колонку "№ таблицы" из заголовков
    clean_headers = [h for i, h in enumerate(headers) if i != table_num_idx]
    
    grouped = {}
    for row in rows:
        if len(row) <= table_num_idx:
            continue
        table_num = str(row[table_num_idx]).strip()
        if not table_num:
            continue
        clean_row = [v for i, v in enumerate(row) if i != table_num_idx]
        if table_num not in grouped:
            grouped[table_num] = []
        grouped[table_num].append(clean_row)
    
    result = {}
    for num, rows_data in grouped.items():
        result[num] = {'headers': clean_headers, 'rows': rows_data}
    
    return result


def geo_filter(headers, rows):
    """
    Фильтрует таблицу: оставляет только нужные колонки и размножает по номерам
    Возвращает: {headers: [...], rows: [[...]], phones: set}
    """
    GEO_COLS = ['Номер', 'Адресс', 'ФИО', 'Дата', 'СНИЛС']
    
    # Находим индексы нужных колонок
    idx_map = {}
    for col in GEO_COLS:
        found = -1
        for i, h in enumerate(headers):
            if h.lower() == col.lower():
                found = i
                break
        if found == -1:
            # Пробуем найти по частичному совпадению
            for i, h in enumerate(headers):
                if col.lower() in h.lower():
                    found = i
                    break
        idx_map[col] = found
    
    # Проверяем что все колонки найдены
    for col, idx in idx_map.items():
        if idx == -1:
            raise ValueError(f'Не найдена колонка: {col}')
    
    out_rows = []
    phones = set()
    
    for row in rows:
        # Извлекаем телефон
        phone_raw = row[idx_map['Номер']] if idx_map['Номер'] < len(row) else ''
        phone_clean = clean_phone(phone_raw)
        
        # Размножаем строку по всем телефонам в ячейке
        phones_in_cell = []
        if phone_clean:
            phones_in_cell.append(phone_clean)
        else:
            # Пробуем извлечь все телефоны из текста
            for p in extract_phones_from_text(str(phone_raw)):
                phones_in_cell.append(p)
        
        if not phones_in_cell:
            phones_in_cell = ['']
        
        for ph in phones_in_cell:
            if ph:
                phones.add(ph)
            out_rows.append([
                ph,
                row[idx_map['Адресс']] if idx_map['Адресс'] < len(row) else '',
                row[idx_map['ФИО']] if idx_map['ФИО'] < len(row) else '',
                row[idx_map['Дата']] if idx_map['Дата'] < len(row) else '',
                row[idx_map['СНИЛС']] if idx_map['СНИЛС'] < len(row) else ''
            ])
    
    return {'headers': GEO_COLS, 'rows': out_rows, 'phones': phones}


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
async def send_confirm_with_buttons(bot_token, chat_id, stage_name, count, confirm_id, show_stop=True, show_skip=True):
    """Отправляет подтверждение с кнопками"""
    global stop_requested
    
    if stop_requested:
        return False
    
    text = f"\u26a0\ufe0f ПОДТВЕРДИТЕ ПРОБИВ\n\nЭтап: {stage_name}\nСтрок: {count}"
    
    buttons = []
    row1 = [
        {"text": "\u2705 ПОДТВЕРДИТЬ", "callback_data": f"confirm_{confirm_id}"},
        {"text": "\u274c ОТМЕНА", "callback_data": f"cancel_{confirm_id}"}
    ]
    buttons.append(row1)
    
    if show_skip:
        buttons.append([{"text": "\u23f8 ПРОПУСТИТЬ ЭТАП", "callback_data": f"skip_{confirm_id}"}])
    
    if show_stop:
        buttons.append([{"text": "\u26d4 ОСТАНОВИТЬ ВСЁ", "callback_data": f"stop_{confirm_id}"}])
    
    kb = {"inline_keyboard": buttons}
    
    try:
        async with aiohttp.ClientSession() as s:
            await s.post(
                f"https://api.telegram.org/bot{bot_token}/sendMessage",
                json={"chat_id": int(chat_id), "text": text, "reply_markup": kb}
            )
        asyncio.create_task(poll_updates_with_buttons(bot_token, chat_id, confirm_id))
        return True
    except Exception as e:
        print(f"[CONFIRM] Ошибка: {e}")
        return False


async def poll_updates_with_buttons(bot_token, chat_id, confirm_id):
    """Ждет нажатия кнопок"""
    global stop_requested
    offset = 0
    start = time.time()
    print(f"[POLL] Начинаю опрос для confirm_id={confirm_id}")
    
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
                                    pending_confirms[confirm_id] = "confirm"
                                    print(f"[POLL] ПОДТВЕРЖДЕНО")
                                    return
                                    
                                elif cb_data == f"cancel_{confirm_id}":
                                    await s.post(
                                        f"https://api.telegram.org/bot{bot_token}/editMessageText",
                                        json={"chat_id": int(chat_id), "message_id": msg_id,
                                              "text": "\u274c ОТМЕНЕНО"}
                                    )
                                    pending_confirms[confirm_id] = "cancel"
                                    print(f"[POLL] ОТМЕНЕНО")
                                    return
                                    
                                elif cb_data == f"skip_{confirm_id}":
                                    await s.post(
                                        f"https://api.telegram.org/bot{bot_token}/editMessageText",
                                        json={"chat_id": int(chat_id), "message_id": msg_id,
                                              "text": "\u23f8 ЭТАП ПРОПУЩЕН"}
                                    )
                                    pending_confirms[confirm_id] = "skip"
                                    print(f"[POLL] ПРОПУЩЕН")
                                    return
                                    
                                elif cb_data == f"stop_{confirm_id}":
                                    await s.post(
                                        f"https://api.telegram.org/bot{bot_token}/editMessageText",
                                        json={"chat_id": int(chat_id), "message_id": msg_id,
                                              "text": "\u26d4 ОСТАНОВЛЕНО! Завершаю..."}
                                    )
                                    pending_confirms[confirm_id] = "stop"
                                    stop_requested = True
                                    print(f"[POLL] ОСТАНОВЛЕНО")
                                    return
        except Exception as e:
            print(f"[POLL] Ошибка: {e}")
        await asyncio.sleep(1)
    
    pending_confirms[confirm_id] = "timeout"
    print(f"[POLL] ТАЙМАУТ")


async def safe_confirm_with_buttons(bot_token, chat_id, stage_name, count, confirm_id, add_log, show_stop=True, show_skip=True):
    """Безопасное подтверждение с кнопками"""
    global stop_requested
    
    if stop_requested:
        add_log("[x] Остановка запрошена")
        return "stop"
    
    if not bot_token or not chat_id:
        add_log("[v] Бот не настроен — авто-продолжаю")
        return "confirm"

    sent = await send_confirm_with_buttons(bot_token, chat_id, stage_name, count, confirm_id, show_stop, show_skip)
    if not sent:
        add_log("[!] Не удалось отправить запрос в бот — авто-продолжаю")
        return "confirm"

    add_log(f"[ОЖИДАНИЕ] Откройте бот: {stage_name}")
    
    start = time.time()
    while time.time() - start < 600:
        if confirm_id in pending_confirms:
            r = pending_confirms.pop(confirm_id)
            if r == "stop":
                stop_requested = True
                add_log("[x] ОСТАНОВКА ВСЕХ ПРОЦЕССОВ")
                return "stop"
            if r == "skip":
                add_log(f"[v] ПРОПУЩЕН: {stage_name}")
                return "skip"
            if r == "confirm":
                add_log(f"[v] ПОДТВЕРЖДЕНО: {stage_name}")
                return "confirm"
            if r == "cancel":
                add_log(f"[x] ОТМЕНЁН: {stage_name}")
                return "cancel"
        await asyncio.sleep(0.5)
    
    add_log("[!] ТАЙМАУТ — продолжаю")
    return "confirm"


async def send_file_to_bot(bot_token, chat_id, filepath, caption=""):
    try:
        async with aiohttp.ClientSession() as s:
            data = aiohttp.FormData()
            data.add_field('chat_id', str(chat_id))
            data.add_field('caption', caption)
            data.add_field('document', open(filepath, 'rb'))
            await s.post(f"https://api.telegram.org/bot{bot_token}/sendDocument", data=data)
            print(f"[BOT] Файл отправлен: {caption}")
    except Exception as e:
        print(f"[BOT] Ошибка: {e}")


async def send_zip_to_bot(bot_token, chat_id, zip_path, caption=""):
    try:
        async with aiohttp.ClientSession() as s:
            data = aiohttp.FormData()
            data.add_field('chat_id', str(chat_id))
            data.add_field('caption', caption)
            data.add_field('document', open(zip_path, 'rb'))
            await s.post(f"https://api.telegram.org/bot{bot_token}/sendDocument", data=data)
            print(f"[BOT] ZIP отправлен: {caption}")
    except Exception as e:
        print(f"[BOT] Ошибка ZIP: {e}")


# ====================== РАБОТА С БОТАМИ ПРОБИВА ======================
async def clear_bot(client, bot):
    try:
        e = await client.get_entity(bot)
        await client.send_message(e, "/start")
        await asyncio.sleep(2)
        print(f"[BOT] /start отправлен в {bot}")
    except Exception as ex:
        print(f"[BOT] Ошибка: {ex}")


async def click_btn(client, bot, text, retries=3):
    e = await client.get_entity(bot)
    for attempt in range(retries):
        if attempt > 0:
            await asyncio.sleep(3)
        try:
            async for msg in client.iter_messages(e, limit=30):
                if msg.buttons and time.time() - msg.date.timestamp() < 300:
                    for row in msg.buttons:
                        for btn in row:
                            if btn.text and text.lower() in btn.text.lower():
                                await btn.click()
                                await asyncio.sleep(2)
                                print(f"[BOT] Нажата кнопка '{btn.text}'")
                                return True
            print(f"[BOT] Кнопка '{text}' не найдена (попытка {attempt+1})")
        except Exception as ex:
            print(f"[BOT] Ошибка: {ex}")
    return False


async def wait_xlsx(client, bot, timeout=180, since_msg_id=None):
    e = await client.get_entity(bot)
    start = time.time()
    print(f"[BOT] Ожидаю XLSX от {bot}...")
    while time.time() - start < timeout:
        msgs = await client.get_messages(e, limit=5)
        for msg in msgs:
            if not msg or not msg.document:
                continue
            if since_msg_id is not None and msg.id <= since_msg_id:
                continue
            for a in msg.document.attributes:
                if isinstance(a, DocumentAttributeFilename) and a.file_name.endswith('.xlsx'):
                    print(f"[BOT] Получен XLSX: {a.file_name}")
                    return msg
        await asyncio.sleep(3)
    print(f"[BOT] XLSX не получен")
    return None


# ====================== ПАРСИНГ XLSX ======================
def parse_xlsx(path):
    res = []
    try:
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        h = {}

        for col in range(1, ws.max_column + 1):
            v = str(ws.cell(row=1, column=col).value or "").upper().strip()
            if any(k in v for k in ['ИНН', 'INN', 'ПАСПОРТ']):
                continue
            if any(k in v for k in ['FIO', 'ФИО', 'ИМЯ', 'ФАМИЛИЯ', 'NAME']):
                h['fio'] = col
            if any(k in v for k in ['ДАТА', 'BIRTH', 'РОЖД', 'DATE']):
                h['date'] = col
            if any(k in v for k in ['ТЕЛЕФОН', 'PHONE', 'ТЕЛ']):
                h['phone'] = col
            elif 'НОМЕР' in v and 'ПАСПОРТ' not in v:
                h['phone'] = col
            if any(k in v for k in ['СНИЛС', 'SNILS']):
                h['snils'] = col

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
                if any(v for v in r.values() if v):
                    res.append(r)
            except Exception:
                pass
        return res
    except Exception as e:
        print(f"[PARSE] Ошибка: {e}")
        return []


# ====================== ЗАПОЛНЕНИЕ ТАБЛИЦЫ ======================
COL_NO = 1
COL_FIO = 2
COL_DATE = 3
COL_PHONE = 4
COL_SNILS = 5
COL_ADDR = 6


def fill_dates_from_response(ws, response_records):
    filled = 0
    print(f"[FILL-DATES] Начинаю заполнение дат. Ответов: {len(response_records)}")
    
    table_phones = {}
    for row in range(2, ws.max_row + 1):
        phone = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
        if phone:
            clean = clean_phone(phone)
            if clean:
                table_phones[clean] = row
    
    for rec in response_records:
        rec_phone = rec.get('phone', '')
        rec_date = rec.get('date', '')
        if not rec_date or not rec_phone:
            continue
        
        clean_rec = clean_phone(rec_phone)
        if not clean_rec:
            continue
        
        if clean_rec in table_phones:
            row = table_phones[clean_rec]
            existing = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
            if not existing or existing == 'None':
                ws.cell(row=row, column=COL_DATE).value = rec_date
                filled += 1
                print(f"[FILL-DATES] Строка {row}: ЗАПОЛНЕНО")
    
    print(f"[FILL-DATES] ИТОГО заполнено дат: {filled}")
    return filled


def fill_phones_from_response(ws, response_records):
    filled = 0
    print(f"[FILL-PHONES] Начинаю заполнение номеров. Ответов: {len(response_records)}")
    
    table_index = {}
    for row in range(2, ws.max_row + 1):
        fio = normalize_fio(str(ws.cell(row=row, column=COL_FIO).value or ""))
        date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        if fio and date and date != 'None':
            table_index[(fio, date)] = row
    
    for rec in response_records:
        rec_fio = normalize_fio(rec.get('fio', ''))
        rec_phone = rec.get('phone', '')
        rec_date = rec.get('date', '')
        
        if not rec_phone or not rec_fio:
            continue
        
        clean_rec = clean_phone(rec_phone)
        if not clean_rec:
            continue
        
        key = (rec_fio, rec_date)
        if key in table_index:
            row = table_index[key]
            existing = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
            if not existing or existing == 'None':
                ws.cell(row=row, column=COL_PHONE).value = clean_rec
                filled += 1
                print(f"[FILL-PHONES] Строка {row}: ЗАПОЛНЕНО")
        else:
            # fallback по ФИО
            for row in range(2, ws.max_row + 1):
                table_fio = normalize_fio(str(ws.cell(row=row, column=COL_FIO).value or ""))
                if table_fio == rec_fio:
                    existing = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
                    if not existing or existing == 'None':
                        ws.cell(row=row, column=COL_PHONE).value = clean_rec
                        filled += 1
                        print(f"[FILL-PHONES] Строка {row} (fallback): ЗАПОЛНЕНО")
                    break
    
    print(f"[FILL-PHONES] ИТОГО заполнено номеров: {filled}")
    return filled


def fill_snils_dates(ws, response_records):
    filled = 0
    print(f"[FILL-SNILS] Начинаю заполнение дат через СНИЛС")
    
    table_snils = {}
    for row in range(2, ws.max_row + 1):
        snils = clean_snils(str(ws.cell(row=row, column=COL_SNILS).value or ""))
        if snils and len(snils) >= 11:
            table_snils[snils] = row
    
    for rec in response_records:
        rec_snils = clean_snils(rec.get('snils', ''))
        rec_fio = normalize_fio(rec.get('fio', ''))
        rec_date = rec.get('date', '')
        
        if not rec_date:
            continue
        
        found = False
        
        if rec_snils and len(rec_snils) >= 11 and rec_snils in table_snils:
            row = table_snils[rec_snils]
            existing = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
            if not existing or existing == 'None':
                ws.cell(row=row, column=COL_DATE).value = rec_date
                filled += 1
                print(f"[FILL-SNILS] Строка {row}: ЗАПОЛНЕНО по СНИЛС")
                found = True
        
        if not found and rec_fio:
            for row in range(2, ws.max_row + 1):
                table_fio = normalize_fio(str(ws.cell(row=row, column=COL_FIO).value or ""))
                if table_fio == rec_fio:
                    existing = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
                    if not existing or existing == 'None':
                        ws.cell(row=row, column=COL_DATE).value = rec_date
                        filled += 1
                        print(f"[FILL-SNILS] Строка {row} (fallback): ЗАПОЛНЕНО")
                    found = True
                    break
    
    print(f"[FILL-SNILS] ИТОГО заполнено дат: {filled}")
    return filled


# ====================== ДОБИВ ЧЕРЕЗ САУРОН ======================
async def dobiv_sauron(client, bot, fio, date, account_id, row_num):
    """Добив номера через саурон (@proverim123_bot)"""
    try:
        query = f"{fio} {date}"
        print(f"[ДОБИВ] Акк {account_id}, строка {row_num}: {query}")
        
        await client.send_message(bot, query)
        await asyncio.sleep(5)
        
        async for msg in client.iter_messages(bot, limit=10):
            if msg.text and ("ОТЧЕТ" in msg.text or "ТЕЛЕФОНЫ" in msg.text):
                phones = extract_phones_from_text(msg.text)
                if phones:
                    print(f"[ДОБИВ] Найдены телефоны: {phones}")
                    return phones
                break
        
        return []
    except Exception as e:
        print(f"[ДОБИВ] Ошибка: {e}")
        return []


# ====================== ПОЛНЫЙ ЦИКЛ ПРОБИВА (7 ЭТАПОВ) ======================
async def run_full_cycle(ss, bot1, bot2, bot_token, chat_id,
                         items_no_date, items_no_phone, items_snils,
                         year_range, original_rows, tables_names=None):
    """
    Основной цикл пробива (7 этапов) с поддержкой ZIP и управления
    """
    global stop_requested
    stop_requested = False
    
    client = await get_client(ss)
    log = []

    def add(msg):
        ts = datetime.now().strftime('%H:%M:%S')
        log.append(f"[{ts}] {msg}")
        print(f"[LOG] {msg}")

    async def send_status(stage_label, file_path=None):
        if bot_token and chat_id:
            if file_path and os.path.exists(file_path):
                await send_file_to_bot(bot_token, chat_id, file_path, f"Таблица после: {stage_label}")
            else:
                wb.save(result_file)
                await send_file_to_bot(bot_token, chat_id, result_file, f"Таблица после: {stage_label}")

    async def send_final_zip():
        """Создаёт и отправляет ZIP архив со всеми таблицами"""
        if not bot_token or not chat_id:
            return
        
        # Разбиваем таблицу по колонке "№ таблицы"
        split_result = split_by_table_num(
            [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)],
            [[ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)] for r in range(2, ws.max_row + 1)]
        )
        
        if not split_result:
            await send_file_to_bot(bot_token, chat_id, result_file, "ИТОГОВЫЙ ФАЙЛ (все этапы)")
            return
        
        # Создаём ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            phones_all = set()
            
            for table_num, table_data in split_result.items():
                # Применяем GEO фильтр
                try:
                    geo_result = geo_filter(table_data['headers'], table_data['rows'])
                    phones_all.update(geo_result['phones'])
                    
                    # Создаём XLSX
                    ws_data = [geo_result['headers']] + geo_result['rows']
                    wb_temp = Workbook()
                    ws_temp = wb_temp.active
                    for row in ws_data:
                        ws_temp.append(row)
                    
                    xlsx_buffer = io.BytesIO()
                    wb_temp.save(xlsx_buffer)
                    xlsx_buffer.seek(0)
                    
                    # Имя файла
                    name = f"ГЕО_{table_num}.xlsx"
                    if tables_names and int(table_num) <= len(tables_names):
                        name = f"ГЕО_{tables_names[int(table_num)-1]}.xlsx"
                    zf.writestr(name, xlsx_buffer.getvalue())
                except Exception as e:
                    add(f"[ZIP] Ошибка таблицы {table_num}: {e}")
                    continue
            
            # Добавляем numbers.txt
            if phones_all:
                zf.writestr('numbers.txt', '\n'.join(sorted(phones_all)))
            else:
                zf.writestr('numbers.txt', '(нет валидных номеров)')
        
        zip_buffer.seek(0)
        zip_path = os.path.join(TEMP_DIR, f"result_{int(time.time())}.zip")
        with open(zip_path, 'wb') as f:
            f.write(zip_buffer.getvalue())
        
        await send_zip_to_bot(bot_token, chat_id, zip_path, "📦 ИТОГОВЫЙ ZIP АРХИВ (все таблицы + numbers.txt)")
        add("[v] ZIP архив отправлен")

    # === СОЗДАЁМ ИТОГОВУЮ ТАБЛИЦУ ===
    result_file = os.path.join(TEMP_DIR, f"result_{int(time.time())}.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.append(["N таблицы", "ФИО", "Дата", "Номер", "СНИЛС", "Адресс"])

    for i, row in enumerate(original_rows):
        ws.append([
            str(row[0] if len(row) > 0 else "").strip(),
            str(row[1] if len(row) > 1 else "").strip(),
            str(row[2] if len(row) > 2 else "").strip(),
            str(row[3] if len(row) > 3 else "").strip(),
            str(row[4] if len(row) > 4 else "").strip(),
            str(row[5] if len(row) > 5 else "").strip()
        ])
    wb.save(result_file)
    add(f"Таблица создана: {ws.max_row - 1} строк")

    # === НОРМАЛИЗАЦИЯ ФИО через DeepSeek ===
    add("[DEEPSEEK] Начинаю нормализацию ФИО...")
    for row in range(2, ws.max_row + 1):
        fio_val = str(ws.cell(row=row, column=COL_FIO).value or "").strip()
        if fio_val and fio_val != 'None':
            normalized = await normalize_fio_deepseek(fio_val)
            ws.cell(row=row, column=COL_FIO).value = normalized
            if row % 10 == 0:
                await asyncio.sleep(0.5)  # Не флудим API
    wb.save(result_file)
    add("[DEEPSEEK] Нормализация завершена")

    # === АНАЛИЗИРУЕМ ЧТО НУЖНО ПРОБИТЬ ===
    real_snils = []
    for row in range(2, ws.max_row + 1):
        existing_date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        snils_val = str(ws.cell(row=row, column=COL_SNILS).value or "").strip()
        if (not existing_date or existing_date == 'None') and snils_val and len(clean_snils(snils_val)) >= 11:
            real_snils.append(clean_snils(snils_val))
    real_snils = list(set(real_snils))

    add(f"К пробиву: дат={len(items_no_date)} номеров={len(items_no_phone)} снилс={len(real_snils)}")

    # ============ ЭТАП 1: ФИО+НОМЕР -> БОТ1 ============
    if items_no_date and not stop_requested:
        cid = f"s1_{int(time.time())}"
        add(f"ЭТАП 1: ФИО+НОМЕР -> бот1 ({len(items_no_date)} строк)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "ЭТАП 1: ФИО+НОМЕР", len(items_no_date), cid, add)
        if result == "stop":
            await send_final_zip()
            return {"ok": True, "log": log, "stopped": True}
        if result == "skip":
            add("[v] ЭТАП 1 ПРОПУЩЕН")
        elif result == "confirm":
            add("[v] ПОДТВЕРЖДЕНО! Начинаю этап 1...")

            txt = "\n".join([f"{it.get('fio','')}\t{it.get('phone','')}" for it in items_no_date])
            tpath = os.path.join(TEMP_DIR, f"t1_{int(time.time())}.txt")
            with open(tpath, 'w', encoding='utf-8') as f:
                f.write(txt)

            await clear_bot(client, bot1)
            e = await client.get_entity(bot1)
            await client.send_message(e, "Пробивы")
            await asyncio.sleep(2)
            await click_btn(client, bot1, "ФИО+номер")
            await asyncio.sleep(2)
            
            last_msgs = await client.get_messages(e, limit=1)
            last_msg_id = last_msgs[0].id if last_msgs else 0
            await client.send_file(e, tpath)
            add("Файл отправлен в бот1, жду ответ...")

            msg = await wait_xlsx(client, bot1, 300, since_msg_id=last_msg_id)
            if msg:
                rpath = os.path.join(TEMP_DIR, f"r1_{int(time.time())}.xlsx")
                await client.download_media(msg, file=rpath)
                recs = parse_xlsx(rpath)
                add(f"Получено ответов: {len(recs)}")
                fill_dates_from_response(ws, recs)
                wb.save(result_file)
                await send_status("этап 1")
            else:
                add("[!] Бот1 не ответил")
    
    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # === ПЕРЕСЧЁТ СНИЛС ===
    real_snils = []
    for row in range(2, ws.max_row + 1):
        existing_date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        snils_val = str(ws.cell(row=row, column=COL_SNILS).value or "").strip()
        if (not existing_date or existing_date == 'None') and snils_val and len(clean_snils(snils_val)) >= 11:
            real_snils.append(clean_snils(snils_val))
    real_snils = list(set(real_snils))
    add(f"[ПЕРЕСЧЁТ] СНИЛС с пустой датой: {len(real_snils)}")

    # ============ ЭТАП 2: СНИЛС -> БОТ1 ============
    if real_snils and not stop_requested:
        cid = f"s2_{int(time.time())}"
        add(f"ЭТАП 2: СНИЛС -> бот1 ({len(real_snils)} снилс)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "ЭТАП 2: СНИЛС (бот1)", len(real_snils), cid, add)
        if result == "stop":
            await send_final_zip()
            return {"ok": True, "log": log, "stopped": True}
        if result == "skip":
            add("[v] ЭТАП 2 ПРОПУЩЕН")
        elif result == "confirm":
            add("[v] ПОДТВЕРЖДЕНО! Начинаю этап 2...")

            txt = "\n".join(real_snils)
            tpath = os.path.join(TEMP_DIR, f"t2_{int(time.time())}.txt")
            with open(tpath, 'w', encoding='utf-8') as f:
                f.write(txt)

            await clear_bot(client, bot1)
            e = await client.get_entity(bot1)
            await client.send_message(e, "Пробивы")
            await asyncio.sleep(4)
            await click_btn(client, bot1, "СНИЛС")
            await asyncio.sleep(3)
            
            last_msgs = await client.get_messages(e, limit=1)
            last_msg_id = last_msgs[0].id if last_msgs else 0
            await client.send_file(e, tpath)
            add("СНИЛС отправлены в бот1, жду ответ...")

            msg = await wait_xlsx(client, bot1, 300, since_msg_id=last_msg_id)
            if msg:
                rpath = os.path.join(TEMP_DIR, f"r2_{int(time.time())}.xlsx")
                await client.download_media(msg, file=rpath)
                recs = parse_xlsx(rpath)
                add(f"Получено через СНИЛС: {len(recs)}")
                fill_snils_dates(ws, recs)
                wb.save(result_file)
                await send_status("этап 2")
            else:
                add("[!] Бот1 не ответил")
    
    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # ============ ЭТАП 3: СНИЛС -> БОТ2 ============
    snils_still_empty = []
    for row in range(2, ws.max_row + 1):
        existing_date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        snils_val = str(ws.cell(row=row, column=COL_SNILS).value or "").strip()
        if (not existing_date or existing_date == 'None') and snils_val and len(clean_snils(snils_val)) >= 11:
            snils_still_empty.append(clean_snils(snils_val))
    snils_still_empty = list(set(snils_still_empty))

    if snils_still_empty and not stop_requested:
        cid = f"s3_{int(time.time())}"
        add(f"ЭТАП 3: СНИЛС -> бот2 ({len(snils_still_empty)} снилс)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "ЭТАП 3: СНИЛС (бот2)", len(snils_still_empty), cid, add)
        if result == "stop":
            await send_final_zip()
            return {"ok": True, "log": log, "stopped": True}
        if result == "skip":
            add("[v] ЭТАП 3 ПРОПУЩЕН")
        elif result == "confirm":
            add("[v] ПОДТВЕРЖДЕНО! Начинаю этап 3...")

            txt = "\n".join(snils_still_empty)
            tpath = os.path.join(TEMP_DIR, f"t3_{int(time.time())}.txt")
            with open(tpath, 'w', encoding='utf-8') as f:
                f.write(txt)

            e = await client.get_entity(bot2)
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
                await send_status("этап 3")
            else:
                add("[!] Бот2 не ответил")
    
    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # ============ ЭТАП 4: ФИЛЬТР ПО ГОДАМ ============
    if year_range and not stop_requested:
        try:
            parts = year_range.split('-')
            yf, yt = int(parts[0]), int(parts[1])
            add(f"ЭТАП 4: Фильтр годов {yf}-{yt}")

            # Сначала проверяем кнопку
            cid = f"s4_{int(time.time())}"
            result = await safe_confirm_with_buttons(bot_token, chat_id, f"ЭТАП 4: Фильтр {yf}-{yt}", ws.max_row - 1, cid, add)
            if result == "stop":
                await send_final_zip()
                return {"ok": True, "log": log, "stopped": True}
            if result == "skip":
                add("[v] ЭТАП 4 ПРОПУЩЕН")
            elif result == "confirm":
                add("[v] ПОДТВЕРЖДЕНО! Фильтрую...")
                
                rows_to_delete = []
                for row in range(2, ws.max_row + 1):
                    date_val = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
                    parts_d = date_val.split('.')
                    if len(parts_d) == 3:
                        try:
                            year = int(parts_d[2])
                            if year < yf or year > yt:
                                rows_to_delete.append(row)
                        except ValueError:
                            rows_to_delete.append(row)
                    else:
                        rows_to_delete.append(row)

                for row in reversed(rows_to_delete):
                    ws.delete_rows(row)
                wb.save(result_file)
                await send_status(f"этап 4 (фильтр {yf}-{yt})")
                add(f"Удалено: {len(rows_to_delete)}, осталось: {ws.max_row - 1}")
        except Exception as e:
            add(f"Ошибка фильтра: {e}")
    
    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # ============ ЭТАП 5: ФИО+ДАТА -> БОТ1 ============
    items_no_phone_after_stage4 = []
    for row in range(2, ws.max_row + 1):
        fio_val = str(ws.cell(row=row, column=COL_FIO).value or "").strip()
        date_val = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        phone_val = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
        if (fio_val and fio_val != 'None'
                and date_val and date_val != 'None' and date_val != '0'
                and (not phone_val or phone_val == 'None' or phone_val == '0')):
            items_no_phone_after_stage4.append({
                'fio': normalize_fio(fio_val),
                'date': date_val
            })

    if items_no_phone_after_stage4 and not stop_requested:
        cid = f"s5_{int(time.time())}"
        add(f"ЭТАП 5: ФИО+ДАТА -> бот1 ({len(items_no_phone_after_stage4)} строк)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "ЭТАП 5: ФИО+ДАТА (бот1)", len(items_no_phone_after_stage4), cid, add)
        if result == "stop":
            await send_final_zip()
            return {"ok": True, "log": log, "stopped": True}
        if result == "skip":
            add("[v] ЭТАП 5 ПРОПУЩЕН")
        elif result == "confirm":
            add("[v] ПОДТВЕРЖДЕНО! Начинаю этап 5...")

            txt = "\n".join([f"{it.get('fio','')}\t{it.get('date','')}" for it in items_no_phone_after_stage4])
            tpath = os.path.join(TEMP_DIR, f"t5_{int(time.time())}.txt")
            with open(tpath, 'w', encoding='utf-8') as f:
                f.write(txt)

            await clear_bot(client, bot1)
            e = await client.get_entity(bot1)
            await client.send_message(e, "Пробивы")
            await asyncio.sleep(2)
            await click_btn(client, bot1, "ФИО+дата")
            await asyncio.sleep(2)
            
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
                await send_status("этап 5")
            else:
                add("[!] Бот1 не ответил")
    
    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # ============ ЭТАП 6: ФИО+ДАТА -> БОТ2 ============
    items_no_phone_after_stage5 = []
    for row in range(2, ws.max_row + 1):
        fio_val = str(ws.cell(row=row, column=COL_FIO).value or "").strip()
        date_val = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        phone_val = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
        if (fio_val and fio_val != 'None'
                and date_val and date_val != 'None' and date_val != '0'
                and (not phone_val or phone_val == 'None' or phone_val == '0')):
            items_no_phone_after_stage5.append({
                'fio': normalize_fio(fio_val),
                'date': date_val
            })

    if items_no_phone_after_stage5 and not stop_requested:
        cid = f"s6_{int(time.time())}"
        add(f"ЭТАП 6: ФИО+ДАТА -> бот2 ({len(items_no_phone_after_stage5)} строк)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "ЭТАП 6: ФИО+ДАТА (бот2)", len(items_no_phone_after_stage5), cid, add)
        if result == "stop":
            await send_final_zip()
            return {"ok": True, "log": log, "stopped": True}
        if result == "skip":
            add("[v] ЭТАП 6 ПРОПУЩЕН")
        elif result == "confirm":
            add("[v] ПОДТВЕРЖДЕНО! Начинаю этап 6...")

            txt = "\n".join([f"{it.get('fio','')}\t{it.get('date','')}" for it in items_no_phone_after_stage5])
            tpath = os.path.join(TEMP_DIR, f"t6_{int(time.time())}.txt")
            with open(tpath, 'w', encoding='utf-8') as f:
                f.write(txt)

            e = await client.get_entity(bot2)
            last_msgs = await client.get_messages(e, limit=1)
            last_msg_id = last_msgs[0].id if last_msgs else 0
            await client.send_file(e, tpath)
            add("Файл отправлен в бот2, жду ответ...")

            msg = await wait_xlsx(client, bot2, 300, since_msg_id=last_msg_id)
            if msg:
                rpath = os.path.join(TEMP_DIR, f"r6_{int(time.time())}.xlsx")
                await client.download_media(msg, file=rpath)
                recs = parse_xlsx(rpath)
                add(f"Получено ответов от бот2: {len(recs)}")
                fill_phones_from_response(ws, recs)
                wb.save(result_file)
                await send_status("этап 6")
            else:
                add("[!] Бот2 не ответил")
    
    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # ============ ЭТАП 7: ДОБИВ -> БОТ2 (саурон) ============
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

    if items_no_phone_final and not stop_requested:
        cid = f"s7_{int(time.time())}"
        add(f"ЭТАП 7: ДОБИВ -> бот2 ({len(items_no_phone_final)} строк)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "ЭТАП 7: ДОБИВ (бот2)", len(items_no_phone_final), cid, add)
        if result == "stop":
            await send_final_zip()
            return {"ok": True, "log": log, "stopped": True}
        if result == "skip":
            add("[v] ЭТАП 7 ПРОПУЩЕН")
        elif result == "confirm":
            add("[v] ПОДТВЕРЖДЕНО! Добиваю через саурон...")

            e = await client.get_entity(bot2)
            for i, it in enumerate(items_no_phone_final):
                if stop_requested:
                    break
                    
                try:
                    fio = it.get('fio', '')
                    date = it.get('date', '')
                    
                    phones = await dobiv_sauron(client, e, fio, date, "1", i+1)
                    
                    if phones:
                        for row in range(2, ws.max_row + 1):
                            table_fio = normalize_fio(str(ws.cell(row=row, column=COL_FIO).value or ""))
                            table_date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
                            existing_phone = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()

                            if table_fio == normalize_fio(fio) and table_date == date:
                                if not existing_phone or existing_phone == 'None':
                                    ws.cell(row=row, column=COL_PHONE).value = phones[0]
                                    add(f"  Добив: {fio} -> {phones[0]}")
                                break

                    if (i + 1) % 5 == 0:
                        wb.save(result_file)
                        add(f"  Сохранено после {i+1} строк")
                        await asyncio.sleep(2)

                except FloodWaitError as e:
                    add(f"  FloodWait: {e.seconds}с")
                    await asyncio.sleep(e.seconds)
                except Exception as e:
                    add(f"  Ошибка добива: {e}")

            wb.save(result_file)
            await send_status("этап 7 (добив)")
            add("Добив завершён")

    # === УДАЛЯЕМ СТРОКИ БЕЗ ДАТЫ ===
    rows_no_date = []
    for row in range(2, ws.max_row + 1):
        date_val = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        if not date_val or date_val == 'None' or date_val == '0':
            rows_no_date.append(row)
    
    if rows_no_date:
        add(f"[ОЧИСТКА] Удаляю {len(rows_no_date)} строк без даты")
        for row in reversed(rows_no_date):
            ws.delete_rows(row)
        wb.save(result_file)
        add(f"Осталось: {ws.max_row - 1} строк")

    # === ФИНАЛ ===
    add("=== ВСЕ ЭТАПЫ ЗАВЕРШЕНЫ ===")
    
    # Отправляем финальный ZIP
    await send_final_zip()
    
    # Статистика
    total_dates = 0
    total_phones = 0
    for row in range(2, ws.max_row + 1):
        if str(ws.cell(row=row, column=COL_DATE).value or "").strip():
            total_dates += 1
        if str(ws.cell(row=row, column=COL_PHONE).value or "").strip():
            total_phones += 1
    add(f"Итого: строк={ws.max_row-1}, с датами={total_dates}, с номерами={total_phones}")

    return {"ok": True, "log": log, "stopped": stop_requested}


# ====================== ЭНДПОИНТЫ ======================
async def handle_health(request):
    return web.json_response({"ok": True, "message": "X Backend v12.0"})


async def handle_root(request):
    try:
        html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "сайт.html")
        with open(html_path, "r", encoding="utf-8") as f:
            html = f.read()
        return web.Response(text=html, content_type="text/html")
    except Exception as e:
        print(f"[ROOT] Ошибка: {e}")
        return web.Response(text="OK", content_type="text/plain")


async def handle_ping(request):
    return web.Response(text="pong", content_type="text/plain")


async def handle_favicon(request):
    return web.Response(status=204)


async def handle_upload_zip(request):
    """Загрузка ZIP архива с таблицами"""
    try:
        reader = await request.multipart()
        field = await reader.next()
        if field.name != 'file':
            return web.json_response({"ok": False, "error": "No file"}, status=400)
        
        data = await field.read()
        zip_path = os.path.join(TEMP_DIR, f"upload_{int(time.time())}.zip")
        with open(zip_path, 'wb') as f:
            f.write(data)
        
        # Распаковываем ZIP
        tables_data = []
        with zipfile.ZipFile(zip_path, 'r') as zf:
            for file_info in zf.filelist:
                if file_info.filename.endswith('.xlsx'):
                    with zf.open(file_info) as f:
                        xlsx_data = f.read()
                        xlsx_path = os.path.join(TEMP_DIR, f"tmp_{int(time.time())}_{file_info.filename}")
                        with open(xlsx_path, 'wb') as out:
                            out.write(xlsx_data)
                        
                        # Парсим XLSX
                        wb = load_workbook(xlsx_path, data_only=True)
                        ws = wb.active
                        headers = [str(cell.value or "").strip() for cell in ws[1]]
                        rows = []
                        for row in ws.iter_rows(min_row=2, values_only=True):
                            if any(cell for cell in row):
                                rows.append([str(cell or "").strip() for cell in row])
                        tables_data.append({'headers': headers, 'rows': rows})
                        os.remove(xlsx_path)
        
        # Объединяем таблицы
        merged = merge_tables(tables_data)
        if not merged:
            return web.json_response({"ok": False, "error": "No tables found"}, status=400)
        
        # Сохраняем объединённую таблицу
        merged_path = os.path.join(TEMP_DIR, f"merged_{int(time.time())}.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.append(merged['headers'])
        for row in merged['rows']:
            ws.append(row)
        wb.save(merged_path)
        
        return web.json_response({
            "ok": True,
            "file": merged_path,
            "headers": merged['headers'],
            "rows": merged['rows'],
            "count": len(merged['rows'])
        })
    except Exception as e:
        print(f"[UPLOAD] Ошибка: {e}")
        traceback.print_exc()
        return web.json_response({"ok": False, "error": str(e)}, status=500)


async def handle_send_code(request):
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
        print(f"[AUTH] Ошибка: {e}")
        return web.json_response({"ok": False, "error": str(e)}, status=400)


async def handle_verify_code(request):
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
        print(f"[AUTH] Успешный вход: {me.first_name}")
        return web.json_response({
            "ok": True,
            "session": ss,
            "phone": phone,
            "name": me.first_name or ""
        })
    except Exception as e:
        print(f"[AUTH] Ошибка: {e}")
        return web.json_response({"ok": False, "error": str(e)}, status=400)


async def handle_full_probev(request):
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
        tables_names = d.get("tables_names", [])

        if not ss:
            return web.json_response({"ok": False, "error": "Нет сессии"}, status=400)

        if ss in active_probevs:
            return web.json_response({"ok": False, "error": "Пробив уже выполняется"}, status=409)
        active_probevs.add(ss)

        try:
            print(f"\n{'='*60}")
            print(f"[PROBEV] ЗАПУСК ПОЛНОГО ЦИКЛА (7 ЭТАПОВ + DeepSeek)")
            print(f"[PROBEV] Бот1: {bot1}, Бот2: {bot2}")
            print(f"[PROBEV] Строк без даты: {len(items_no_date)}")
            print(f"[PROBEV] Строк без номера: {len(items_no_phone)}")
            print(f"{'='*60}\n")

            result = await run_full_cycle(
                ss, bot1, bot2, bot_token, chat_id,
                items_no_date, items_no_phone, items_snils,
                year_range, original_rows, tables_names
            )
            return web.json_response(result)
        finally:
            active_probevs.discard(ss)
    except Exception as e:
        traceback.print_exc()
        return web.json_response({"ok": False, "error": str(e)}, status=400)


async def handle_stop(request):
    """Принудительная остановка"""
    global stop_requested
    stop_requested = True
    return web.json_response({"ok": True, "message": "Stop requested"})


# ====================== ЗАПУСК ======================
app = web.Application(middlewares=[log_and_cors], client_max_size=100 * 1024 * 1024)
app.router.add_get("/", handle_root)
app.router.add_get("/health", handle_health)
app.router.add_get("/ping", handle_ping)
app.router.add_get("/favicon.ico", handle_favicon)
app.router.add_post("/upload-zip", handle_upload_zip)
app.router.add_post("/send-code", handle_send_code)
app.router.add_post("/verify-code", handle_verify_code)
app.router.add_post("/full-probev", handle_full_probev)
app.router.add_post("/stop", handle_stop)


async def on_startup(app):
    port = app["port"]
    msg = (
        "=" * 60 + "\n"
        f"X Backend v12.0 ЗАПУЩЕН (DeepSeek + 7 этапов + ZIP)\n"
        f"Host: 0.0.0.0  |  Port: {port}\n"
        + "=" * 60
    )
    print(msg, flush=True)


async def on_shutdown(app):
    print("[SHUTDOWN] Закрываю соединения...", flush=True)
    for ss, client in list(user_clients.items()):
        try:
            if client.is_connected():
                await client.disconnect()
        except Exception:
            pass
    user_clients.clear()
    sessions.clear()
    pending_confirms.clear()
    print("[SHUTDOWN] Завершено.", flush=True)


if __name__ == "__main__":
    import sys, signal as _signal
    try:
        port = int(os.environ.get("PORT", 4545))
        app["port"] = port
        app.on_startup.append(on_startup)
        app.on_shutdown.append(on_shutdown)

        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

        runner = web.AppRunner(app)
        loop.run_until_complete(runner.setup())

        import socket as _socket
        sock = _socket.socket(_socket.AF_INET, _socket.SOCK_STREAM)
        sock.setsockopt(_socket.SOL_SOCKET, _socket.SO_REUSEADDR, 1)
        try:
            sock.setsockopt(_socket.SOL_SOCKET, _socket.SO_REUSEPORT, 1)
        except:
            pass
        sock.bind(("0.0.0.0", port))
        sock.listen(128)
        sock.setblocking(False)

        site = web.SockSite(runner, sock)
        loop.run_until_complete(site.start())

        print(f"[START] СЕРВЕР ГОТОВ — порт {port}", flush=True)

        def shutdown():
            print("[SIGNAL] Остановка...", flush=True)
            loop.create_task(_do_shutdown(runner, sock, loop))

        async def _do_shutdown(runner, sock, loop):
            await runner.cleanup()
            sock.close()
            loop.stop()

        for sig in (_signal.SIGTERM, _signal.SIGINT):
            try:
                loop.add_signal_handler(sig, shutdown)
            except NotImplementedError:
                pass

        loop.run_forever()
        loop.close()
    except Exception as e:
        print(f"[FATAL] {e}", flush=True)
        traceback.print_exc()
        sys.exit(1)