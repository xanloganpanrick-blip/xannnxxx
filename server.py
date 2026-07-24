# server.py - X Backend v14.0 (ПЕРЕИМЕНОВАНИЕ ФАЙЛОВ ПО АДРЕСУ + ПОИСК ИНН В ГРУППЕ)
# Установка: pip install aiohttp telethon openpyxl
# Запуск: python server.py
# Порт: 8765

import asyncio
import json
import os
import re
import time
import traceback
import io
import zipfile
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
active_probevs = {}
probev_lock = asyncio.Lock()
stop_requested = False

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
def normalize_fio_local(raw):
    if not raw:
        return ""
    cleaned = re.sub(r'[^A-Za-zА-ЯЁа-яё\-]', ' ', str(raw))
    words = [w.strip() for w in cleaned.split() if w.strip()]
    if not words:
        return ""
    return ' '.join(words).upper().replace('Ё', 'Е')


def normalize_address_local(raw):
    """Локальная нормализация адреса (без API)"""
    if not raw:
        return ""
    # Убираем лишние пробелы, приводим к нижнему регистру
    s = str(raw).strip().lower()
    # Заменяем распространённые сокращения
    replacements = {
        'обл': 'область', 'обл.': 'область',
        'г': 'город', 'г.': 'город',
        'ул': 'улица', 'ул.': 'улица',
        'д': 'дом', 'д.': 'дом',
        'кв': 'квартира', 'кв.': 'квартира',
        'корп': 'корпус', 'корп.': 'корпус',
        'стр': 'строение', 'стр.': 'строение',
    }
    for old, new in replacements.items():
        s = s.replace(old, new)
    # Убираем лишние пробелы
    s = re.sub(r'\s+', ' ', s).strip()
    return s


async def normalize_batch_deepseek(items, prompt_type='fio'):
    """Пакетная нормализация через DeepSeek"""
    if not items:
        return []
    
    if len(items) <= 2:
        if prompt_type == 'fio':
            return [normalize_fio_local(f) for f in items]
        else:
            return [normalize_address_local(a) for a in items]
    
    try:
        async with aiohttp.ClientSession() as session:
            url = "https://api.deepseek.com/v1/chat/completions"
            headers = {
                "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
                "Content-Type": "application/json"
            }
            
            if prompt_type == 'fio':
                system_prompt = "Normalize each FIO to format: LASTNAME FIRSTNAME PATRONYMIC. Return only normalized list, one per line, numbered. Use UPPERCASE."
                items_text = "\n".join([f"{i+1}. {f}" for i, f in enumerate(items)])
            elif prompt_type == 'address':
                system_prompt = "Normalize each address to format: CITY, STREET, HOUSE, KV. Example: 'Тольятти, Голосова, 26, кв. 33'. Remove 'кв.' from output. Return only normalized list, one per line, numbered."
                items_text = "\n".join([f"{i+1}. {a}" for i, a in enumerate(items)])
            else:
                return items
            
            payload = {
                "model": "deepseek-chat",
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"Normalize these:\n{items_text}"}
                ],
                "temperature": 0.1,
                "max_tokens": 800
            }
            async with session.post(url, headers=headers, json=payload, timeout=15) as resp:
                data = await resp.json()
                if data.get("choices"):
                    text = data["choices"][0]["message"]["content"].strip()
                    result = []
                    for line in text.split('\n'):
                        line = line.strip()
                        if not line:
                            continue
                        if re.match(r'^\d+\.', line):
                            line = re.sub(r'^\d+\.\s*', '', line)
                        if prompt_type == 'fio':
                            result.append(line.upper())
                        else:
                            # Убираем "кв." из адреса
                            line = re.sub(r'\s*кв\.?\s*', ', ', line)
                            line = re.sub(r',\s*,', ',', line)
                            result.append(line)
                    return result
    except Exception as e:
        print(f"[DEEPSEEK] Batch error: {e}")
    
    # Fallback
    if prompt_type == 'fio':
        return [normalize_fio_local(f) for f in items]
    return [normalize_address_local(a) for a in items]


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
    phones = []
    for phone in re.findall(r'(?:\+?79\d{9})', text):
        clean = re.sub(r'[^0-9]', '', phone)
        if len(clean) == 11 and clean.startswith('79'):
            phones.append(clean)
    return list(set(phones))


def extract_inn_from_text(text):
    """Извлекает ИНН (12 цифр) из текста"""
    inn_match = re.search(r'\b\d{12}\b', text)
    if inn_match:
        return inn_match.group()
    return None


# ====================== MERGE / SPLIT ======================
def merge_tables(tables_data):
    if not tables_data:
        return None
    
    base_headers = list(tables_data[0]['headers'])
    addr_idx = -1
    for i, h in enumerate(base_headers):
        if h.lower() in ['адресс', 'адрес', 'address']:
            addr_idx = i
            break
    
    table_num_col = 'N tablicy'
    new_headers = []
    if addr_idx >= 0:
        new_headers = base_headers[:addr_idx] + [table_num_col] + base_headers[addr_idx:]
    else:
        new_headers = [table_num_col] + base_headers
    
    all_rows = []
    for table_idx, table in enumerate(tables_data, 1):
        for row in table['rows']:
            row_map = {}
            for i, h in enumerate(table['headers']):
                if i < len(row):
                    row_map[h.lower()] = row[i]
            
            new_row = []
            for h in new_headers:
                if h == table_num_col:
                    new_row.append(str(table_idx))
                else:
                    new_row.append(row_map.get(h.lower(), ''))
            all_rows.append(new_row)
    
    return {'headers': new_headers, 'rows': all_rows}


def split_by_table_num(headers, rows):
    table_num_idx = -1
    for i, h in enumerate(headers):
        if h == 'N tablicy':
            table_num_idx = i
            break
    
    if table_num_idx == -1:
        return None
    
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
    GEO_COLS = ['Nomer', 'Adress', 'FIO', 'Data', 'SNILS']
    
    idx_map = {}
    for col in GEO_COLS:
        found = -1
        for i, h in enumerate(headers):
            if h.lower() == col.lower():
                found = i
                break
        if found == -1:
            for i, h in enumerate(headers):
                if col.lower() in h.lower():
                    found = i
                    break
        idx_map[col] = found
    
    for col, idx in idx_map.items():
        if idx == -1:
            raise ValueError(f'Column not found: {col}')
    
    out_rows = []
    phones = set()
    
    for row in rows:
        phone_raw = row[idx_map['Nomer']] if idx_map['Nomer'] < len(row) else ''
        phone_clean = clean_phone(phone_raw)
        
        phones_in_cell = []
        if phone_clean:
            phones_in_cell.append(phone_clean)
        else:
            for p in extract_phones_from_text(str(phone_raw)):
                phones_in_cell.append(p)
        
        if not phones_in_cell:
            phones_in_cell = ['']
        
        for ph in phones_in_cell:
            if ph:
                phones.add(ph)
            out_rows.append([
                ph,
                row[idx_map['Adress']] if idx_map['Adress'] < len(row) else '',
                row[idx_map['FIO']] if idx_map['FIO'] < len(row) else '',
                row[idx_map['Data']] if idx_map['Data'] < len(row) else '',
                row[idx_map['SNILS']] if idx_map['SNILS'] < len(row) else ''
            ])
    
    return {'headers': GEO_COLS, 'rows': out_rows, 'phones': phones}


# ====================== TELEGRAM CLIENT ======================
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


# ====================== BOT CONFIRMATIONS ======================
async def send_confirm_with_buttons(bot_token, chat_id, stage_name, count, confirm_id, topic_id=None):
    global stop_requested
    
    if stop_requested:
        return False
    
    text = f"CONFIRM PROBEV\n\nStage: {stage_name}\nRows: {count}"
    
    buttons = [
        [{"text": "CONFIRM", "callback_data": f"confirm_{confirm_id}"}],
        [{"text": "SKIP", "callback_data": f"skip_{confirm_id}"}],
        [{"text": "STOP ALL", "callback_data": f"stop_{confirm_id}"}],
        [{"text": "AGAIN", "callback_data": f"again_{confirm_id}"}]
    ]
    
    kb = {"inline_keyboard": buttons}
    
    payload = {"chat_id": int(chat_id), "text": text, "reply_markup": kb}
    if topic_id:
        payload["message_thread_id"] = int(topic_id)
    
    try:
        async with aiohttp.ClientSession() as s:
            await s.post(
                f"https://api.telegram.org/bot{bot_token}/sendMessage",
                json=payload
            )
        asyncio.create_task(poll_updates_with_buttons(bot_token, chat_id, confirm_id, topic_id))
        return True
    except Exception as e:
        print(f"[CONFIRM] Error: {e}")
        return False


async def poll_updates_with_buttons(bot_token, chat_id, confirm_id, topic_id=None):
    global stop_requested
    offset = 0
    print(f"[POLL] Starting poll for confirm_id={confirm_id}")
    
    while True:
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
                            if not cb:
                                continue
                            
                            msg = cb.get("message", {})
                            msg_chat_id = str(msg.get("chat", {}).get("id"))
                            msg_topic_id = msg.get("message_thread_id")
                            
                            if msg_chat_id != str(chat_id):
                                continue
                            if topic_id and msg_topic_id != int(topic_id):
                                continue
                            
                            cb_data = cb.get("data", "")
                            await s.post(
                                f"https://api.telegram.org/bot{bot_token}/answerCallbackQuery",
                                json={"callback_query_id": cb["id"]}
                            )
                            msg_id = msg.get("message_id")
                            
                            payload = {"chat_id": int(chat_id), "message_id": msg_id}
                            if topic_id:
                                payload["message_thread_id"] = int(topic_id)
                            
                            if cb_data == f"confirm_{confirm_id}":
                                payload["text"] = "CONFIRMED! Executing..."
                                await s.post(
                                    f"https://api.telegram.org/bot{bot_token}/editMessageText",
                                    json=payload
                                )
                                pending_confirms[confirm_id] = "confirm"
                                print(f"[POLL] CONFIRMED")
                                return
                                    
                            elif cb_data == f"skip_{confirm_id}":
                                payload["text"] = "STAGE SKIPPED"
                                await s.post(
                                    f"https://api.telegram.org/bot{bot_token}/editMessageText",
                                    json=payload
                                )
                                pending_confirms[confirm_id] = "skip"
                                print(f"[POLL] SKIPPED")
                                return
                                    
                            elif cb_data == f"stop_{confirm_id}":
                                payload["text"] = "STOPPED! Finishing..."
                                await s.post(
                                    f"https://api.telegram.org/bot{bot_token}/editMessageText",
                                    json=payload
                                )
                                pending_confirms[confirm_id] = "stop"
                                stop_requested = True
                                print(f"[POLL] STOPPED")
                                return
                                    
                            elif cb_data == f"again_{confirm_id}":
                                payload["text"] = "AGAIN requested! Resending..."
                                await s.post(
                                    f"https://api.telegram.org/bot{bot_token}/editMessageText",
                                    json=payload
                                )
                                pending_confirms[confirm_id] = "again"
                                print(f"[POLL] AGAIN")
                                return
        except Exception as e:
            print(f"[POLL] Error: {e}")
        await asyncio.sleep(1)


async def safe_confirm_with_buttons(bot_token, chat_id, stage_name, count, confirm_id, add_log, topic_id=None):
    global stop_requested
    
    if stop_requested:
        add_log("[x] Stop requested")
        return "stop"
    
    if not bot_token or not chat_id:
        add_log("[v] Bot not configured - auto continue")
        return "confirm"

    while True:
        sent = await send_confirm_with_buttons(bot_token, chat_id, stage_name, count, confirm_id, topic_id)
        if not sent:
            add_log("[!] Failed to send to bot - retrying in 5s...")
            await asyncio.sleep(5)
            continue

        add_log(f"[WAITING] Open bot for: {stage_name}")
        
        while True:
            if confirm_id in pending_confirms:
                r = pending_confirms.pop(confirm_id)
                if r == "stop":
                    stop_requested = True
                    add_log("[x] STOP ALL PROCESSES")
                    return "stop"
                if r == "skip":
                    add_log(f"[v] SKIPPED: {stage_name}")
                    return "skip"
                if r == "confirm":
                    add_log(f"[v] CONFIRMED: {stage_name}")
                    return "confirm"
                if r == "again":
                    add_log(f"[v] AGAIN - resending confirmation for: {stage_name}")
                    break
            await asyncio.sleep(0.5)


async def send_file_to_bot(bot_token, chat_id, filepath, caption="", topic_id=None):
    try:
        async with aiohttp.ClientSession() as s:
            data = aiohttp.FormData()
            data.add_field('chat_id', str(chat_id))
            data.add_field('caption', caption)
            data.add_field('document', open(filepath, 'rb'))
            if topic_id:
                data.add_field('message_thread_id', str(topic_id))
            await s.post(f"https://api.telegram.org/bot{bot_token}/sendDocument", data=data)
            print(f"[BOT] File sent: {caption}")
    except Exception as e:
        print(f"[BOT] Error: {e}")


async def send_zip_to_bot(bot_token, chat_id, zip_path, caption="", topic_id=None):
    try:
        async with aiohttp.ClientSession() as s:
            data = aiohttp.FormData()
            data.add_field('chat_id', str(chat_id))
            data.add_field('caption', caption)
            data.add_field('document', open(zip_path, 'rb'))
            if topic_id:
                data.add_field('message_thread_id', str(topic_id))
            await s.post(f"https://api.telegram.org/bot{bot_token}/sendDocument", data=data)
            print(f"[BOT] ZIP sent: {caption}")
    except Exception as e:
        print(f"[BOT] ZIP error: {e}")


# ====================== BOT OPERATIONS ======================
async def clear_bot(client, bot):
    try:
        e = await client.get_entity(bot)
        await client.send_message(e, "/start")
        await asyncio.sleep(2)
        print(f"[BOT] /start sent to {bot}")
    except Exception as ex:
        print(f"[BOT] Error: {ex}")


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
                                print(f"[BOT] Clicked '{btn.text}'")
                                return True
            print(f"[BOT] Button '{text}' not found (attempt {attempt+1})")
        except Exception as ex:
            print(f"[BOT] Error: {ex}")
    return False


async def wait_xlsx(client, bot, timeout=180, since_msg_id=None):
    e = await client.get_entity(bot)
    start = time.time()
    print(f"[BOT] Waiting for XLSX from {bot}...")
    while time.time() - start < timeout:
        msgs = await client.get_messages(e, limit=5)
        for msg in msgs:
            if not msg or not msg.document:
                continue
            if since_msg_id is not None and msg.id <= since_msg_id:
                continue
            for a in msg.document.attributes:
                if isinstance(a, DocumentAttributeFilename) and a.file_name.endswith('.xlsx'):
                    print(f"[BOT] Received XLSX: {a.file_name}")
                    return msg
        await asyncio.sleep(3)
    print(f"[BOT] XLSX not received")
    return None


# ====================== PARSE XLSX ======================
def parse_xlsx(path):
    res = []
    try:
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        h = {}

        for col in range(1, ws.max_column + 1):
            v = str(ws.cell(row=1, column=col).value or "").upper().strip()
            if any(k in v for k in ['INN', 'PASSPORT']):
                continue
            if any(k in v for k in ['FIO', 'NAME']):
                h['fio'] = col
            if any(k in v for k in ['DATE', 'BIRTH']):
                h['date'] = col
            if any(k in v for k in ['PHONE', 'TEL']):
                h['phone'] = col
            elif 'NOMER' in v and 'PASSPORT' not in v:
                h['phone'] = col
            if any(k in v for k in ['SNILS']):
                h['snils'] = col
            if any(k in v for k in ['ADDR', 'АДРЕС', 'АДРЕСС']):
                h['addr'] = col

        for row in range(2, ws.max_row + 1):
            try:
                r = {}
                if 'fio' in h:
                    r['fio'] = normalize_fio_local(ws.cell(row=row, column=h['fio']).value)
                if 'date' in h:
                    r['date'] = parse_date(ws.cell(row=row, column=h['date']).value)
                if 'phone' in h:
                    r['phone'] = clean_phone(ws.cell(row=row, column=h['phone']).value)
                if 'snils' in h:
                    r['snils'] = clean_snils(ws.cell(row=row, column=h['snils']).value)
                if 'addr' in h:
                    r['addr'] = str(ws.cell(row=row, column=h['addr']).value or "").strip()
                if any(v for v in r.values() if v):
                    res.append(r)
            except Exception:
                pass
        return res
    except Exception as e:
        print(f"[PARSE] Error: {e}")
        return []


# ====================== FILL TABLE ======================
COL_NO = 1
COL_FIO = 2
COL_DATE = 3
COL_PHONE = 4
COL_SNILS = 5
COL_ADDR = 6


def fill_dates_from_response(ws, response_records):
    filled = 0
    print(f"[FILL-DATES] Starting. Responses: {len(response_records)}")
    
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
                print(f"[FILL-DATES] Row {row}: FILLED")
    
    print(f"[FILL-DATES] Total filled: {filled}")
    return filled


def fill_phones_from_response(ws, response_records):
    filled = 0
    print(f"[FILL-PHONES] Starting. Responses: {len(response_records)}")
    
    table_index = {}
    for row in range(2, ws.max_row + 1):
        fio = normalize_fio_local(str(ws.cell(row=row, column=COL_FIO).value or ""))
        date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        if fio and date and date != 'None':
            table_index[(fio, date)] = row
    
    for rec in response_records:
        rec_fio = normalize_fio_local(rec.get('fio', ''))
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
                print(f"[FILL-PHONES] Row {row}: FILLED")
        else:
            for row in range(2, ws.max_row + 1):
                table_fio = normalize_fio_local(str(ws.cell(row=row, column=COL_FIO).value or ""))
                if table_fio == rec_fio:
                    existing = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
                    if not existing or existing == 'None':
                        ws.cell(row=row, column=COL_PHONE).value = clean_rec
                        filled += 1
                        print(f"[FILL-PHONES] Row {row} (fallback): FILLED")
                    break
    
    print(f"[FILL-PHONES] Total filled: {filled}")
    return filled


def fill_snils_dates(ws, response_records):
    filled = 0
    print(f"[FILL-SNILS] Starting")
    
    table_snils = {}
    for row in range(2, ws.max_row + 1):
        snils = clean_snils(str(ws.cell(row=row, column=COL_SNILS).value or ""))
        if snils and len(snils) >= 11:
            table_snils[snils] = row
    
    for rec in response_records:
        rec_snils = clean_snils(rec.get('snils', ''))
        rec_fio = normalize_fio_local(rec.get('fio', ''))
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
                print(f"[FILL-SNILS] Row {row}: FILLED by SNILS")
                found = True
        
        if not found and rec_fio:
            for row in range(2, ws.max_row + 1):
                table_fio = normalize_fio_local(str(ws.cell(row=row, column=COL_FIO).value or ""))
                if table_fio == rec_fio:
                    existing = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
                    if not existing or existing == 'None':
                        ws.cell(row=row, column=COL_DATE).value = rec_date
                        filled += 1
                        print(f"[FILL-SNILS] Row {row} (fallback): FILLED")
                    found = True
                    break
    
    print(f"[FILL-SNILS] Total filled: {filled}")
    return filled


# ====================== DOBIV SAURON ======================
async def dobiv_sauron(client, bot, fio, date, account_id, row_num, ws, wb, result_file, add_log):
    try:
        norm_date = parse_date(date)
        if not norm_date:
            norm_date = date
        
        query = f"{fio} {norm_date}"
        add_log(f"[DOBIV] Acc {account_id}, row {row_num}: {query}")
        
        await client.send_message(bot, query)
        await asyncio.sleep(5)
        
        async for msg in client.iter_messages(bot, limit=10):
            if msg.text and ("REPORT" in msg.text or "PHONES" in msg.text):
                phones = extract_phones_from_text(msg.text)
                if phones:
                    add_log(f"[DOBIV] Found phones: {phones}")
                    return phones
                break
        
        return []
    except Exception as e:
        add_log(f"[DOBIV] Error: {e}")
        return []


# ====================== ПОИСК ИНН В ГРУППЕ ======================
async def find_inn_in_group(client, group_id, address, topic_id=None):
    """
    Ищет в группе/теме сообщение с адресом, извлекает ИНН
    Возвращает: (inn, found) или (None, False)
    """
    try:
        # Получаем сущность группы
        entity = await client.get_entity(int(group_id))
        
        # Ищем сообщения с адресом (ищем по частям адреса)
        search_parts = address.split(',')
        search_terms = []
        for part in search_parts:
            part = part.strip()
            if part and len(part) > 3:
                search_terms.append(part)
        
        # Ищем сообщения
        async for msg in client.iter_messages(entity, limit=100):
            if not msg.text:
                continue
            
            # Проверяем, есть ли адрес в сообщении
            msg_lower = msg.text.lower()
            address_lower = address.lower()
            
            # Ищем совпадение по ключевым частям
            match_count = 0
            for term in search_terms[:3]:  # берём первые 3 части
                if term.lower() in msg_lower:
                    match_count += 1
            
            # Если совпало хотя бы 2 части - считаем, что это оно
            if match_count >= 2:
                # Ищем ИНН
                inn = extract_inn_from_text(msg.text)
                if inn:
                    return inn, True
                else:
                    return None, True  # сообщение найдено, но ИНН нет
        
        return None, False
    except Exception as e:
        print(f"[GROUP] Ошибка поиска: {e}")
        return None, False


# ====================== ПЕРЕИМЕНОВАНИЕ ФАЙЛОВ ======================
async def rename_files_by_address(ws, client, group_id, topic_id, add_log, tables_names):
    """
    Переименовывает файлы по адресу из колонки "Адресс"
    Ищет ИНН в группе/теме
    Имя файла: ИНН_Адрес.xlsx или УК_Адрес.xlsx
    """
    if not group_id:
        add_log("[RENAME] Группа не указана - пропускаю переименование")
        return tables_names
    
    add_log("[RENAME] Начинаю переименование файлов по адресам...")
    
    # Собираем уникальные адреса из таблицы
    address_map = {}  # {table_num: address}
    table_nums = {}
    
    for row in range(2, ws.max_row + 1):
        table_num = str(ws.cell(row=row, column=COL_NO).value or "").strip()
        addr = str(ws.cell(row=row, column=COL_ADDR).value or "").strip()
        
        if table_num and addr and addr != 'None':
            if table_num not in address_map:
                address_map[table_num] = addr
                table_nums[table_num] = row
    
    add_log(f"[RENAME] Найдено {len(address_map)} уникальных адресов")
    
    # Для каждого адреса ищем ИНН в группе
    new_names = {}
    for table_num, addr in address_map.items():
        # Нормализуем адрес для поиска (убираем "кв.")
        clean_addr = re.sub(r'\s*кв\.?\s*\d+', '', addr).strip()
        clean_addr = re.sub(r',\s*,', ',', clean_addr)
        
        add_log(f"[RENAME] Ищу ИНН для: {clean_addr}")
        
        inn, found = await find_inn_in_group(client, group_id, clean_addr, topic_id)
        
        if found and inn:
            new_name = f"{inn}_{clean_addr}"
            add_log(f"[RENAME] Найден ИНН: {inn} -> {new_name}")
        else:
            new_name = f"УК_{clean_addr}"
            add_log(f"[RENAME] ИНН не найден -> {new_name}")
        
        # Очищаем имя файла от недопустимых символов
        new_name = re.sub(r'[<>:"/\\|?*]', '_', new_name)
        new_names[table_num] = new_name
    
    # Обновляем имена таблиц
    updated_names = []
    for i, name in enumerate(tables_names):
        table_num = str(i + 1)
        if table_num in new_names:
            updated_names.append(new_names[table_num])
        else:
            updated_names.append(name)
    
    add_log(f"[RENAME] Переименовано {len(new_names)} файлов")
    return updated_names


# ====================== FULL CYCLE (7 STAGES) ======================
async def run_full_cycle(ss, bot1, bot2, bot_token, chat_id,
                         items_no_date, items_no_phone, items_snils,
                         year_range, original_rows, tables_names=None, topic_id=None, group_id=None):
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
                await send_file_to_bot(bot_token, chat_id, file_path, f"Table after: {stage_label}", topic_id)
            else:
                wb.save(result_file)
                await send_file_to_bot(bot_token, chat_id, result_file, f"Table after: {stage_label}", topic_id)

    async def send_final_zip():
        if not bot_token or not chat_id:
            return
        
        split_result = split_by_table_num(
            [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)],
            [[ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)] for r in range(2, ws.max_row + 1)]
        )
        
        if not split_result:
            await send_file_to_bot(bot_token, chat_id, result_file, "FINAL FILE (all stages)", topic_id)
            return
        
        # Переименовываем файлы
        final_names = await rename_files_by_address(ws, client, group_id, topic_id, add, tables_names)
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            phones_all = set()
            
            for table_num, table_data in split_result.items():
                try:
                    geo_result = geo_filter(table_data['headers'], table_data['rows'])
                    phones_all.update(geo_result['phones'])
                    
                    ws_data = [geo_result['headers']] + geo_result['rows']
                    wb_temp = Workbook()
                    ws_temp = wb_temp.active
                    for row in ws_data:
                        ws_temp.append(row)
                    
                    xlsx_buffer = io.BytesIO()
                    wb_temp.save(xlsx_buffer)
                    xlsx_buffer.seek(0)
                    
                    # Используем переименованное имя
                    idx = int(table_num) - 1
                    if idx < len(final_names):
                        name = f"{final_names[idx]}.xlsx"
                    else:
                        name = f"GEO_{table_num}.xlsx"
                    zf.writestr(name, xlsx_buffer.getvalue())
                except Exception as e:
                    add(f"[ZIP] Table {table_num} error: {e}")
                    continue
            
            if phones_all:
                zf.writestr('numbers.txt', '\n'.join(sorted(phones_all)))
            else:
                zf.writestr('numbers.txt', '(no valid phones)')
        
        zip_buffer.seek(0)
        zip_path = os.path.join(TEMP_DIR, f"result_{int(time.time())}.zip")
        with open(zip_path, 'wb') as f:
            f.write(zip_buffer.getvalue())
        
        await send_zip_to_bot(bot_token, chat_id, zip_path, "FINAL ZIP ARCHIVE (all tables + numbers.txt)", topic_id)
        add("[v] ZIP archive sent")

    # === CREATE TABLE ===
    result_file = os.path.join(TEMP_DIR, f"result_{int(time.time())}.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.append(["N tablicy", "FIO", "Data", "Nomer", "SNILS", "Adress"])

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
    add(f"Table created: {ws.max_row - 1} rows")

    # === ПАКЕТНАЯ НОРМАЛИЗАЦИЯ ФИО ===
    add("[DEEPSEEK] Normalizing FIO in batches...")
    fio_rows = []
    fio_values = []
    for row in range(2, ws.max_row + 1):
        fio_val = str(ws.cell(row=row, column=COL_FIO).value or "").strip()
        if fio_val and fio_val != 'None':
            fio_rows.append(row)
            fio_values.append(fio_val)
    
    batch_size = 40
    for batch_idx in range(0, len(fio_values), batch_size):
        batch = fio_values[batch_idx:batch_idx + batch_size]
        batch_rows = fio_rows[batch_idx:batch_idx + batch_size]
        
        normalized = await normalize_batch_deepseek(batch, 'fio')
        for j, norm_val in enumerate(normalized):
            if j < len(batch_rows):
                ws.cell(row=batch_rows[j], column=COL_FIO).value = norm_val
        
        add(f"[DEEPSEEK] FIO batch {batch_idx//batch_size + 1}/{(len(fio_values)+batch_size-1)//batch_size}: {len(batch)} names")
        await asyncio.sleep(0.3)
    wb.save(result_file)
    add("[DEEPSEEK] FIO normalization complete")

    # === ПАКЕТНАЯ НОРМАЛИЗАЦИЯ АДРЕСОВ ===
    add("[DEEPSEEK] Normalizing addresses in batches...")
    addr_rows = []
    addr_values = []
    for row in range(2, ws.max_row + 1):
        addr_val = str(ws.cell(row=row, column=COL_ADDR).value or "").strip()
        if addr_val and addr_val != 'None' and len(addr_val) > 5:
            addr_rows.append(row)
            addr_values.append(addr_val)
    
    for batch_idx in range(0, len(addr_values), batch_size):
        batch = addr_values[batch_idx:batch_idx + batch_size]
        batch_rows = addr_rows[batch_idx:batch_idx + batch_size]
        
        normalized = await normalize_batch_deepseek(batch, 'address')
        for j, norm_val in enumerate(normalized):
            if j < len(batch_rows):
                ws.cell(row=batch_rows[j], column=COL_ADDR).value = norm_val
        
        add(f"[DEEPSEEK] Address batch {batch_idx//batch_size + 1}/{(len(addr_values)+batch_size-1)//batch_size}: {len(batch)} addresses")
        await asyncio.sleep(0.3)
    wb.save(result_file)
    add("[DEEPSEEK] Address normalization complete")

    # === ANALYZE ===
    real_snils = []
    for row in range(2, ws.max_row + 1):
        existing_date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        snils_val = str(ws.cell(row=row, column=COL_SNILS).value or "").strip()
        if (not existing_date or existing_date == 'None') and snils_val and len(clean_snils(snils_val)) >= 11:
            real_snils.append(clean_snils(snils_val))
    real_snils = list(set(real_snils))

    add(f"To probe: dates={len(items_no_date)} phones={len(items_no_phone)} snils={len(real_snils)}")

    # ============ STAGE 1: FIO+PHONE -> BOT1 ============
    if items_no_date and not stop_requested:
        cid = f"s1_{int(time.time())}"
        add(f"STAGE 1: FIO+PHONE -> bot1 ({len(items_no_date)} rows)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "STAGE 1: FIO+PHONE", len(items_no_date), cid, add, topic_id)
        if result == "stop":
            await send_final_zip()
            return {"ok": True, "log": log, "stopped": True}
        if result == "skip":
            add("[v] STAGE 1 SKIPPED")
        elif result == "confirm":
            add("[v] CONFIRMED! Starting stage 1...")

            txt = "\n".join([f"{it.get('fio','')}\t{it.get('phone','')}" for it in items_no_date])
            tpath = os.path.join(TEMP_DIR, f"t1_{int(time.time())}.txt")
            with open(tpath, 'w', encoding='utf-8') as f:
                f.write(txt)

            await clear_bot(client, bot1)
            e = await client.get_entity(bot1)
            await client.send_message(e, "Probev")
            await asyncio.sleep(2)
            await click_btn(client, bot1, "FIO+phone")
            await asyncio.sleep(2)
            
            last_msgs = await client.get_messages(e, limit=1)
            last_msg_id = last_msgs[0].id if last_msgs else 0
            await client.send_file(e, tpath)
            add("File sent to bot1, waiting...")

            msg = await wait_xlsx(client, bot1, 300, since_msg_id=last_msg_id)
            if msg:
                rpath = os.path.join(TEMP_DIR, f"r1_{int(time.time())}.xlsx")
                await client.download_media(msg, file=rpath)
                recs = parse_xlsx(rpath)
                add(f"Received responses: {len(recs)}")
                fill_dates_from_response(ws, recs)
                wb.save(result_file)
                await send_status("stage 1")
            else:
                add("[!] Bot1 did not respond")
    
    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # === RECALCULATE SNILS ===
    real_snils = []
    for row in range(2, ws.max_row + 1):
        existing_date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        snils_val = str(ws.cell(row=row, column=COL_SNILS).value or "").strip()
        if (not existing_date or existing_date == 'None') and snils_val and len(clean_snils(snils_val)) >= 11:
            real_snils.append(clean_snils(snils_val))
    real_snils = list(set(real_snils))
    add(f"[RECALC] SNILS with empty date: {len(real_snils)}")

    # ============ STAGE 2: SNILS -> BOT1 ============
    if real_snils and not stop_requested:
        cid = f"s2_{int(time.time())}"
        add(f"STAGE 2: SNILS -> bot1 ({len(real_snils)} snils)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "STAGE 2: SNILS (bot1)", len(real_snils), cid, add, topic_id)
        if result == "stop":
            await send_final_zip()
            return {"ok": True, "log": log, "stopped": True}
        if result == "skip":
            add("[v] STAGE 2 SKIPPED")
        elif result == "confirm":
            add("[v] CONFIRMED! Starting stage 2...")

            txt = "\n".join(real_snils)
            tpath = os.path.join(TEMP_DIR, f"t2_{int(time.time())}.txt")
            with open(tpath, 'w', encoding='utf-8') as f:
                f.write(txt)

            await clear_bot(client, bot1)
            e = await client.get_entity(bot1)
            await client.send_message(e, "Probev")
            await asyncio.sleep(4)
            await click_btn(client, bot1, "SNILS")
            await asyncio.sleep(3)
            
            last_msgs = await client.get_messages(e, limit=1)
            last_msg_id = last_msgs[0].id if last_msgs else 0
            await client.send_file(e, tpath)
            add("SNILS sent to bot1, waiting...")

            msg = await wait_xlsx(client, bot1, 300, since_msg_id=last_msg_id)
            if msg:
                rpath = os.path.join(TEMP_DIR, f"r2_{int(time.time())}.xlsx")
                await client.download_media(msg, file=rpath)
                recs = parse_xlsx(rpath)
                add(f"Received via SNILS: {len(recs)}")
                fill_snils_dates(ws, recs)
                wb.save(result_file)
                await send_status("stage 2")
            else:
                add("[!] Bot1 did not respond")
    
    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # ============ STAGE 3: SNILS -> BOT2 ============
    snils_still_empty = []
    for row in range(2, ws.max_row + 1):
        existing_date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        snils_val = str(ws.cell(row=row, column=COL_SNILS).value or "").strip()
        if (not existing_date or existing_date == 'None') and snils_val and len(clean_snils(snils_val)) >= 11:
            snils_still_empty.append(clean_snils(snils_val))
    snils_still_empty = list(set(snils_still_empty))

    if snils_still_empty and not stop_requested:
        cid = f"s3_{int(time.time())}"
        add(f"STAGE 3: SNILS -> bot2 ({len(snils_still_empty)} snils)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "STAGE 3: SNILS (bot2)", len(snils_still_empty), cid, add, topic_id)
        if result == "stop":
            await send_final_zip()
            return {"ok": True, "log": log, "stopped": True}
        if result == "skip":
            add("[v] STAGE 3 SKIPPED")
        elif result == "confirm":
            add("[v] CONFIRMED! Starting stage 3...")

            txt = "\n".join(snils_still_empty)
            tpath = os.path.join(TEMP_DIR, f"t3_{int(time.time())}.txt")
            with open(tpath, 'w', encoding='utf-8') as f:
                f.write(txt)

            e = await client.get_entity(bot2)
            last_msgs = await client.get_messages(e, limit=1)
            last_msg_id = last_msgs[0].id if last_msgs else 0
            await client.send_file(e, tpath)
            add("SNILS sent to bot2, waiting...")

            msg = await wait_xlsx(client, bot2, 300, since_msg_id=last_msg_id)
            if msg:
                rpath = os.path.join(TEMP_DIR, f"r3_{int(time.time())}.xlsx")
                await client.download_media(msg, file=rpath)
                recs = parse_xlsx(rpath)
                add(f"Received via SNILS bot2: {len(recs)}")
                fill_snils_dates(ws, recs)
                wb.save(result_file)
                await send_status("stage 3")
            else:
                add("[!] Bot2 did not respond")
    
    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # ============ STAGE 4: YEAR FILTER ============
    if year_range and not stop_requested:
        try:
            parts = year_range.split('-')
            yf, yt = int(parts[0]), int(parts[1])
            add(f"STAGE 4: Year filter {yf}-{yt}")

            cid = f"s4_{int(time.time())}"
            result = await safe_confirm_with_buttons(bot_token, chat_id, f"STAGE 4: Filter {yf}-{yt}", ws.max_row - 1, cid, add, topic_id)
            if result == "stop":
                await send_final_zip()
                return {"ok": True, "log": log, "stopped": True}
            if result == "skip":
                add("[v] STAGE 4 SKIPPED")
            elif result == "confirm":
                add("[v] CONFIRMED! Filtering...")
                
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
                await send_status(f"stage 4 (filter {yf}-{yt})")
                add(f"Deleted: {len(rows_to_delete)}, remaining: {ws.max_row - 1}")
        except Exception as e:
            add(f"Filter error: {e}")
    
    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # ============ STAGE 5: FIO+DATE -> BOT1 ============
    items_no_phone_after_stage4 = []
    for row in range(2, ws.max_row + 1):
        fio_val = str(ws.cell(row=row, column=COL_FIO).value or "").strip()
        date_val = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        phone_val = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
        if (fio_val and fio_val != 'None'
                and date_val and date_val != 'None' and date_val != '0'
                and (not phone_val or phone_val == 'None' or phone_val == '0')):
            items_no_phone_after_stage4.append({
                'fio': normalize_fio_local(fio_val),
                'date': date_val
            })

    if items_no_phone_after_stage4 and not stop_requested:
        cid = f"s5_{int(time.time())}"
        add(f"STAGE 5: FIO+DATE -> bot1 ({len(items_no_phone_after_stage4)} rows)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "STAGE 5: FIO+DATE (bot1)", len(items_no_phone_after_stage4), cid, add, topic_id)
        if result == "stop":
            await send_final_zip()
            return {"ok": True, "log": log, "stopped": True}
        if result == "skip":
            add("[v] STAGE 5 SKIPPED")
        elif result == "confirm":
            add("[v] CONFIRMED! Starting stage 5...")

            txt = "\n".join([f"{it.get('fio','')}\t{it.get('date','')}" for it in items_no_phone_after_stage4])
            tpath = os.path.join(TEMP_DIR, f"t5_{int(time.time())}.txt")
            with open(tpath, 'w', encoding='utf-8') as f:
                f.write(txt)

            await clear_bot(client, bot1)
            e = await client.get_entity(bot1)
            await client.send_message(e, "Probev")
            await asyncio.sleep(2)
            await click_btn(client, bot1, "FIO+date")
            await asyncio.sleep(2)
            
            last_msgs = await client.get_messages(e, limit=1)
            last_msg_id = last_msgs[0].id if last_msgs else 0
            await client.send_file(e, tpath)
            add("File sent to bot1, waiting...")

            msg = await wait_xlsx(client, bot1, 300, since_msg_id=last_msg_id)
            if msg:
                rpath = os.path.join(TEMP_DIR, f"r5_{int(time.time())}.xlsx")
                await client.download_media(msg, file=rpath)
                recs = parse_xlsx(rpath)
                add(f"Received responses: {len(recs)}")
                fill_phones_from_response(ws, recs)
                wb.save(result_file)
                await send_status("stage 5")
            else:
                add("[!] Bot1 did not respond")
    
    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # ============ STAGE 6: FIO+DATE -> BOT2 ============
    items_no_phone_after_stage5 = []
    for row in range(2, ws.max_row + 1):
        fio_val = str(ws.cell(row=row, column=COL_FIO).value or "").strip()
        date_val = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        phone_val = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
        if (fio_val and fio_val != 'None'
                and date_val and date_val != 'None' and date_val != '0'
                and (not phone_val or phone_val == 'None' or phone_val == '0')):
            items_no_phone_after_stage5.append({
                'fio': normalize_fio_local(fio_val),
                'date': date_val
            })

    if items_no_phone_after_stage5 and not stop_requested:
        cid = f"s6_{int(time.time())}"
        add(f"STAGE 6: FIO+DATE -> bot2 ({len(items_no_phone_after_stage5)} rows)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "STAGE 6: FIO+DATE (bot2)", len(items_no_phone_after_stage5), cid, add, topic_id)
        if result == "stop":
            await send_final_zip()
            return {"ok": True, "log": log, "stopped": True}
        if result == "skip":
            add("[v] STAGE 6 SKIPPED")
        elif result == "confirm":
            add("[v] CONFIRMED! Starting stage 6...")

            txt = "\n".join([f"{it.get('fio','')}\t{it.get('date','')}" for it in items_no_phone_after_stage5])
            tpath = os.path.join(TEMP_DIR, f"t6_{int(time.time())}.txt")
            with open(tpath, 'w', encoding='utf-8') as f:
                f.write(txt)

            e = await client.get_entity(bot2)
            last_msgs = await client.get_messages(e, limit=1)
            last_msg_id = last_msgs[0].id if last_msgs else 0
            await client.send_file(e, tpath)
            add("File sent to bot2, waiting...")

            msg = await wait_xlsx(client, bot2, 300, since_msg_id=last_msg_id)
            if msg:
                rpath = os.path.join(TEMP_DIR, f"r6_{int(time.time())}.xlsx")
                await client.download_media(msg, file=rpath)
                recs = parse_xlsx(rpath)
                add(f"Received responses from bot2: {len(recs)}")
                fill_phones_from_response(ws, recs)
                wb.save(result_file)
                await send_status("stage 6")
            else:
                add("[!] Bot2 did not respond")
    
    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # ============ STAGE 7: DOBIV -> BOT2 (SAURON) ============
    items_no_phone_final = []
    for row in range(2, ws.max_row + 1):
        fio_val = str(ws.cell(row=row, column=COL_FIO).value or "").strip()
        date_val = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        phone_val = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
        if (fio_val and fio_val != 'None'
                and date_val and date_val != 'None' and date_val != '0'
                and (not phone_val or phone_val == 'None' or phone_val == '0')):
            items_no_phone_final.append({
                'fio': normalize_fio_local(fio_val),
                'date': date_val
            })

    if items_no_phone_final and not stop_requested:
        cid = f"s7_{int(time.time())}"
        add(f"STAGE 7: DOBIV -> bot2 ({len(items_no_phone_final)} rows)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "STAGE 7: DOBIV (bot2)", len(items_no_phone_final), cid, add, topic_id)
        if result == "stop":
            await send_final_zip()
            return {"ok": True, "log": log, "stopped": True}
        if result == "skip":
            add("[v] STAGE 7 SKIPPED")
        elif result == "confirm":
            add("[v] CONFIRMED! Dobiv via sauron...")

            e = await client.get_entity(bot2)
            for i, it in enumerate(items_no_phone_final):
                if stop_requested:
                    break
                    
                try:
                    fio = it.get('fio', '')
                    date = it.get('date', '')
                    
                    phones = await dobiv_sauron(client, e, fio, date, "1", i+1, ws, wb, result_file, add)
                    
                    if phones:
                        for row in range(2, ws.max_row + 1):
                            table_fio = normalize_fio_local(str(ws.cell(row=row, column=COL_FIO).value or ""))
                            table_date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
                            existing_phone = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()

                            if table_fio == normalize_fio_local(fio) and table_date == date:
                                if not existing_phone or existing_phone == 'None':
                                    ws.cell(row=row, column=COL_PHONE).value = phones[0]
                                    add(f"[DOBIV] {fio} -> {phones[0]}")
                                break

                    if (i + 1) % 5 == 0:
                        wb.save(result_file)
                        add(f"[DOBIV] Saved after {i+1} rows")
                        await asyncio.sleep(2)

                except FloodWaitError as e:
                    add(f"[DOBIV] FloodWait: {e.seconds}s")
                    await asyncio.sleep(e.seconds)
                except Exception as e:
                    add(f"[DOBIV] Error: {e}")

            wb.save(result_file)
            await send_status("stage 7 (dobiv)")
            add("Dobiv completed")

    # === DELETE ROWS WITHOUT DATE ===
    rows_no_date = []
    for row in range(2, ws.max_row + 1):
        date_val = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        if not date_val or date_val == 'None' or date_val == '0':
            rows_no_date.append(row)
    
    if rows_no_date:
        add(f"[CLEANUP] Deleting {len(rows_no_date)} rows without date")
        for row in reversed(rows_no_date):
            ws.delete_rows(row)
        wb.save(result_file)
        add(f"Remaining: {ws.max_row - 1} rows")

    # === FINAL ===
    add("=== ALL STAGES COMPLETED ===")
    
    await send_final_zip()
    
    total_dates = 0
    total_phones = 0
    for row in range(2, ws.max_row + 1):
        if str(ws.cell(row=row, column=COL_DATE).value or "").strip():
            total_dates += 1
        if str(ws.cell(row=row, column=COL_PHONE).value or "").strip():
            total_phones += 1
    add(f"Total: rows={ws.max_row-1}, with dates={total_dates}, with phones={total_phones}")

    return {"ok": True, "log": log, "stopped": stop_requested}


# ====================== ENDPOINTS ======================
async def handle_health(request):
    return web.json_response({"ok": True, "message": "X Backend v14.0"})


async def handle_root(request):
    try:
        html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "сайт.html")
        with open(html_path, "r", encoding="utf-8") as f:
            html = f.read()
        return web.Response(text=html, content_type="text/html")
    except Exception as e:
        print(f"[ROOT] Error: {e}")
        return web.Response(text="OK", content_type="text/plain")


async def handle_upload_zip(request):
    try:
        reader = await request.multipart()
        field = await reader.next()
        if field.name != 'file':
            return web.json_response({"ok": False, "error": "No file"}, status=400)
        
        data = await field.read()
        zip_path = os.path.join(TEMP_DIR, f"upload_{int(time.time())}.zip")
        with open(zip_path, 'wb') as f:
            f.write(data)
        
        tables_data = []
        with zipfile.ZipFile(zip_path, 'r') as zf:
            for file_info in zf.filelist:
                if file_info.filename.endswith('.xlsx'):
                    with zf.open(file_info) as f:
                        xlsx_data = f.read()
                        xlsx_path = os.path.join(TEMP_DIR, f"tmp_{int(time.time())}_{file_info.filename}")
                        with open(xlsx_path, 'wb') as out:
                            out.write(xlsx_data)
                        
                        wb = load_workbook(xlsx_path, data_only=True)
                        ws = wb.active
                        headers = [str(cell.value or "").strip() for cell in ws[1]]
                        rows = []
                        for row in ws.iter_rows(min_row=2, values_only=True):
                            if any(cell for cell in row):
                                rows.append([str(cell or "").strip() for cell in row])
                        tables_data.append({'headers': headers, 'rows': rows, 'name': file_info.filename})
                        os.remove(xlsx_path)
        
        merged = merge_tables(tables_data)
        if not merged:
            return web.json_response({"ok": False, "error": "No tables found"}, status=400)
        
        merged_path = os.path.join(TEMP_DIR, f"merged_{int(time.time())}.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.append(merged['headers'])
        for row in merged['rows']:
            ws.append(row)
        wb.save(merged_path)
        
        names = [t.get('name', f'Table_{i+1}') for i, t in enumerate(tables_data)]
        
        return web.json_response({
            "ok": True,
            "file": merged_path,
            "headers": merged['headers'],
            "rows": merged['rows'],
            "count": len(merged['rows']),
            "tables_count": len(tables_data),
            "names": names
        })
    except Exception as e:
        print(f"[UPLOAD] Error: {e}")
        traceback.print_exc()
        return web.json_response({"ok": False, "error": str(e)}, status=500)


async def handle_send_code(request):
    try:
        d = await request.json()
        phone = d.get("phone", "").strip()
        if not phone:
            return web.json_response({"ok": False, "error": "Enter phone"}, status=400)

        c = TelegramClient(StringSession(), API_ID, API_HASH)
        await c.connect()
        r = await c.send_code_request(phone)
        sessions[phone] = {"client": c, "hash": r.phone_code_hash}
        print(f"[AUTH] Code sent to {phone}")
        return web.json_response({"ok": True})
    except Exception as e:
        print(f"[AUTH] Error: {e}")
        return web.json_response({"ok": False, "error": str(e)}, status=400)


async def handle_verify_code(request):
    try:
        d = await request.json()
        phone = d.get("phone", "").strip()
        code = d.get("code", "").strip()
        password = d.get("password", "").strip()

        if phone not in sessions:
            return web.json_response({"ok": False, "error": "Send code first"}, status=400)

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
        print(f"[AUTH] Login: {me.first_name}")
        return web.json_response({
            "ok": True,
            "session": ss,
            "phone": phone,
            "name": me.first_name or ""
        })
    except Exception as e:
        print(f"[AUTH] Error: {e}")
        return web.json_response({"ok": False, "error": str(e)}, status=400)


async def handle_full_probev(request):
    try:
        d = await request.json()
        ss = d.get("session", "")
        bot1 = d.get("bot1", "@osint_pam_pam_bot")
        bot2 = d.get("bot2", "@proverim123_bot")
        bot_token = d.get("bot_token", "")
        chat_id = d.get("chat_id", "")
        topic_id = d.get("topic_id", None)
        group_id = d.get("group_id", None)  # ID группы для поиска ИНН
        year_range = d.get("year_range", "") or "1945-1975"
        items_no_date = d.get("items_no_date", [])
        items_no_phone = d.get("items_no_phone", [])
        items_snils = d.get("items_snils", [])
        original_rows = d.get("original_rows", [])
        tables_names = d.get("tables_names", [])

        if not ss:
            return web.json_response({"ok": False, "error": "No session"}, status=400)

        async with probev_lock:
            if ss in active_probevs:
                task = active_probevs[ss]
                if not task.done():
                    return web.json_response({"ok": False, "error": "Probev already running"}, status=409)
                else:
                    del active_probevs[ss]

        print(f"\n{'='*60}")
        print(f"[PROBEV] START FULL CYCLE (7 STAGES + DeepSeek + RENAME)")
        print(f"[PROBEV] Bot1: {bot1}, Bot2: {bot2}")
        print(f"[PROBEV] Rows without date: {len(items_no_date)}")
        print(f"[PROBEV] Rows without phone: {len(items_no_phone)}")
        print(f"[PROBEV] Group ID: {group_id or 'None'}")
        print(f"[PROBEV] Topic ID: {topic_id or 'None'}")
        print(f"{'='*60}\n")

        async def run_and_cleanup():
            try:
                result = await run_full_cycle(
                    ss, bot1, bot2, bot_token, chat_id,
                    items_no_date, items_no_phone, items_snils,
                    year_range, original_rows, tables_names, topic_id, group_id
                )
                return result
            finally:
                async with probev_lock:
                    if ss in active_probevs:
                        del active_probevs[ss]
                print(f"[PROBEV] Cleanup completed for session {ss[:10]}...")

        task = asyncio.create_task(run_and_cleanup())
        async with probev_lock:
            active_probevs[ss] = task

        result = await task
        return web.json_response(result)

    except Exception as e:
        traceback.print_exc()
        async with probev_lock:
            if ss in active_probevs:
                del active_probevs[ss]
        return web.json_response({"ok": False, "error": str(e)}, status=400)


async def handle_stop(request):
    global stop_requested
    stop_requested = True
    return web.json_response({"ok": True, "message": "Stop requested"})


# ====================== STARTUP ======================
app = web.Application(middlewares=[log_and_cors], client_max_size=200 * 1024 * 1024)
app.router.add_get("/", handle_root)
app.router.add_get("/health", handle_health)
app.router.add_post("/upload-zip", handle_upload_zip)
app.router.add_post("/send-code", handle_send_code)
app.router.add_post("/verify-code", handle_verify_code)
app.router.add_post("/full-probev", handle_full_probev)
app.router.add_post("/stop", handle_stop)


async def on_startup(app):
    port = app["port"]
    msg = (
        "=" * 60 + "\n"
        f"X Backend v14.0 STARTED (DeepSeek BATCH + ПЕРЕИМЕНОВАНИЕ ПО АДРЕСУ)\n"
        f"Host: 0.0.0.0  |  Port: {port}\n"
        + "=" * 60
    )
    print(msg, flush=True)


async def on_shutdown(app):
    print("[SHUTDOWN] Closing connections...", flush=True)
    for ss, client in list(user_clients.items()):
        try:
            if client.is_connected():
                await client.disconnect()
        except Exception:
            pass
    user_clients.clear()
    sessions.clear()
    pending_confirms.clear()
    print("[SHUTDOWN] Done.", flush=True)


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

        print(f"[START] SERVER READY — port {port}", flush=True)

        def shutdown():
            print("[SIGNAL] Shutting down...", flush=True)
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