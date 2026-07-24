# server.py - X Backend v17.0 (ГЛУБОКАЯ ИНТЕГРАЦИЯ: DeepSeek Квартиры + Сессии + Добив)
# Установка: pip install aiohttp telethon openpyxl
# Запуск: python server.py
# Порт: 4545

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

# Резервный API ключ (можно добавить OpenAI/Anthropic при необходимости)
BACKUP_API_KEY = None  # "sk-...备用"

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
def capitalize_words(text):
    """Каждое слово с прописной буквы (кроме заголовков)"""
    if not text:
        return text
    words = str(text).split()
    return ' '.join([w.capitalize() for w in words])


def normalize_fio_local(raw):
    """Приводит ФИО к формату 'Каждое Слово С Прописной'"""
    if not raw:
        return ""
    cleaned = re.sub(r'[^A-Za-zА-ЯЁа-яё\-]', ' ', str(raw))
    words = [w.strip() for w in cleaned.split() if w.strip()]
    if not words:
        return ""
    # Каждое слово с прописной, остальные строчные
    return ' '.join([w[0].upper() + w[1:].lower() if w else '' for w in words])


def normalize_address_local(raw):
    """Приводит адрес к формату 'Город, Улица, Дом, Квартира' с прописной буквы"""
    if not raw:
        return ""
    s = str(raw).strip().lower()
    # Удаляем лишние префиксы: область, город, улица и т.д.
    s = re.sub(r'\b(область|обл|край|республика|ао|район|р-н)\b\.?\s*', '', s)
    s = re.sub(r'\b(г|город|гор)\b\.?\s*', '', s)
    s = re.sub(r'\b(ул|улица|пр-т|проспект|пер|переулок|пр|проезд|б-р|бульвар|пл|площадь|наб|набережная|ш|шоссе)\b\.?\s*', '', s)
    s = re.sub(r'\b(д|дом|влд|владение)\b\.?\s*', '', s)
    s = re.sub(r'\b(кв|квартира|кв-ра)\b\.?\s*', '', s)
    s = re.sub(r'\b(корп|корпус|стр|строение)\b\.?\s*', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    s = re.sub(r'\s*,\s*', ', ', s)
    # Каждое слово с прописной буквы
    parts = []
    for p in s.split(','):
        p = p.strip()
        if p:
            words = p.split()
            cap_words = [w[0].upper() + w[1:].lower() if w else '' for w in words]
            parts.append(' '.join(cap_words))
    return ', '.join(parts)


async def normalize_batch_deepseek(items, prompt_type='fio', retry_count=0):
    """Пакетная нормализация через DeepSeek с резервной моделью"""
    if not items:
        return []
    
    if len(items) <= 2:
        if prompt_type == 'fio':
            return [normalize_fio_local(f) for f in items]
        else:
            return [normalize_address_local(a) for a in items]
    
    # Пробуем DeepSeek
    try:
        async with aiohttp.ClientSession() as session:
            url = "https://api.deepseek.com/v1/chat/completions"
            headers = {
                "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
                "Content-Type": "application/json"
            }
            
            if prompt_type == 'fio':
                system_prompt = "Приведи каждое ФИО к формату: Фамилия Имя Отчество (каждое слово с прописной буквы, остальные строчные). Верни только список, по одному на строку, с номерами."
                items_text = "\n".join([f"{i+1}. {f}" for i, f in enumerate(items)])
            elif prompt_type == 'address':
                system_prompt = "Приведи каждый адрес к формату: Город, Улица, Дом, Квартира. Убери 'область', 'город', 'улица', 'дом', 'квартира'. Пример: 'Мурманск, Старостина, 69, 112'. Каждое слово с прописной буквы. Верни только список, по одному на строку, с номерами."
                items_text = "\n".join([f"{i+1}. {a}" for i, a in enumerate(items)])
            elif prompt_type == 'find_apartment':
                system_prompt = "Для каждого адреса без квартиры найди квартиру из примеров. Сопоставляй по городу, улице и дому. Верни адрес с квартирой в формате: Город, Улица, Дом, Квартира. Каждое слово с прописной буквы. Верни только список, по одному на строку, с номерами. Если квартира не найдена, оставь адрес как есть."
                items_text = "\n".join([f"{i+1}. {a}" for i, a in enumerate(items)])
            else:
                return items
            
            payload = {
                "model": "deepseek-chat",
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"Нормализуй:\n{items_text}"}
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
                            # Приводим к формату "Каждое Слово С Прописной"
                            result.append(capitalize_words(line))
                        else:
                            result.append(capitalize_words(line))
                    return result
    except Exception as e:
        print(f"[DEEPSEEK] Ошибка: {e}")
        
        # Если есть резервный ключ и это первая попытка
        if BACKUP_API_KEY and retry_count == 0:
            print("[DEEPSEEK] Переключение на резервную модель...")
            # Здесь можно добавить OpenAI/Anthropic
            pass
    
    # Fallback: локальная нормализация
    if prompt_type == 'fio':
        return [normalize_fio_local(f) for f in items]
    elif prompt_type == 'address' or prompt_type == 'find_apartment':
        return [normalize_address_local(a) for a in items]
    return items


def clean_phone(text):
    if not text:
        return ""
    d = re.sub(r'[^0-9]', '', str(text))
    if len(d) >= 10:
        return '+7' + d[-10:]
    return d


def clean_phone_without_plus(text):
    """Очищает телефон и возвращает без +7 (только цифры)"""
    if not text:
        return ""
    d = re.sub(r'[^0-9]', '', str(text))
    if len(d) >= 10:
        return '7' + d[-10:]
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
    inn_match = re.search(r'\b\d{12}\b', text)
    if inn_match:
        return inn_match.group()
    return None


def extract_address_from_report(text):
    """Извлекает адрес из отчёта саурона"""
    patterns = [
        r'Адрес[:\s]+([^\n]+)',
        r'Контактный адрес[:\s]+([^\n]+)',
        r'Адрес регистрации[:\s]+([^\n]+)',
        r'АДРЕС[:\s]+([^\n]+)',
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return None


def extract_apartment_from_report(text, target_address):
    """Извлекает квартиру из отчёта по адресу"""
    # Ищем все адреса в отчёте
    address_pattern = r'(?:Адрес|Контактный адрес|Адрес регистрации|АДРЕС)[:\s]+([^\n]+)'
    matches = re.findall(address_pattern, text, re.IGNORECASE)
    
    # Нормализуем целевой адрес (без квартиры)
    target_clean = re.sub(r',?\s*кв\.?\s*\d+', '', target_address).strip().lower()
    
    for addr in matches:
        addr_clean = re.sub(r',?\s*кв\.?\s*\d+', '', addr).strip().lower()
        if target_clean in addr_clean or addr_clean in target_clean:
            # Ищем квартиру в этом адресе
            apt_match = re.search(r'кв\.?\s*(\d+)', addr, re.IGNORECASE)
            if apt_match:
                return apt_match.group(1)
    return None


async def find_apartments_via_deepseek(addresses_without_apt, addresses_with_apt):
    """Использует DeepSeek API для сопоставления адресов и поиска квартир"""
    if not addresses_without_apt or not addresses_with_apt:
        return {}
    
    # Формируем запрос: список с квартирами как примеры, список без квартир для поиска
    examples = "\n".join([f"  С квартирой: {a}" for a in addresses_with_apt[:20]])
    to_find = "\n".join([f"{i+1}. {a}" for i, a in enumerate(addresses_without_apt)])
    
    prompt = f"""Есть адреса с квартирами (примеры):
{examples}

Для каждого адреса ниже найди квартиру, сопоставив с примерами по городу, улице и дому.
Если точное совпадение не найдено — оставь адрес без квартиры.
Формат ответа: номер. Город, Улица, Дом, Квартира (каждое слово с прописной).
Адреса для поиска квартир:
{to_find}"""
    
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
                    {"role": "system", "content": "Ты помощник для сопоставления адресов и поиска квартир. Отвечай строго в формате: номер. Адрес с квартирой. Каждое слово с прописной буквы."},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.1,
                "max_tokens": 2000
            }
            async with session.post(url, headers=headers, json=payload, timeout=30) as resp:
                data = await resp.json()
                if data.get("choices"):
                    text = data["choices"][0]["message"]["content"].strip()
                    result = {}
                    for line in text.split('\n'):
                        line = line.strip()
                        if not line:
                            continue
                        # Извлекаем порядковый номер и адрес
                        m = re.match(r'^(\d+)[\.\)]\s*(.+)', line)
                        if m:
                            idx = int(m.group(1)) - 1
                            addr_with_apt = m.group(2).strip()
                            if idx < len(addresses_without_apt):
                                # Извлекаем квартиру
                                apt_match = re.search(r',\s*(\d+)\s*$', addr_with_apt)
                                if apt_match:
                                    apartment = apt_match.group(1)
                                    result[addresses_without_apt[idx]] = apartment
                    return result
    except Exception as e:
        print(f"[DEEPSEEK-APT] Ошибка: {e}")
    
    # Fallback: локальное сопоставление
    result = {}
    for addr_wo in addresses_without_apt:
        # Упрощаем адрес без квартиры до "город, улица, дом"
        addr_key = re.sub(r',?\s*\d*\s*$', '', addr_wo).strip().lower()
        addr_key = addr_key.rstrip(',').strip()
        
        for addr_w in addresses_with_apt:
            if addr_key in addr_w.lower():
                apt_match = re.search(r',\s*(\d+)\s*$', addr_w.strip())
                if apt_match:
                    result[addr_wo] = apt_match.group(1)
                    break
    
    return result


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
    
    table_num_col = 'N таблицы'
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
        if h == 'N таблицы':
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
    GEO_COLS = ['Номер', 'Адресс', 'ФИО', 'Дата', 'СНИЛС']
    
    idx_map = {}
    for col in GEO_COLS:
        found = -1
        for i, h in enumerate(headers):
            if h.lower() == col.lower():
                found = i
                break
        if found == -1:
            for i, h in enumerate(headers):
                hl = h.lower()
                if col.lower() in hl:
                    if col.lower() == 'номер' and ('паспорт' in hl or 'инн' in hl or 'passport' in hl or 'inn' in hl):
                        continue
                    found = i
                    break
        idx_map[col] = found
    
    for col, idx in idx_map.items():
        if idx == -1:
            raise ValueError(f'Колонка не найдена: {col}')
    
    out_rows = []
    phones = set()
    
    for row in rows:
        phone_raw = row[idx_map['Номер']] if idx_map['Номер'] < len(row) else ''
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
                row[idx_map['Адресс']] if idx_map['Адресс'] < len(row) else '',
                row[idx_map['ФИО']] if idx_map['ФИО'] < len(row) else '',
                row[idx_map['Дата']] if idx_map['Дата'] < len(row) else '',
                row[idx_map['СНИЛС']] if idx_map['СНИЛС'] < len(row) else ''
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
    raise Exception("Сессия недействительна")


# ====================== ПОДТВЕРЖДЕНИЯ ЧЕРЕЗ БОТА ======================
async def send_confirm_with_buttons(bot_token, chat_id, stage_name, count, confirm_id, topic_id=None):
    global stop_requested
    
    if stop_requested:
        return False
    
    text = f"ПОДТВЕРДИТЕ ПРОБИВ\n\nЭтап: {stage_name}\nСтрок: {count}"
    
    buttons = [
        [{"text": "ПОДТВЕРДИТЬ", "callback_data": f"confirm_{confirm_id}"}],
        [{"text": "ПРОПУСТИТЬ", "callback_data": f"skip_{confirm_id}"}],
        [{"text": "ОСТАНОВИТЬ ВСЁ", "callback_data": f"stop_{confirm_id}"}],
        [{"text": "ЕЩЁ РАЗ", "callback_data": f"again_{confirm_id}"}]
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
        print(f"[CONFIRM] Ошибка: {e}")
        return False


async def poll_updates_with_buttons(bot_token, chat_id, confirm_id, topic_id=None):
    global stop_requested
    offset = 0
    print(f"[POLL] Начинаю опрос для confirm_id={confirm_id}")
    
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
                                payload["text"] = "ПОДТВЕРЖДЕНО! Выполняю..."
                                await s.post(
                                    f"https://api.telegram.org/bot{bot_token}/editMessageText",
                                    json=payload
                                )
                                pending_confirms[confirm_id] = "confirm"
                                print(f"[POLL] ПОДТВЕРЖДЕНО")
                                return
                                    
                            elif cb_data == f"skip_{confirm_id}":
                                payload["text"] = "ЭТАП ПРОПУЩЕН"
                                await s.post(
                                    f"https://api.telegram.org/bot{bot_token}/editMessageText",
                                    json=payload
                                )
                                pending_confirms[confirm_id] = "skip"
                                print(f"[POLL] ПРОПУЩЕН")
                                return
                                    
                            elif cb_data == f"stop_{confirm_id}":
                                payload["text"] = "ОСТАНОВЛЕНО! Завершаю..."
                                await s.post(
                                    f"https://api.telegram.org/bot{bot_token}/editMessageText",
                                    json=payload
                                )
                                pending_confirms[confirm_id] = "stop"
                                stop_requested = True
                                print(f"[POLL] ОСТАНОВЛЕНО")
                                return
                                    
                            elif cb_data == f"again_{confirm_id}":
                                payload["text"] = "ЕЩЁ РАЗ! Отправляю заново..."
                                await s.post(
                                    f"https://api.telegram.org/bot{bot_token}/editMessageText",
                                    json=payload
                                )
                                pending_confirms[confirm_id] = "again"
                                print(f"[POLL] ЕЩЁ РАЗ")
                                return
        except Exception as e:
            print(f"[POLL] Ошибка: {e}")
        await asyncio.sleep(1)


async def safe_confirm_with_buttons(bot_token, chat_id, stage_name, count, confirm_id, add_log, topic_id=None):
    global stop_requested
    
    if stop_requested:
        add_log("[x] Остановка запрошена")
        return "stop"
    
    if not bot_token or not chat_id:
        add_log("[v] Бот не настроен - продолжаю автоматически")
        return "confirm"

    while True:
        sent = await send_confirm_with_buttons(bot_token, chat_id, stage_name, count, confirm_id, topic_id)
        if not sent:
            add_log("[!] Не удалось отправить в бот - повтор через 5с...")
            await asyncio.sleep(5)
            continue

        add_log(f"[ОЖИДАНИЕ] Откройте бот для: {stage_name}")
        
        while True:
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
                if r == "again":
                    add_log(f"[v] ЕЩЁ РАЗ - повтор для: {stage_name}")
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
            print(f"[BOT] Файл отправлен: {caption}")
    except Exception as e:
        print(f"[BOT] Ошибка: {e}")


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
            print(f"[BOT] ZIP отправлен: {caption}")
    except Exception as e:
        print(f"[BOT] ZIP ошибка: {e}")


async def send_txt_to_bot(bot_token, chat_id, content, filename, caption="", topic_id=None):
    """Отправляет TXT файл в бот"""
    try:
        txt_buffer = io.BytesIO(content.encode('utf-8'))
        txt_buffer.seek(0)
        
        async with aiohttp.ClientSession() as s:
            data = aiohttp.FormData()
            data.add_field('chat_id', str(chat_id))
            data.add_field('caption', caption)
            data.add_field('document', txt_buffer, filename=filename)
            if topic_id:
                data.add_field('message_thread_id', str(topic_id))
            await s.post(f"https://api.telegram.org/bot{bot_token}/sendDocument", data=data)
            print(f"[BOT] TXT отправлен: {filename}")
    except Exception as e:
        print(f"[BOT] TXT ошибка: {e}")


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
                                print(f"[BOT] Нажата кнопка '{btn.text}' в {bot}")
                                return True
            print(f"[BOT] Кнопка '{text}' не найдена в {bot} (попытка {attempt+1})")
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


async def wait_txt(client, bot, timeout=180, since_msg_id=None):
    """Ожидает TXT файл от бота"""
    e = await client.get_entity(bot)
    start = time.time()
    print(f"[BOT] Ожидаю TXT от {bot}...")
    while time.time() - start < timeout:
        msgs = await client.get_messages(e, limit=5)
        for msg in msgs:
            if not msg or not msg.document:
                continue
            if since_msg_id is not None and msg.id <= since_msg_id:
                continue
            for a in msg.document.attributes:
                if isinstance(a, DocumentAttributeFilename) and a.file_name.endswith('.txt'):
                    print(f"[BOT] Получен TXT: {a.file_name}")
                    return msg
        await asyncio.sleep(3)
    print(f"[BOT] TXT не получен")
    return None


async def wait_report(client, bot, phone, timeout=300):
    """Ожидает отчёт от бота по номеру телефона"""
    e = await client.get_entity(bot)
    start = time.time()
    print(f"[BOT] Ожидаю отчёт по номеру {phone} от {bot}...")
    
    # Отправляем запрос
    await client.send_message(e, phone)
    
    while time.time() - start < timeout:
        msgs = await client.get_messages(e, limit=10)
        for msg in msgs:
            if not msg or not msg.text:
                continue
            # Проверяем, что это отчёт по нашему номеру
            if phone in msg.text and ("ОТЧЕТ" in msg.text or "ЗАПРОС" in msg.text):
                print(f"[BOT] Получен отчёт по номеру {phone}")
                return msg.text
        await asyncio.sleep(5)
    print(f"[BOT] Отчёт по номеру {phone} не получен")
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
            if any(k in v for k in ['ИНН', 'INN', 'ПАСПОРТ', 'PASSPORT']):
                continue
            if any(k in v for k in ['ФИО', 'FIO', 'ИМЯ', 'ФАМИЛИЯ', 'NAME']):
                h['fio'] = col
            if any(k in v for k in ['ДАТА', 'DATE', 'РОЖД', 'BIRTH']):
                h['date'] = col
            if any(k in v for k in ['ТЕЛЕФОН', 'PHONE', 'ТЕЛ']):
                h['phone'] = col
            elif 'НОМЕР' in v and 'ПАСПОРТ' not in v and 'ИНН' not in v:
                h['phone'] = col
            if any(k in v for k in ['СНИЛС', 'SNILS']):
                h['snils'] = col
            if any(k in v for k in ['АДРЕС', 'АДРЕСС', 'ADDR', 'ADDRESS']):
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
                print(f"[FILL-PHONES] Строка {row}: ЗАПОЛНЕНО")
        else:
            for row in range(2, ws.max_row + 1):
                table_fio = normalize_fio_local(str(ws.cell(row=row, column=COL_FIO).value or ""))
                if table_fio == rec_fio:
                    existing = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
                    if not existing or existing == 'None':
                        ws.cell(row=row, column=COL_PHONE).value = clean_rec
                        filled += 1
                        print(f"[FILL-PHONES] Строка {row} (запасной вариант): ЗАПОЛНЕНО")
                    break
    
    print(f"[FILL-PHONES] ИТОГО заполнено номеров: {filled}")
    return filled


def fill_snils_dates(ws, response_records):
    filled = 0
    print(f"[FILL-SNILS] Начинаю заполнение дат по СНИЛС")
    
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
                print(f"[FILL-SNILS] Строка {row}: ЗАПОЛНЕНО по СНИЛС")
                found = True
        
        if not found and rec_fio:
            for row in range(2, ws.max_row + 1):
                table_fio = normalize_fio_local(str(ws.cell(row=row, column=COL_FIO).value or ""))
                if table_fio == rec_fio:
                    existing = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
                    if not existing or existing == 'None':
                        ws.cell(row=row, column=COL_DATE).value = rec_date
                        filled += 1
                        print(f"[FILL-SNILS] Строка {row} (запасной вариант): ЗАПОЛНЕНО")
                    found = True
                    break
    
    print(f"[FILL-SNILS] ИТОГО заполнено дат: {filled}")
    return filled


def fill_phones_by_numbers(ws, phone_to_fio_map):
    """Заполняет ФИО по номерам из TXT (как в minecraft.py)"""
    filled = 0
    print(f"[FILL-PHONES-BY-NUMBERS] Начинаю заполнение ФИО по номерам. Карта: {len(phone_to_fio_map)}")
    
    for row in range(2, ws.max_row + 1):
        phone_val = clean_phone(str(ws.cell(row=row, column=COL_PHONE).value or "").strip())
        if not phone_val:
            continue
        
        # Ищем номер в карте (без +7)
        phone_digits = re.sub(r'[^0-9]', '', phone_val)
        if len(phone_digits) >= 10:
            phone_key = '7' + phone_digits[-10:]
        else:
            phone_key = phone_digits
        
        if phone_key in phone_to_fio_map:
            fio_val = phone_to_fio_map[phone_key]
            if fio_val:
                ws.cell(row=row, column=COL_FIO).value = normalize_fio_local(fio_val)
                filled += 1
                print(f"[FILL-PHONES-BY-NUMBERS] Строка {row}: ЗАПОЛНЕНО ФИО для {phone_key}")
    
    print(f"[FILL-PHONES-BY-NUMBERS] ИТОГО заполнено ФИО: {filled}")
    return filled


def fill_apartments_from_report(ws, address_to_apartment_map):
    """Заполняет квартиры по нормализованному адресу из словаря"""
    filled = 0
    print(f"[FILL-APARTMENTS] Начинаю заполнение квартир. Карта: {len(address_to_apartment_map)}")
    
    for row in range(2, ws.max_row + 1):
        addr_val = str(ws.cell(row=row, column=COL_ADDR).value or "").strip()
        if not addr_val or addr_val == 'None':
            continue
        
        # Нормализуем адрес для сравнения
        addr_clean = normalize_address_local(addr_val)
        
        # Проверяем ключи в мапе
        for key, apartment in address_to_apartment_map.items():
            key_clean = normalize_address_local(key)
            if key_clean == addr_clean or key_clean in addr_clean or addr_clean in key_clean:
                # Добавляем квартиру после запятой
                ws.cell(row=row, column=COL_ADDR).value = f"{addr_val.rstrip(',')}, {apartment}"
                filled += 1
                print(f"[FILL-APARTMENTS] Строка {row}: +кв {apartment}")
                break
    
    print(f"[FILL-APARTMENTS] ИТОГО заполнено квартир: {filled}")
    return filled


# ====================== ДОБИВ ЧЕРЕЗ САУРОН ======================
async def dobiv_sauron(client, bot, fio, date, account_id, row_num, ws, wb, result_file, add_log):
    try:
        norm_date = parse_date(date)
        if not norm_date:
            norm_date = date
        
        query = f"{fio} {norm_date}"
        add_log(f"[ДОБИВ] Акк {account_id}, строка {row_num}: {query}")
        
        await client.send_message(bot, query)
        await asyncio.sleep(5)
        
        async for msg in client.iter_messages(bot, limit=10):
            if msg.text and ("ОТЧЕТ" in msg.text or "ТЕЛЕФОНЫ" in msg.text):
                phones = extract_phones_from_text(msg.text)
                if phones:
                    add_log(f"[ДОБИВ] Найдены телефоны: {phones}")
                    return phones
                break
        
        return []
    except Exception as e:
        add_log(f"[ДОБИВ] Ошибка: {e}")
        return []


async def dobiv_by_numbers(client, bot, phone, add_log):
    """Пробивает номер телефона через бот 2 и возвращает отчёт"""
    try:
        # Добавляем 7 в начале, если её нет
        phone_clean = clean_phone(phone)
        phone_for_send = clean_phone_without_plus(phone)
        if not phone_for_send.startswith('7'):
            phone_for_send = '7' + phone_for_send
        
        add_log(f"[ДОБИВ-КВАРТИР] Отправляю номер: {phone_for_send}")
        
        # Ждём отчёт
        report = await wait_report(client, bot, phone_for_send, timeout=300)
        if report:
            add_log(f"[ДОБИВ-КВАРТИР] Получен отчёт для {phone_for_send}")
            return report
        
        return None
    except Exception as e:
        add_log(f"[ДОБИВ-КВАРТИР] Ошибка: {e}")
        return None


# ====================== ПОИСК ИНН В ГРУППЕ ======================
async def find_inn_in_group(client, group_id, address, topic_id=None):
    try:
        entity = await client.get_entity(int(group_id))
        
        search_parts = address.split(',')
        search_terms = []
        for part in search_parts:
            part = part.strip()
            if part and len(part) > 3:
                search_terms.append(part)
        
        async for msg in client.iter_messages(entity, limit=100):
            if not msg.text:
                continue
            
            msg_lower = msg.text.lower()
            address_lower = address.lower()
            
            match_count = 0
            for term in search_terms[:3]:
                if term.lower() in msg_lower:
                    match_count += 1
            
            if match_count >= 2:
                inn = extract_inn_from_text(msg.text)
                if inn:
                    return inn, True
                else:
                    return None, True
        
        return None, False
    except Exception as e:
        print(f"[GROUP] Ошибка поиска: {e}")
        return None, False


# ====================== ПЕРЕИМЕНОВАНИЕ ФАЙЛОВ ======================
async def rename_files_by_address(ws, client, group_id, topic_id, add_log, tables_names):
    if not group_id:
        add_log("[ПЕРЕИМЕНОВАНИЕ] Группа не указана - пропускаю")
        return tables_names
    
    add_log("[ПЕРЕИМЕНОВАНИЕ] Начинаю переименование по адресам...")
    
    address_map = {}
    table_nums = {}
    
    for row in range(2, ws.max_row + 1):
        table_num = str(ws.cell(row=row, column=COL_NO).value or "").strip()
        addr = str(ws.cell(row=row, column=COL_ADDR).value or "").strip()
        
        if table_num and addr and addr != 'None':
            if table_num not in address_map:
                address_map[table_num] = addr
                table_nums[table_num] = row
    
    add_log(f"[ПЕРЕИМЕНОВАНИЕ] Найдено {len(address_map)} уникальных адресов")
    
    new_names = {}
    for table_num, addr in address_map.items():
        clean_addr = re.sub(r',?\s*кв\.?\s*\d+', '', addr).strip()
        clean_addr = re.sub(r',\s*,', ',', clean_addr)
        
        add_log(f"[ПЕРЕИМЕНОВАНИЕ] Ищу ИНН для: {clean_addr}")
        
        inn, found = await find_inn_in_group(client, group_id, clean_addr, topic_id)
        
        if found and inn:
            new_name = f"{inn}_{clean_addr}"
            add_log(f"[ПЕРЕИМЕНОВАНИЕ] Найден ИНН: {inn} -> {new_name}")
        else:
            new_name = f"УК_{clean_addr}"
            add_log(f"[ПЕРЕИМЕНОВАНИЕ] ИНН не найден -> {new_name}")
        
        new_name = re.sub(r'[<>:"/\\|?*]', '_', new_name)
        new_names[table_num] = new_name
    
    updated_names = []
    for i, name in enumerate(tables_names):
        table_num = str(i + 1)
        if table_num in new_names:
            updated_names.append(new_names[table_num])
        else:
            updated_names.append(name)
    
    add_log(f"[ПЕРЕИМЕНОВАНИЕ] Переименовано {len(new_names)} файлов")
    return updated_names


# ====================== ФИЛЬТР ПО НОМЕРАМ (как в сименамиmax.html) ======================
def filter_by_phones(headers, rows, phone_list):
    """Фильтрует таблицу по списку номеров (как в сименамиmax.html)"""
    if not phone_list:
        return rows
    
    # Находим колонку с номером
    phone_idx = -1
    for i, h in enumerate(headers):
        hl = h.lower()
        if hl in ['номер', 'телефон', 'phone', 'tel', 'nomer']:
            phone_idx = i
            break
        if 'номер' in hl or 'телефон' in hl or 'phone' in hl:
            phone_idx = i
            break
    
    if phone_idx == -1:
        return rows
    
    # Создаём множество номеров для быстрого поиска
    phone_set = set()
    for p in phone_list:
        clean = clean_phone(p)
        if clean:
            digits = re.sub(r'[^0-9]', '', clean)
            if len(digits) >= 10:
                phone_set.add('7' + digits[-10:])
    
    # Фильтруем строки
    filtered = []
    for row in rows:
        if phone_idx >= len(row):
            continue
        phone_val = clean_phone(str(row[phone_idx] or ''))
        if not phone_val:
            continue
        digits = re.sub(r'[^0-9]', '', phone_val)
        if len(digits) >= 10:
            key = '7' + digits[-10:]
            if key in phone_set:
                filtered.append(row)
    
    return filtered


# ====================== ПОЛНЫЙ ЦИКЛ ПРОБИВА (8 ЭТАПОВ) ======================
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
                await send_file_to_bot(bot_token, chat_id, file_path, f"Таблица после: {stage_label}", topic_id)
            else:
                wb.save(result_file)
                await send_file_to_bot(bot_token, chat_id, result_file, f"Таблица после: {stage_label}", topic_id)

    async def send_final_zip(complete_session=False):
        if not bot_token or not chat_id:
            return
        
        split_result = split_by_table_num(
            [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)],
            [[ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)] for r in range(2, ws.max_row + 1)]
        )
        
        if not split_result:
            await send_file_to_bot(bot_token, chat_id, result_file, "ИТОГОВЫЙ ФАЙЛ (все этапы)", topic_id)
            return
        
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
                    
                    idx = int(table_num) - 1
                    if idx < len(final_names):
                        name = f"{final_names[idx]}.xlsx"
                    else:
                        name = f"ГЕО_{table_num}.xlsx"
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
        
        caption = "ИТОГОВЫЙ ZIP АРХИВ (все таблицы + numbers.txt)"
        if complete_session:
            caption = "СЕССИЯ ЗАВЕРШЕНА! Все файлы с именами ИНН/УК_Адрес."
        
        await send_zip_to_bot(bot_token, chat_id, zip_path, caption, topic_id)
        add("[v] ZIP архив отправлен в бот")
        
        return zip_path

    async def send_txt_for_max_check():
        """Отправляет TXT с номерами для чека максов"""
        if not bot_token or not chat_id:
            return
        
        phones = []
        for row in range(2, ws.max_row + 1):
            phone_val = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
            if phone_val and phone_val != 'None' and phone_val != '0':
                clean = clean_phone_without_plus(phone_val)
                if clean:
                    phones.append(clean)
        
        if phones:
            phones = list(set(phones))
            content = '\n'.join(phones)
            caption = "ЧЕК МАКСОВ — проверьте эти номера.\n\nДАЛЕЕ: отправьте TXT с форматом 'Номер Имя' для добива ФИО, или нажмите 'Завершить сессию' на сайте."
            await send_txt_to_bot(bot_token, chat_id, content, "check_max.txt", caption, topic_id)
            add(f"[TXT] Отправлен check_max.txt с {len(phones)} номерами")

    # === СОЗДАЁМ ТАБЛИЦУ ===
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

    # === НОРМАЛИЗАЦИЯ ФИО ===
    add("[DEEPSEEK] Нормализация ФИО...")
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
        
        add(f"[DEEPSEEK] Пакет ФИО {batch_idx//batch_size + 1}/{(len(fio_values)+batch_size-1)//batch_size}: {len(batch)} имен")
        await asyncio.sleep(0.3)
    wb.save(result_file)
    add("[DEEPSEEK] Нормализация ФИО завершена")

    # === НОРМАЛИЗАЦИЯ АДРЕСОВ ===
    add("[DEEPSEEK] Нормализация адресов...")
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
        
        add(f"[DEEPSEEK] Пакет адресов {batch_idx//batch_size + 1}/{(len(addr_values)+batch_size-1)//batch_size}: {len(batch)} адресов")
        await asyncio.sleep(0.3)
    wb.save(result_file)
    add("[DEEPSEEK] Нормализация адресов завершена")

    # === АНАЛИЗ ===
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

        result = await safe_confirm_with_buttons(bot_token, chat_id, "ЭТАП 1: ФИО+НОМЕР", len(items_no_date), cid, add, topic_id)
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
                add("[!] Бот1 не ответил на этапе 1")
    
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

        result = await safe_confirm_with_buttons(bot_token, chat_id, "ЭТАП 2: СНИЛС (бот1)", len(real_snils), cid, add, topic_id)
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
                add(f"Получено по СНИЛС: {len(recs)}")
                fill_snils_dates(ws, recs)
                wb.save(result_file)
                await send_status("этап 2")
            else:
                add("[!] Бот1 не ответил на этапе 2")
    
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

        result = await safe_confirm_with_buttons(bot_token, chat_id, "ЭТАП 3: СНИЛС (бот2)", len(snils_still_empty), cid, add, topic_id)
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
                add(f"Получено по СНИЛС бот2: {len(recs)}")
                fill_snils_dates(ws, recs)
                wb.save(result_file)
                await send_status("этап 3")
            else:
                add("[!] Бот2 не ответил на этапе 3")
    
    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # ============ ЭТАП 3.5: ПРОБИВ ПО НОМЕРУ (НОВЫЙ ЭТАП) ============
    # Собираем номера для строк, у которых всё ещё нет даты
    phones_for_probev = []
    phone_rows = []
    for row in range(2, ws.max_row + 1):
        existing_date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
        if not existing_date or existing_date == 'None' or existing_date == '0':
            phone_val = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
            if phone_val and phone_val != 'None' and phone_val != '0':
                clean = clean_phone(phone_val)
                if clean:
                    phones_for_probev.append(clean)
                    phone_rows.append(row)
    
    if phones_for_probev and not stop_requested:
        cid = f"s35_{int(time.time())}"
        add(f"ЭТАП 3.5: ПРОБИВ ПО НОМЕРУ -> бот2 ({len(phones_for_probev)} номеров)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "ЭТАП 3.5: ПРОБИВ ПО НОМЕРУ", len(phones_for_probev), cid, add, topic_id)
        if result == "stop":
            await send_final_zip()
            return {"ok": True, "log": log, "stopped": True}
        if result == "skip":
            add("[v] ЭТАП 3.5 ПРОПУЩЕН")
        elif result == "confirm":
            add("[v] ПОДТВЕРЖДЕНО! Начинаю пробив по номерам...")

            # Формируем TXT с номерами (без +7)
            txt_lines = []
            for phone in phones_for_probev:
                clean = clean_phone_without_plus(phone)
                if clean:
                    txt_lines.append(clean)
            
            txt_content = '\n'.join(txt_lines)
            tpath = os.path.join(TEMP_DIR, f"t35_{int(time.time())}.txt")
            with open(tpath, 'w', encoding='utf-8') as f:
                f.write(txt_content)

            e = await client.get_entity(bot2)
            last_msgs = await client.get_messages(e, limit=1)
            last_msg_id = last_msgs[0].id if last_msgs else 0
            await client.send_file(e, tpath)
            add("Номера отправлены в бот2, жду ответ...")

            msg = await wait_xlsx(client, bot2, 300, since_msg_id=last_msg_id)
            if msg:
                rpath = os.path.join(TEMP_DIR, f"r35_{int(time.time())}.xlsx")
                await client.download_media(msg, file=rpath)
                recs = parse_xlsx(rpath)
                add(f"Получено по номерам: {len(recs)}")
                fill_dates_from_response(ws, recs)
                wb.save(result_file)
                await send_status("этап 3.5 (пробив по номеру)")
            else:
                add("[!] Бот2 не ответил на этапе 3.5")
    
    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # ============ ЭТАП 4: ФИЛЬТР ПО ГОДАМ ============
    if year_range and not stop_requested:
        try:
            parts = year_range.split('-')
            yf, yt = int(parts[0]), int(parts[1])
            add(f"ЭТАП 4: Фильтр годов {yf}-{yt}")

            cid = f"s4_{int(time.time())}"
            result = await safe_confirm_with_buttons(bot_token, chat_id, f"ЭТАП 4: Фильтр {yf}-{yt}", ws.max_row - 1, cid, add, topic_id)
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
                'fio': normalize_fio_local(fio_val),
                'date': date_val
            })

    if items_no_phone_after_stage4 and not stop_requested:
        cid = f"s5_{int(time.time())}"
        add(f"ЭТАП 5: ФИО+ДАТА -> бот1 ({len(items_no_phone_after_stage4)} строк)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "ЭТАП 5: ФИО+ДАТА (бот1)", len(items_no_phone_after_stage4), cid, add, topic_id)
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
                add("[!] Бот1 не ответил на этапе 5")
    
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
                'fio': normalize_fio_local(fio_val),
                'date': date_val
            })

    if items_no_phone_after_stage5 and not stop_requested:
        cid = f"s6_{int(time.time())}"
        add(f"ЭТАП 6: ФИО+ДАТА -> бот2 ({len(items_no_phone_after_stage5)} строк)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "ЭТАП 6: ФИО+ДАТА (бот2)", len(items_no_phone_after_stage5), cid, add, topic_id)
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
                add(f"Получено от бот2: {len(recs)}")
                fill_phones_from_response(ws, recs)
                wb.save(result_file)
                await send_status("этап 6")
            else:
                add("[!] Бот2 не ответил на этапе 6")
    
    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # ============ ЭТАП 7: ДОБИВ -> БОТ2 (как в minecraft.py) ============
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
        add(f"ЭТАП 7: ДОБИВ -> бот2 ({len(items_no_phone_final)} строк)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "ЭТАП 7: ДОБИВ (бот2)", len(items_no_phone_final), cid, add, topic_id)
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
                    
                    phones = await dobiv_sauron(client, e, fio, date, "1", i+1, ws, wb, result_file, add)
                    
                    if phones:
                        for row in range(2, ws.max_row + 1):
                            table_fio = normalize_fio_local(str(ws.cell(row=row, column=COL_FIO).value or ""))
                            table_date = str(ws.cell(row=row, column=COL_DATE).value or "").strip()
                            existing_phone = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()

                            if table_fio == normalize_fio_local(fio) and table_date == date:
                                if not existing_phone or existing_phone == 'None':
                                    ws.cell(row=row, column=COL_PHONE).value = phones[0]
                                    add(f"[ДОБИВ] {fio} -> {phones[0]}")
                                break

                    if (i + 1) % 5 == 0:
                        wb.save(result_file)
                        add(f"[ДОБИВ] Сохранено после {i+1} строк")
                        await asyncio.sleep(2)

                except FloodWaitError as e:
                    add(f"[ДОБИВ] FloodWait: {e.seconds}с")
                    await asyncio.sleep(e.seconds)
                except Exception as e:
                    add(f"[ДОБИВ] Ошибка: {e}")

            wb.save(result_file)
            await send_status("этап 7 (добив)")
            add("Добив завершён")

    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # ============ ЭТАП 8: ДОБИВ КВАРТИР (ГЛУБОКАЯ ИНТЕГРАЦИЯ) ============
    # Этап 8a: Собираем все адреса — и с квартирами, и без
    all_addresses = []
    addresses_with_apt = []  # Адреса, где уже есть квартира (примеры)
    addresses_without_apt = []  # Адреса без квартиры
    rows_without_apartment = []
    
    for row in range(2, ws.max_row + 1):
        addr_val = str(ws.cell(row=row, column=COL_ADDR).value or "").strip()
        if not addr_val or addr_val == 'None':
            continue
        
        has_apt = bool(re.search(r',\s*\d+\s*$', addr_val))
        addr_normalized = normalize_address_local(addr_val)
        
        if has_apt:
            if addr_normalized not in addresses_with_apt:
                addresses_with_apt.append(addr_normalized)
        else:
            phone_val = str(ws.cell(row=row, column=COL_PHONE).value or "").strip()
            if phone_val and phone_val != 'None' and phone_val != '0':
                if addr_normalized not in addresses_without_apt:
                    addresses_without_apt.append(addr_normalized)
                rows_without_apartment.append({
                    'row': row,
                    'address': addr_val,
                    'address_normalized': addr_normalized,
                    'phone': clean_phone(phone_val)
                })

    if rows_without_apartment and not stop_requested:
        cid = f"s8_{int(time.time())}"
        add(f"ЭТАП 8: ДОБИВ КВАРТИР ({len(rows_without_apartment)} адресов без кв, {len(addresses_with_apt)} с кв для примера)")

        result = await safe_confirm_with_buttons(bot_token, chat_id, "ЭТАП 8: ДОБИВ КВАРТИР", len(rows_without_apartment), cid, add, topic_id)
        if result == "stop":
            await send_final_zip()
            return {"ok": True, "log": log, "stopped": True}
        if result == "skip":
            add("[v] ЭТАП 8 ПРОПУЩЕН")
        elif result == "confirm":
            add("[v] ПОДТВЕРЖДЕНО! Начинаю добив квартир...")

            e = await client.get_entity(bot2)
            address_to_apartment = {}
            
            # Этап 8a: Пробив номеров через бот2 для получения адресов с квартирами
            unique_phones = list(set([item['phone'] for item in rows_without_apartment]))
            add(f"[ДОБИВ-КВАРТИР] Уникальных номеров для пробива: {len(unique_phones)}")
            
            probed_addresses = []  # Адреса, полученные из отчётов
            
            for i, phone in enumerate(unique_phones):
                if stop_requested:
                    break
                
                try:
                    # Добавляем 7 при отправке в бот
                    phone_for_send = clean_phone_without_plus(phone)
                    if not phone_for_send.startswith('7'):
                        phone_for_send = '7' + phone_for_send
                    
                    add(f"[ДОБИВ-КВАРТИР] Пробив номера: {phone_for_send}")
                    
                    report = await dobiv_by_numbers(client, e, phone, add)
                    
                    if report:
                        # Сохраняем отчёт в файл с именем, содержащим номер
                        report_filename = f"report_{phone_for_send}.txt"
                        report_path = os.path.join(TEMP_DIR, report_filename)
                        with open(report_path, 'w', encoding='utf-8') as f:
                            f.write(report)
                        
                        # Отправляем отчёт в бот
                        await send_file_to_bot(bot_token, chat_id, report_path, f"Отчёт по номеру {phone_for_send}", topic_id)
                        
                        # Извлекаем адрес из отчёта
                        report_addr = extract_address_from_report(report)
                        if report_addr:
                            normalized = normalize_address_local(report_addr)
                            # Проверяем, есть ли квартира
                            if re.search(r',\s*\d+\s*$', normalized):
                                if normalized not in probed_addresses:
                                    probed_addresses.append(normalized)
                                    add(f"[ДОБИВ-КВАРТИР] Найден адрес с кв: {normalized}")
                            elif normalized not in probed_addresses:
                                probed_addresses.append(normalized)
                    
                    if (i + 1) % 5 == 0:
                        add(f"[ДОБИВ-КВАРТИР] Пробито {i+1}/{len(unique_phones)} номеров")
                        await asyncio.sleep(2)
                        
                except FloodWaitError as e:
                    add(f"[ДОБИВ-КВАРТИР] FloodWait: {e.seconds}с")
                    await asyncio.sleep(e.seconds)
                except Exception as e:
                    add(f"[ДОБИВ-КВАРТИР] Ошибка пробива {phone}: {e}")
            
            # Этап 8b: Сопоставление адресов через DeepSeek
            all_examples = list(set(addresses_with_apt + probed_addresses))
            add(f"[ДОБИВ-КВАРТИР] Примеров адресов с кв: {len(all_examples)}")
            add(f"[ДОБИВ-КВАРТИР] Адресов без кв: {len(addresses_without_apt)}")
            
            if addresses_without_apt and all_examples:
                add("[ДОБИВ-КВАРТИР] Запуск DeepSeek для сопоставления адресов...")
                apt_map = await find_apartments_via_deepseek(addresses_without_apt, all_examples)
                
                if apt_map:
                    # Заполняем квартиры в таблице
                    for item in rows_without_apartment:
                        addr_norm = item.get('address_normalized', '')
                        if addr_norm in apt_map:
                            apartment = apt_map[addr_norm]
                            row = item['row']
                            current_addr = str(ws.cell(row=row, column=COL_ADDR).value or "").strip()
                            # Добавляем квартиру после запятой
                            ws.cell(row=row, column=COL_ADDR).value = f"{current_addr.rstrip(',')}, {apartment}"
                            address_to_apartment[current_addr] = apartment
                            add(f"[ДОБИВ-КВАРТИР] Строка {row}: +кв {apartment} -> {current_addr}, {apartment}")
                    
                    wb.save(result_file)
                    await send_status("этап 8 (добив квартир)")
                    add(f"[ДОБИВ-КВАРТИР] Заполнено квартир через DeepSeek: {len(apt_map)}")
                else:
                    add("[ДОБИВ-КВАРТИР] DeepSeek не нашёл совпадений")
            else:
                add("[ДОБИВ-КВАРТИР] Недостаточно данных для сопоставления")

    if stop_requested:
        await send_final_zip()
        return {"ok": True, "log": log, "stopped": True}

    # === ОТПРАВЛЯЕМ TXT ДЛЯ ЧЕКА МАКСОВ ===
    await send_txt_for_max_check()

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
    return web.json_response({"ok": True, "message": "X Backend v17.0"})


async def handle_root(request):
    try:
        html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "сайт.html")
        with open(html_path, "r", encoding="utf-8") as f:
            html = f.read()
        return web.Response(text=html, content_type="text/html")
    except Exception as e:
        print(f"[ROOT] Ошибка: {e}")
        return web.Response(text="OK", content_type="text/plain")


async def handle_upload_zip(request):
    try:
        reader = await request.multipart()
        field = await reader.next()
        if field.name != 'file':
            return web.json_response({"ok": False, "error": "Нет файла"}, status=400)
        
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
            return web.json_response({"ok": False, "error": "Таблицы не найдены"}, status=400)
        
        merged_path = os.path.join(TEMP_DIR, f"merged_{int(time.time())}.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.append(merged['headers'])
        for row in merged['rows']:
            ws.append(row)
        wb.save(merged_path)
        
        names = [t.get('name', f'Таблица_{i+1}') for i, t in enumerate(tables_data)]
        
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
        print(f"[AUTH] Вход: {me.first_name}")
        return web.json_response({
            "ok": True,
            "session": ss,
            "phone": phone,
            "name": me.first_name or ""
        })
    except Exception as e:
        print(f"[AUTH] Ошибка: {e}")
        return web.json_response({"ok": False, "error": str(e)}, status=400)


async def handle_upload_check_txt(request):
    """Загрузка TXT с номерами для добива (как в minecraft.py)"""
    try:
        reader = await request.multipart()
        field = await reader.next()
        if field.name != 'file':
            return web.json_response({"ok": False, "error": "Нет файла"}, status=400)
        
        data = await field.read()
        txt_content = data.decode('utf-8', errors='ignore')
        
        # Парсим TXT: каждая строка - номер или "номер имя"
        phone_to_fio = {}
        for line in txt_content.split('\n'):
            line = line.strip()
            if not line:
                continue
            
            # Пробуем разбить на номер и имя
            parts = line.split(maxsplit=1)
            if len(parts) >= 2:
                phone_raw = parts[0]
                fio = parts[1]
                phone_clean = clean_phone_without_plus(phone_raw)
                if phone_clean and fio:
                    phone_to_fio[phone_clean] = fio
            else:
                # Только номер
                phone_clean = clean_phone_without_plus(line)
                if phone_clean:
                    phone_to_fio[phone_clean] = ''
        
        return web.json_response({
            "ok": True,
            "phone_to_fio": phone_to_fio,
            "count": len(phone_to_fio)
        })
    except Exception as e:
        print(f"[UPLOAD-TXT] Ошибка: {e}")
        traceback.print_exc()
        return web.json_response({"ok": False, "error": str(e)}, status=500)


async def handle_full_probev(request):
    try:
        d = await request.json()
        ss = d.get("session", "")
        bot1 = d.get("bot1", "@osint_pam_pam_bot")
        bot2 = d.get("bot2", "@proverim123_bot")
        bot_token = d.get("bot_token", "")
        chat_id = d.get("chat_id", "")
        topic_id = d.get("topic_id", None)
        group_id = d.get("group_id", None)
        year_range = d.get("year_range", "") or "1945-1975"
        items_no_date = d.get("items_no_date", [])
        items_no_phone = d.get("items_no_phone", [])
        items_snils = d.get("items_snils", [])
        original_rows = d.get("original_rows", [])
        tables_names = d.get("tables_names", [])

        if not ss:
            return web.json_response({"ok": False, "error": "Нет сессии"}, status=400)

        async with probev_lock:
            if ss in active_probevs:
                task = active_probevs[ss]
                if not task.done():
                    return web.json_response({"ok": False, "error": "Пробив уже выполняется"}, status=409)
                else:
                    del active_probevs[ss]

        print(f"\n{'='*60}")
        print(f"[PROBEV] ЗАПУСК ПОЛНОГО ЦИКЛА (8 ЭТАПОВ + DeepSeek + ДОБИВ КВАРТИР)")
        print(f"[PROBEV] Бот1: {bot1}, Бот2: {bot2}")
        print(f"[PROBEV] Строк без даты: {len(items_no_date)}")
        print(f"[PROBEV] Строк без номера: {len(items_no_phone)}")
        print(f"[PROBEV] Группа: {group_id or 'Нет'}")
        print(f"[PROBEV] Тема: {topic_id or 'Нет'}")
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
                print(f"[PROBEV] Очистка завершена для сессии {ss[:10]}...")

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
    return web.json_response({"ok": True, "message": "Остановка запрошена"})


async def handle_finish_session(request):
    """Завершение сессии: отправка всех актуальных файлов с именами ИНН/УК_Адрес в бот"""
    try:
        d = await request.json()
        ss = d.get("session", "")
        tables_names = d.get("tables_names", [])
        bot_token = d.get("bot_token", "")
        chat_id = d.get("chat_id", "")
        group_id = d.get("group_id", None)
        topic_id = d.get("topic_id", None)
        
        if not ss:
            return web.json_response({"ok": False, "error": "Нет сессии"}, status=400)
        
        # Ищем самый свежий result файл
        result_files = []
        for f in os.listdir(TEMP_DIR):
            if f.startswith('result_') and f.endswith('.xlsx'):
                result_files.append(os.path.join(TEMP_DIR, f))
        
        if not result_files:
            return web.json_response({"ok": True, "message": "Нет файлов для отправки"})
        
        # Сортируем по времени создания (самый новый первый)
        result_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        latest_file = result_files[0]
        
        # Загружаем таблицу
        wb = load_workbook(latest_file, data_only=True)
        ws = wb.active
        
        # Переименовываем файлы по ИНН/адресу
        if bot_token and chat_id and group_id:
            try:
                client = await get_client(ss)
                renamed = await rename_files_by_address(ws, client, group_id, topic_id, None, tables_names)
                
                # Отправляем каждый файл с новым именем
                split_result = split_by_table_num(
                    [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)],
                    [[ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)] for r in range(2, ws.max_row + 1)]
                )
                
                if split_result:
                    for table_num, table_data in split_result.items():
                        geo_result = geo_filter(table_data['headers'], table_data['rows'])
                        ws_data = [geo_result['headers']] + geo_result['rows']
                        
                        wb_temp = Workbook()
                        ws_temp = wb_temp.active
                        for row in ws_data:
                            ws_temp.append(row)
                        
                        xlsx_path = os.path.join(TEMP_DIR, f"session_{int(time.time())}_{table_num}.xlsx")
                        wb_temp.save(xlsx_path)
                        
                        idx = int(table_num) - 1
                        name = renamed[idx] if idx < len(renamed) else f"ГЕО_{table_num}"
                        
                        # Отправляем в бот
                        caption = f"Сессия завершена: {name}"
                        await send_file_to_bot(bot_token, chat_id, xlsx_path, caption, topic_id)
            except Exception as e:
                print(f"[FINISH] Ошибка отправки: {e}")
        else:
            # Просто отправляем итоговый файл
            if bot_token and chat_id:
                await send_file_to_bot(bot_token, chat_id, latest_file, "Сессия завершена (Итоговая таблица)", topic_id)
        
        return web.json_response({
            "ok": True,
            "message": "Сессия завершена, файлы отправлены в бот",
            "files_count": len(result_files)
        })
    except Exception as e:
        print(f"[FINISH] Ошибка: {e}")
        traceback.print_exc()
        return web.json_response({"ok": False, "error": str(e)}, status=500)


async def handle_normalize_addresses(request):
    """Нормализация адресов через DeepSeek API (вызывается из вкладки Адреса)"""
    try:
        d = await request.json()
        addresses = d.get("addresses", [])
        if not addresses:
            return web.json_response({"ok": False, "error": "Нет адресов"}, status=400)
        
        # Используем ту же функцию нормализации что и в основном цикле
        normalized = await normalize_batch_deepseek(addresses, 'address')
        
        return web.json_response({
            "ok": True,
            "normalized": normalized,
            "count": len(normalized)
        })
    except Exception as e:
        print(f"[NORM-ADDR] Ошибка: {e}")
        traceback.print_exc()
        return web.json_response({"ok": False, "error": str(e)}, status=500)


# ====================== ЗАПУСК ======================
app = web.Application(middlewares=[log_and_cors], client_max_size=200 * 1024 * 1024)
app.router.add_get("/", handle_root)
app.router.add_get("/health", handle_health)
app.router.add_post("/upload-zip", handle_upload_zip)
app.router.add_post("/upload-check-txt", handle_upload_check_txt)
app.router.add_post("/send-code", handle_send_code)
app.router.add_post("/verify-code", handle_verify_code)
app.router.add_post("/full-probev", handle_full_probev)
app.router.add_post("/stop", handle_stop)
app.router.add_post("/finish-session", handle_finish_session)
app.router.add_post("/normalize-addresses", handle_normalize_addresses)


async def on_startup(app):
    port = app["port"]
    msg = (
        "=" * 60 + "\n"
        f"X Backend v17.0 ЗАПУЩЕН (8 этапов + DeepSeek квартиры + сессии)\n"
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
    print("[SHUTDOWN] Готово.", flush=True)


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
            print("[SIGNAL] Останавливаю...", flush=True)
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