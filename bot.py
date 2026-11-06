import asyncio
import json
import logging
import os
import re
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import aiohttp
import pandas as pd
import pytz
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardMarkup,
)
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from bs4 import BeautifulSoup

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)

# Переменные окружения из Railway
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
RAW_GROUP_ID = os.getenv("TARGET_GROUP_ID", "").strip()
RAW_THREAD_ID = os.getenv("TARGET_THREAD_ID", "").strip()  # ID вкладки (топика) в группе, если есть

TIMEZONE = pytz.timezone("Asia/Tbilisi")

EXCEL_SOURCE_PATH = "Брони.xlsx"
STATE_DB_PATH = "fines_state.json"
REPORT_EXCEL_PATH = "Отчет_штрафы_Батуми.xlsx"

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())
scheduler = AsyncIOScheduler(timezone=TIMEZONE)

# Флаг, чтобы нельзя было запустить две проверки одновременно
is_checking_now = False


def get_target_group_id() -> Optional[int | str]:
    if not RAW_GROUP_ID:
        return None
    if RAW_GROUP_ID.startswith("-") and RAW_GROUP_ID[1:].isdigit():
        return int(RAW_GROUP_ID)
    if RAW_GROUP_ID.isdigit():
        return int(RAW_GROUP_ID)
    return RAW_GROUP_ID


def get_target_thread_id() -> Optional[int]:
    if RAW_THREAD_ID and RAW_THREAD_ID.isdigit():
        return int(RAW_THREAD_ID)
    return None


# =====================================================================
# 1. ЧТЕНИЕ И ОЧИСТКА БАЗЫ МАШИН ИЗ ВКЛАДКИ "Номера" (Брони.xlsx)
# =====================================================================
def clean_plate(val: any) -> str:
    """Приводит госномер к единому формату (например, 'NZ-001-GU' -> 'NZ001GU', 'Ew147we' -> 'EW147WE')."""
    if pd.isna(val):
        return ""
    s = re.sub(r"[^A-Za-z0-9]", "", str(val)).upper()
    # Грузинский стандарт: 2 буквы + 3 цифры + 2 буквы (7 символов)
    if len(s) == 7 and re.match(r"^[A-Z]{2}\d{3}[A-Z]{2}$", s):
        return s
    return ""


def clean_tp(val: any) -> str:
    """Очищает номер техпаспорта (например, 'Bva1707845 ' -> 'BVA1707845')."""
    if pd.isna(val):
        return ""
    s = re.sub(r"[^A-Za-z0-9]", "", str(val)).upper()
    # Техпаспорта в Грузии: 2-3 буквы (BVA, AJA, DVA, EVA, KVA, AQA, AM, AJ, BV) + 6-8 цифр
    if re.match(r"^[A-Z]{2,3}\d{6,8}$", s):
        return s
    return ""


def load_fleet_from_excel(filepath: str = EXCEL_SOURCE_PATH) -> Dict[str, Dict]:
    """
    Читает вкладку 'Номера' из файла Брони.xlsx.
    Возвращает словарь:
    {
      "PG397PP": {"name": "Эскейп Макса 397", "tps": ["BVA1627475", "AJA4346779"]},
      ...
    }
    """
    if not os.path.exists(filepath):
        logging.warning(f"Файл {filepath} не найден.")
        return {}

    try:
        df = pd.read_excel(filepath, sheet_name="Номера")
    except Exception as e:
        logging.error(f"Ошибка при чтении листа 'Номера' из {filepath}: {e}")
        return {}

    cars: Dict[str, Dict] = {}

    # Проход 1: Основная таблица справа (колонки 4..8: Название, Госномер, ТП1, ТП2, ТП3)
    # + резервные/старые техпаспорта из колонок 0 и 1 той же строки
    for idx, row in df.iterrows():
        ncols = len(row)
        raw_name = str(row.iloc[4]).strip() if ncols > 4 and pd.notna(row.iloc[4]) else ""
        plate_main = clean_plate(row.iloc[5]) if ncols > 5 else ""

        if plate_main:
            tps_ordered = []
            # Сначала берем актуальные техпаспорта справа (колонки 6, 7, 8)
            for c_idx in [6, 7, 8]:
                if ncols > c_idx:
                    tp = clean_tp(row.iloc[c_idx])
                    if tp and tp not in tps_ordered:
                        tps_ordered.append(tp)

            # Затем добавляем старые/запасные техпаспорта слева (колонки 0, 1)
            for c_idx in [0, 1]:
                if ncols > c_idx:
                    tp = clean_tp(row.iloc[c_idx])
                    if tp and tp not in tps_ordered:
                        tps_ordered.append(tp)

            if plate_main not in cars:
                cars[plate_main] = {
                    "name": raw_name if raw_name and raw_name.lower() != "nan" else f"Авто {plate_main}",
                    "tps": tps_ordered,
                }
            else:
                if raw_name and cars[plate_main]["name"].startswith("Авто "):
                    cars[plate_main]["name"] = raw_name
                for tp in tps_ordered:
                    if tp not in cars[plate_main]["tps"]:
                        cars[plate_main]["tps"].append(tp)

    # Проход 2: Дополнительный список машин слева внизу (колонка 0 = Госномер типа NZ-001-GU, колонка 1 = ТП)
    for idx, row in df.iterrows():
        ncols = len(row)
        plate_left = clean_plate(row.iloc[0]) if ncols > 0 else ""
        tp_left = clean_tp(row.iloc[1]) if ncols > 1 else ""

        if plate_left and tp_left:
            if plate_left not in cars:
                cars[plate_left] = {
                    "name": f"Авто {plate_left}",
                    "tps": [tp_left],
                }
            else:
                if tp_left not in cars[plate_left]["tps"]:
                    cars[plate_left]["tps"].append(tp_left)

    return cars


# =====================================================================
# 2. БАЗА СОСТОЯНИЯ ШТРАФОВ (fines_state.json) ДЛЯ ОТСЛЕЖИВАНИЯ ДИНАМИКИ
# =====================================================================
def load_state() -> dict:
    if os.path.exists(STATE_DB_PATH):
        try:
            with open(STATE_DB_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logging.error(f"Ошибка чтения {STATE_DB_PATH}: {e}")
    return {
        "last_check_time": None,
        "last_total_debt": 0.0,
        "known_fines": {},  # protocol_id -> dict со сведениями о штрафе
        "cars_summary": {},  # plate -> dict с итогами по машине
    }


def save_state(state: dict):
    try:
        with open(STATE_DB_PATH, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logging.error(f"Ошибка сохранения {STATE_DB_PATH}: {e}")


# =====================================================================
# 3. МОДУЛЬ ПРОВЕРКИ ШТРАФОВ НА PROTOCOLS.GE (С ПЕРЕБОРОМ ВСЕХ ТП)
# =====================================================================
def is_driver_or_inactive_fine(raw_item: dict | str) -> bool:
    """
    Отсекает:
    1) Серые / неактивные штрафы, выписанные лично на права водителя (personal/driver fine).
    2) Уже оплаченные (paid / გადახდილია) или отмененные штрафы.
    """
    text_repr = str(raw_item).lower()
    # Признаки оплаченного штрафа
    if any(w in text_repr for w in ["paid", "გადახდილი", "оплачен", "status_paid", "inactive", "canceled"]):
        return True
    # Признаки водительского штрафа (выписанного на личный ID / права, который недоступен по техпаспорту)
    if any(w in text_repr for w in ["driver_only", "personal_fine", "disabled", "greyed"]):
        return True
    return False


def calculate_days_and_deadline(issue_date_str: str, explicit_deadline_str: str = "") -> Tuple[int, str, str]:
    """
    Считает, сколько дней прошло или осталось до повышения штрафа.
    Возвращает: (days_left, deadline_date_str, urgency_status)
    urgency_status: 'URGENT_DISCOUNT' (скидка 20% сгорает <= 3 дн),
                    'URGENT_PENALTY' (штраф удвоится <= 5 дн),
                    'OVERDUE' (просрочен!),
                    'NORMAL' (в пределах срока)
    """
    now = datetime.now(TIMEZONE).replace(tzinfo=None)

    # Пытаемся распознать дату выписки или дедлайна (форматы DD.MM.YYYY, YYYY-MM-DD)
    parsed_issue = None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            parsed_issue = datetime.strptime(issue_date_str.strip()[:10], fmt)
            break
        except Exception:
            pass

    parsed_deadline = None
    if explicit_deadline_str:
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                parsed_deadline = datetime.strptime(explicit_deadline_str.strip()[:10], fmt)
                break
            except Exception:
                pass

    if parsed_deadline:
        days_left = (parsed_deadline - now).days
        dl_str = parsed_deadline.strftime("%d.%m.%Y")
        if days_left < 0:
            return days_left, dl_str, "OVERDUE"
        elif days_left <= 5:
            return days_left, dl_str, "URGENT_PENALTY"
        return days_left, dl_str, "NORMAL"

    if parsed_issue:
        days_passed = (now - parsed_issue).days
        # Стандартный срок оплаты в Грузии — 30 дней (10 дней на скидку 20%)
        days_left_30 = 30 - days_passed
        days_left_10 = 10 - days_passed
        dl_30_str = (parsed_issue + pd.Timedelta(days=30)).strftime("%d.%m.%Y")

        if days_left_30 < 0:
            return days_left_30, dl_30_str, "OVERDUE"
        elif days_left_30 <= 5:
            return days_left_30, dl_30_str, "URGENT_PENALTY"
        elif 0 <= days_left_10 <= 3:
            return days_left_30, dl_30_str, "URGENT_DISCOUNT"
        return days_left_30, dl_30_str, "NORMAL"

    return 30, "Не указан", "NORMAL"


async def fetch_fines_single_tp(
    session: aiohttp.ClientSession, plate: str, tp: str
) -> List[Dict]:
    """
    Отправляет запрос на protocols.ge для одной связки (Госномер + Техпаспорт).
    Поддерживает как JSON-ответ API, так и разбор HTML-таблицы.
    Автоматически отсекает серые водительские строки (class='disabled' / 'gray' / без суммы).
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
        "Accept": "application/json, text/html, */*",
        "Referer": "https://protocols.ge/",
    }

    # Поддержка вариантов параметров, которые использует форма protocols.ge
    params_variants = [
        {"plate": plate, "tech_passport": tp},
        {"car_number": plate, "document_number": tp},
        {"number": plate, "passport": tp},
    ]

    found_fines: List[Dict] = []

    for payload in params_variants:
        for method, url in [
            ("GET", "https://protocols.ge/"),
            ("POST", "https://protocols.ge/"),
            ("POST", "https://protocols.ge/api/search"),
        ]:
            try:
                if method == "GET":
                    req_ctx = session.get(url, params=payload, headers=headers, timeout=12)
                else:
                    req_ctx = session.post(url, data=payload, headers=headers, timeout=12)

                async with req_ctx as resp:
                    if resp.status != 200:
                        continue

                    content_type = resp.headers.get("Content-Type", "").lower()
                    body_text = await resp.text()

                    # Вариант А: Сервер вернул JSON
                    if "application/json" in content_type or body_text.strip().startswith(("{", "[")):
                        try:
                            data = json.loads(body_text)
                            items = (
                                data.get("fines")
                                or data.get("protocols")
                                or data.get("data")
                                or (data if isinstance(data, list) else [])
                            )
                            for item in items:
                                if not isinstance(item, dict) or is_driver_or_inactive_fine(item):
                                    continue
                                raw_amt = str(item.get("amount") or item.get("penalty") or item.get("sum") or "0")
                                clean_amt = re.sub(r"[^\d.]", "", raw_amt.replace(",", "."))
                                amount = float(clean_amt) if clean_amt else 0.0
                                if amount <= 0:
                                    continue

                                prot_id = str(
                                    item.get("protocol_number")
                                    or item.get("receipt_number")
                                    or item.get("id")
                                    or f"{plate}_{amount}_{item.get('date', '')}"
                                ).strip()
                                issue_date = str(item.get("date") or item.get("issue_date") or datetime.now(TIMEZONE).strftime("%d.%m.%Y"))
                                dl_date = str(item.get("deadline") or item.get("due_date") or "")
                                days_left, dl_str, urg = calculate_days_and_deadline(issue_date, dl_date)

                                found_fines.append({
                                    "protocol_id": prot_id,
                                    "plate": plate,
                                    "tp_used": tp,
                                    "amount": amount,
                                    "issue_date": issue_date,
                                    "deadline": dl_str,
                                    "days_left": days_left,
                                    "urgency": urg,
                                    "description": str(item.get("article") or item.get("description") or "Штраф по авто"),
                                })
                            if found_fines:
                                return found_fines
                        except Exception:
                            pass

                    # Вариант Б: Сервер вернул HTML-страницу с таблицей протоколов
                    soup = BeautifulSoup(body_text, "html.parser")
                    rows = soup.find_all("tr")
                    for tr in rows:
                        tr_classes = " ".join(tr.get("class", [])).lower()
                        tr_style = (tr.get("style") or "").lower()
                        # Пропускаем зачеркнутые / серые (водительские) строки
                        if any(k in tr_classes or k in tr_style for k in ["disabled", "gray", "grey", "muted", "line-through", "driver", "paid"]):
                            continue

                        cols = [td.get_text(" ", strip=True) for td in tr.find_all(["td", "th"])]
                        if len(cols) < 2:
                            continue
                        row_joined = " | ".join(cols)
                        if is_driver_or_inactive_fine(row_joined):
                            continue

                        # Ищем сумму в лари (GEL / ლ / число с точкой)
                        amt_match = re.search(r"(\d{1,4}(?:[.,]\d{1,2})?)\s*(?:GEL|ლ|лари|ল)?", row_joined, re.I)
                        # Ищем номер квитанции/протокола (обычно от 6 до 12 символов/цифр)
                        prot_match = re.search(r"\b([A-ZА-Я]{0,3}\d{6,12})\b", row_joined, re.I)
                        date_match = re.search(r"\b(\d{2}\.\d{2}\.\d{4}|\d{4}-\d{2}-\d{2})\b", row_joined)

                        if amt_match and prot_match:
                            amount = float(amt_match.group(1).replace(",", "."))
                            if amount <= 0:
                                continue
                            prot_id = prot_match.group(1).upper()
                            issue_date = date_match.group(1) if date_match else datetime.now(TIMEZONE).strftime("%d.%m.%Y")
                            days_left, dl_str, urg = calculate_days_and_deadline(issue_date)

                            found_fines.append({
                                "protocol_id": prot_id,
                                "plate": plate,
                                "tp_used": tp,
                                "amount": amount,
                                "issue_date": issue_date,
                                "deadline": dl_str,
                                "days_left": days_left,
                                "urgency": urg,
                                "description": "Активный протокол",
                            })
                    if found_fines:
                        return found_fines

            except Exception as e:
                logging.debug(f"Попытка {method} {url} для {plate}/{tp}: {e}")

    return found_fines


async def check_car_all_tps(
    session: aiohttp.ClientSession, plate: str, tps: List[str]
) -> Tuple[List[Dict], List[str]]:
    """
    Перебирает ВСЕ техпаспорта машины (основные и старые).
    Собирает уникальные штрафы по номеру протокола (чтобы один штраф не посчитался дважды).
    """
    unique_fines: Dict[str, Dict] = {}
    checked_tps: List[str] = []

    for tp in tps:
        checked_tps.append(tp)
        fines = await fetch_fines_single_tp(session, plate, tp)
        for f in fines:
            pid = f["protocol_id"]
            if pid not in unique_fines:
                unique_fines[pid] = f
        await asyncio.sleep(0.25)  # Мягкая задержка, чтобы сайт не блокировал запросы

    return list(unique_fines.values()), checked_tps


# =====================================================================
# 4. ГЕНЕРАЦИЯ ЧИСТОЙ EXCEL-ТАБЛИЦЫ (БЕЗ МУСОРА, ТОЛЬКО ЦИФРЫ И ДАТЫ)
# =====================================================================
def generate_clean_excel_report(cars_summary: Dict[str, Dict], total_debt: float, diff_debt: float) -> str:
    """
    Создает красивую таблицу 'Отчет_штрафы_Батуми.xlsx' с 2 вкладками:
    1) 'Только со штрафами' — машины, где долг > 0 (отсортированы по убыванию долга).
    2) 'Весь автопарк (166 авто)' — полная сводка по всем машинам и их техпаспортам.
    """
    rows_all = []
    rows_fines_only = []

    for plate, info in cars_summary.items():
        row_dict = {
            "Госномер": plate,
            "Автомобиль": info["name"],
            "Кол-во штрафов (шт)": info["fines_count"],
            "Сумма долга (GEL)": info["total_amount"],
            "Ближайший срок оплаты": info["nearest_deadline"],
            "Статус срочности": info["status_label"],
            "Номера квитанций": ", ".join(info["protocol_ids"]) if info["protocol_ids"] else "—",
            "Проверенные техпаспорта": ", ".join(info["tps"]),
            "Активный ТП (где найден штраф)": info.get("active_tp") or "—",
        }
        rows_all.append(row_dict)
        if info["total_amount"] > 0:
            rows_fines_only.append(row_dict)

    df_all = pd.DataFrame(rows_all).sort_values(by=["Сумма долга (GEL)", "Госномер"], ascending=[False, True])
    if rows_fines_only:
        df_fines = pd.DataFrame(rows_fines_only).sort_values(by=["Сумма долга (GEL)"], ascending=False)
    else:
        df_fines = pd.DataFrame(columns=[
            "Госномер", "Автомобиль", "Кол-во штрафов (шт)", "Сумма долга (GEL)",
            "Ближайший срок оплаты", "Статус срочности", "Номера квитанций",
            "Проверенные техпаспорта", "Активный ТП (где найден штраф)"
        ])

    # Добавляем итоговую строку сверху или снизу
    summary_row = pd.DataFrame([{
        "Госномер": "ИТОГО ПО ПАРКУ",
        "Автомобиль": f"Динамика за сутки: {'+' if diff_debt >= 0 else ''}{diff_debt:.0f} GEL",
        "Кол-во штрафов (шт)": int(df_all["Кол-во штрафов (шт)"].sum()) if not df_all.empty else 0,
        "Сумма долга (GEL)": total_debt,
        "Ближайший срок оплаты": datetime.now(TIMEZONE).strftime("%d.%m.%Y %H:%M"),
        "Статус срочности": "",
        "Номера квитанций": "",
        "Проверенные техпаспорта": "",
        "Активный ТП (где найден штраф)": "",
    }])

    df_fines_out = pd.concat([summary_row, df_fines], ignore_index=True)
    df_all_out = pd.concat([summary_row, df_all], ignore_index=True)

    with pd.ExcelWriter(REPORT_EXCEL_PATH, engine="openpyxl") as writer:
        df_fines_out.to_excel(writer, sheet_name="Только со штрафами", index=False)
        df_all_out.to_excel(writer, sheet_name="Весь автопарк", index=False)

    return REPORT_EXCEL_PATH


# =====================================================================
# 5. ГЛАВНЫЙ ЦИКЛ ПОЛНОЙ ПРОВЕРКИ ПАРКА И ОТПРАВКИ УВЕДОМЛЕНИЙ
# =====================================================================
async def run_full_fleet_check(notify_chat_id: Optional[int | str] = None, notify_thread_id: Optional[int] = None, is_scheduled: bool = False):
    global is_checking_now
    if is_checking_now:
        if notify_chat_id:
            await bot.send_message(
                chat_id=notify_chat_id,
                message_thread_id=notify_thread_id,
                text="⏳ Проверка уже выполняется в данный момент, подождите окончания...",
            )
        return

    is_checking_now = True
    try:
        fleet = load_fleet_from_excel()
        if not fleet:
            if notify_chat_id:
                await bot.send_message(
                    chat_id=notify_chat_id,
                    message_thread_id=notify_thread_id,
                    text="⚠️ Файл `Брони.xlsx` не найден или вкладка `Номера` пуста! Просто отправьте файл `Брони.xlsx` сюда в чат.",
                    parse_mode="Markdown",
                )
            return

        prev_state = load_state()
        prev_known_fines: Dict[str, Dict] = prev_state.get("known_fines", {})
        prev_total_debt: float = float(prev_state.get("last_total_debt", 0.0))

        new_known_fines: Dict[str, Dict] = {}
        cars_summary: Dict[str, Dict] = {}
        newly_appeared_fines: List[Dict] = []
        urgent_fines: List[Dict] = []

        total_debt = 0.0
        total_fines_count = 0

        async with aiohttp.ClientSession() as session:
            for plate, car_data in fleet.items():
                car_name = car_data["name"]
                tps = car_data["tps"]

                fines, checked_tps = await check_car_all_tps(session, plate, tps)

                car_amount = sum(f["amount"] for f in fines)
                total_debt += car_amount
                total_fines_count += len(fines)

                nearest_dl = "—"
                status_label = "✅ Чисто"
                active_tp_used = ""

                if fines:
                    # Сортируем по остатку дней до повышения
                    fines.sort(key=lambda x: x["days_left"])
                    nearest_dl = f"{fines[0]['deadline']} ({fines[0]['days_left']} дн.)"
                    active_tp_used = fines[0]["tp_used"]

                    if any(f["urgency"] == "OVERDUE" for f in fines):
                        status_label = "🔴 ПРОСРОЧЕН!"
                    elif any(f["urgency"] == "URGENT_PENALTY" for f in fines):
                        status_label = "🔥 Срочно (скоро удвоение!)"
                    elif any(f["urgency"] == "URGENT_DISCOUNT" for f in fines):
                        status_label = "⚠️ Сгорает скидка 20%"
                    else:
                        status_label = "🟡 Есть штраф"

                for f in fines:
                    pid = f["protocol_id"]
                    f_record = {**f, "car_name": car_name}
                    new_known_fines[pid] = f_record

                    if pid not in prev_known_fines:
                        newly_appeared_fines.append(f_record)

                    if f["urgency"] in ("OVERDUE", "URGENT_PENALTY", "URGENT_DISCOUNT"):
                        urgent_fines.append(f_record)

                cars_summary[plate] = {
                    "name": car_name,
                    "tps": checked_tps,
                    "active_tp": active_tp_used,
                    "fines_count": len(fines),
                    "total_amount": car_amount,
                    "nearest_deadline": nearest_dl,
                    "status_label": status_label,
                    "protocol_ids": [f["protocol_id"] for f in fines],
                }

        diff_debt = total_debt - prev_total_debt
        now_str = datetime.now(TIMEZONE).strftime("%d.%m.%Y %H:%M")

        # Сохраняем новое состояние
        new_state = {
            "last_check_time": now_str,
            "last_total_debt": total_debt,
            "known_fines": new_known_fines,
            "cars_summary": cars_summary,
        }
        save_state(new_state)

        # Генерируем чистый Excel-отчет
        report_path = generate_clean_excel_report(cars_summary, total_debt, diff_debt)

        # Формируем компактное сообщение без спама
        cars_with_fines_cnt = sum(1 for c in cars_summary.values() if c["total_amount"] > 0)
        diff_sign = "+" if diff_debt > 0 else ""
        diff_emoji = "📈" if diff_debt > 0 else ("📉" if diff_debt < 0 else "➖")

        summary_lines = [
            f"📊 **СВОДКА ПО ШТРАФАМ АВТОПАРКА ({now_str})**",
            f"🚗 Проверено машин: `{len(fleet)}` (учтено техпаспортов: `{sum(len(x['tps']) for x in fleet.values())}`)",
            f"💰 **Общая сумма штрафов:** `{total_debt:.0f} GEL`",
            f"{diff_emoji} **Динамика с прошлой проверки:** `{diff_sign}{diff_debt:.0f} GEL`",
            f"🔴 **Машин со штрафами:** `{cars_with_fines_cnt}` из `{len(fleet)}` (всего протоколов: `{total_fines_count}`)",
        ]

        # Блок 1: Новые штрафы (если появились)
        if newly_appeared_fines:
            summary_lines.append("\n🚨 **НОВЫЕ ШТРАФЫ:**")
            for nf in newly_appeared_fines[:15]:
                summary_lines.append(
                    f"• `{nf['plate']}` ({nf['car_name']}) — **{nf['amount']:.0f} GEL** "
                    f"(Квитанция `{nf['protocol_id']}`, ТП: `{nf['tp_used']}`, до `{nf['deadline']}`)"
                )
            if len(newly_appeared_fines) > 15:
                summary_lines.append(f"...и еще {len(newly_appeared_fines) - 15} шт. (см. в таблице ниже)")

        # Блок 2: Горящие штрафы (истекает срок)
        if urgent_fines:
            summary_lines.append("\n⏳ **ГОРЯЩИЕ СРОКИ (ОПЛАТИТЬ СРОЧНО):**")
            for uf in urgent_fines[:15]:
                summary_lines.append(
                    f"• `{uf['plate']}` ({uf['car_name']}) — **{uf['amount']:.0f} GEL** | "
                    f"Срок до `{uf['deadline']}` (осталось `{uf['days_left']} дн.`)"
                )

        summary_lines.append("\n📎 *Полная таблица по каждой машине прикреплена файлом ниже.*")
        final_text = "\n".join(summary_lines)

        target_chat = notify_chat_id or get_target_group_id()
        target_thread = notify_thread_id if notify_chat_id else get_target_thread_id()

        if target_chat:
            doc_file = FSInputFile(report_path)
            await bot.send_document(
                chat_id=target_chat,
                message_thread_id=target_thread,
                document=doc_file,
                caption=final_text,
                parse_mode="Markdown",
            )

    except Exception as e:
        logging.error(f"Ошибка в цикле проверки: {e}")
        if notify_chat_id:
            await bot.send_message(
                chat_id=notify_chat_id,
                message_thread_id=notify_thread_id,
                text=f"⚠️ Произошла ошибка при проверке: `{e}`",
                parse_mode="Markdown",
            )
    finally:
        is_checking_now = False


# =====================================================================
# 6. КНОПКИ И КОМАНДЫ БОТА В TELEGRAM
# =====================================================================
def main_keyboard() -> ReplyKeyboardMarkup:
    kb = [
        [KeyboardButton(text="🔍 Проверить штрафы сейчас")],
        [KeyboardButton(text="📊 Скачать таблицу штрафов"), KeyboardButton(text="📈 Общая сводка и баланс")],
        [KeyboardButton(text="🚗 Список машин и ТП"), KeyboardButton(text="ℹ️ Инструкция и ID чата")],
    ]
    return ReplyKeyboardMarkup(keyboard=kb, resize_keyboard=True)


@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    fleet = load_fleet_from_excel()
    total_tps = sum(len(v["tps"]) for v in fleet.values())
    await message.answer(
        "👋 **Привет! Это бот мониторинга штрафов автопарка (Батуми).**\n\n"
        f"🚗 **Машин в базе (`Брони.xlsx`):** `{len(fleet)}`\n"
        f"🪪 **Всего привязанных техпаспортов (включая старые):** `{total_tps}`\n\n"
        "📌 **Что умеет бот:**\n"
        "• Автоматически раз в сутки (и каждые 4 часа для срочных штрафов) проверяет все машины по всем старым и новым техпаспортам.\n"
        "• Игнорирует чужие/серые водительские штрафы.\n"
        "• Формирует чистую Excel-таблицу без спама и считает ежедневную динамику долга (`+` / `-` GEL).\n"
        "• Если нужно обновить список машин — просто скиньте свежий файл `Брони.xlsx` прямо в этот чат!",
        reply_markup=main_keyboard(),
        parse_mode="Markdown",
    )


@dp.message(F.text == "ℹ️ Инструкция и ID чата")
@dp.message(Command("chatid"))
async def show_chat_info(message: types.Message):
    chat_id = message.chat.id
    thread_id = message.message_thread_id
    await message.answer(
        "ℹ️ **Данные текущего чата (для настройки Railway Variables):**\n\n"
        f"• **ID этого чата (`TARGET_GROUP_ID`):** `{chat_id}`\n"
        f"• **ID этой вкладки/темы (`TARGET_THREAD_ID`):** `{thread_id if thread_id else 'Нет (общий чат)'}`\n\n"
        "💡 *Если вы хотите, чтобы бот присылал уведомления о новых и горящих штрафах в конкретную вкладку вашей рабочей группы, напишите команду `/chatid` прямо внутри нужной вкладки группы и скопируйте оба числа в переменные Railway.*",
        parse_mode="Markdown",
    )


@dp.message(F.text == "🚗 Список машин и ТП")
async def show_fleet_stats(message: types.Message):
    fleet = load_fleet_from_excel()
    if not fleet:
        await message.answer("⚠️ База машин пуста. Отправьте файл `Брони.xlsx` в чат.")
        return

    multi_tp_cars = [f"`{p}` ({d['name']}): {', '.join(d['tps'])}" for p, d in fleet.items() if len(d["tps"]) > 1]
    text = (
        f"🚗 **Всего уникальных автомобилей в базе:** `{len(fleet)}`\n"
        f"🪪 **Машин с несколькими техпаспортами (старый + новый):** `{len(multi_tp_cars)}`\n\n"
        f"**Примеры связок (первые 15):**\n" + "\n".join(multi_tp_cars[:15]) +
        "\n\n🔍 *Для проверки конкретной машины введите команду:* `/car PG397PP`"
    )
    await message.answer(text, parse_mode="Markdown")


@dp.message(Command("car"))
async def check_single_car_cmd(message: types.Message):
    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("⚠️ Укажите госномер машины, например: `/car PG397PP`", parse_mode="Markdown")
        return

    plate = clean_plate(parts[1])
    fleet = load_fleet_from_excel()
    if plate not in fleet:
        await message.answer(f"⚠️ Машина `{plate}` не найдена во вкладке 'Номера'.", parse_mode="Markdown")
        return

    car_info = fleet[plate]
    wait_msg = await message.answer(
        f"🔍 Проверяю `{plate}` ({car_info['name']}) по техпаспортам: `{', '.join(car_info['tps'])}`...",
        parse_mode="Markdown",
    )

    async with aiohttp.ClientSession() as session:
        fines, _ = await check_car_all_tps(session, plate, car_info["tps"])

    if not fines:
        await wait_msg.edit_text(
            f"✅ **Машина `{plate}` ({car_info['name']}):**\n"
            f"Проверены техпаспорта: `{', '.join(car_info['tps'])}`\n"
            f"Активных штрафов на автомобиле нет!",
            parse_mode="Markdown",
        )
    else:
        total_c = sum(f["amount"] for f in fines)
        lines = [
            f"🔴 **Машина `{plate}` ({car_info['name']})**",
            f"Найдено штрафов: `{len(fines)}` на сумму **{total_c:.0f} GEL**\n",
        ]
        for f in fines:
            lines.append(
                f"• Квитанция `{f['protocol_id']}`: **{f['amount']:.0f} GEL** "
                f"(от `{f['issue_date']}`, оплатить до `{f['deadline']}`, ТП: `{f['tp_used']}`)"
            )
        await wait_msg.edit_text("\n".join(lines), parse_mode="Markdown")


@dp.message(F.text == "📈 Общая сводка и баланс")
async def show_current_balance(message: types.Message):
    state = load_state()
    if not state.get("last_check_time"):
        await message.answer("ℹ️ Проверка еще ни разу не запускалась. Нажмите **«🔍 Проверить штрафы сейчас»**.")
        return

    cars_sum = state.get("cars_summary", {})
    fines_cars = [(p, d) for p, d in cars_sum.items() if d.get("total_amount", 0) > 0]
    fines_cars.sort(key=lambda x: x[1]["total_amount"], reverse=True)

    lines = [
        f"📈 **ТЕКУЩИЙ БАЛАНС ШТРАФОВ АВТОПАРКА**",
        f"🕒 Последнее обновление: `{state['last_check_time']}`",
        f"💰 **Общая задолженность:** `{state.get('last_total_debt', 0):.0f} GEL`",
        f"🚗 **Машин со штрафами:** `{len(fines_cars)}` из `{len(cars_sum)}`\n",
    ]
    if fines_cars:
        lines.append("**Топ машин по сумме долга:**")
        for p, d in fines_cars[:20]:
            lines.append(f"• `{p}` ({d['name']}): **{d['total_amount']:.0f} GEL** ({d['fines_count']} шт., до {d['nearest_deadline']})")
    else:
        lines.append("✅ На данный момент активных штрафов в базе не числится.")

    await message.answer("\n".join(lines), parse_mode="Markdown")


@dp.message(F.text == "📊 Скачать таблицу штрафов")
async def send_excel_report_cmd(message: types.Message):
    state = load_state()
    if not os.path.exists(REPORT_EXCEL_PATH):
        fleet = load_fleet_from_excel()
        empty_sum = {
            p: {
                "name": d["name"],
                "tps": d["tps"],
                "active_tp": "",
                "fines_count": 0,
                "total_amount": 0.0,
                "nearest_deadline": "—",
                "status_label": "Ожидает проверки",
                "protocol_ids": [],
            }
            for p, d in fleet.items()
        }
        generate_clean_excel_report(empty_sum, 0.0, 0.0)

    doc = FSInputFile(REPORT_EXCEL_PATH)
    await message.answer_document(
        document=doc,
        caption=f"📊 Актуальная таблица штрафов автопарка.\nПоследняя проверка: `{state.get('last_check_time') or 'Только база номеров'}`",
        parse_mode="Markdown",
    )


@dp.message(F.text == "🔍 Проверить штрафы сейчас")
@dp.message(Command("check"))
async def manual_check_cmd(message: types.Message):
    await message.answer("🚀 Запускаю полную проверку всех 166 автомобилей по всем связкам техпаспортов. По окончании пришлю сводку и Excel-файл...")
    await run_full_fleet_check(
        notify_chat_id=message.chat.id,
        notify_thread_id=message.message_thread_id,
        is_scheduled=False,
    )


@dp.message(F.document)
async def handle_new_excel_upload(message: types.Message):
    """Позволяет скинуть свежий файл Брони.xlsx прямо в чат бота для обновления базы машин."""
    doc = message.document
    if not doc.file_name.lower().endswith(".xlsx"):
        return

    wait_msg = await message.answer("⏳ Загружаю и проверяю новый файл базы автомобилей...")
    file_info = await bot.get_file(doc.file_id)
    await bot.download_file(file_info.file_path, EXCEL_SOURCE_PATH)

    fleet = load_fleet_from_excel(EXCEL_SOURCE_PATH)
    total_tps = sum(len(v["tps"]) for v in fleet.values())
    await wait_msg.edit_text(
        f"✅ **Файл `Брони.xlsx` успешно обновлен!**\n"
        f"• Найдено уникальных машин во вкладке `Номера`: `{len(fleet)}`\n"
        f"• Всего связок техпаспортов: `{total_tps}`",
        parse_mode="Markdown",
    )


async def scheduled_daily_job():
    """Ежедневная автоматическая проверка по расписанию."""
    group_id = get_target_group_id()
    thread_id = get_target_thread_id()
    await run_full_fleet_check(notify_chat_id=group_id, notify_thread_id=thread_id, is_scheduled=True)


async def main():
    # Настраиваем автоматический запуск каждый день в 10:00 утра по времени Батуми
    scheduler.add_job(scheduled_daily_job, "cron", hour=10, minute=0)
    scheduler.start()
    logging.info("Бот мониторинга штрафов запущен! Расписание активно.")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
