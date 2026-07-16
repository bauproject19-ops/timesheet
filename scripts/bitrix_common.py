import os
import re
import base64
import requests
from datetime import date
from urllib.parse import urlsplit, urlunsplit, quote

WEBHOOK = os.environ["BITRIX_WEBHOOK"]  # напр. https://bau-project.bitrix24.ru/rest/5/xxxxxxxx/
ENTITY_TYPE_ID = 1068
BOT_ID = 995
CLIENT_ID = "timesheet_bot_client_2026"
REPORTS_FOLDER_ID = 918023  # Общий диск -> "Отчеты Учет времени"
FORM_URL = "https://bauproject19-ops.github.io/timesheet/timesheet.html"
TRIP_ENUM_ID = 1527

USER_NAMES = {
    5: "Гохаев Денис",
    681: "Костерева Ольга",
    951: "Архангельский Егор",
    257: "Гайворонский Илья",
    787: "Калинин Дмитрий",
    249: "Пак Александр",
    399: "Сидоров Петр",
}

REMINDER_USER_IDS = [951, 257, 787, 249, 399]

# Нерабочие праздничные и перенесённые дни РФ на 2026 год
# (Постановление Правительства РФ от 24.09.2025 N 1466)
# ВАЖНО: в конце 2026 года нужно обновить на список 2027 года.
HOLIDAYS_2026 = {
    date(2025, 12, 31),
    date(2026, 1, 1), date(2026, 1, 2), date(2026, 1, 3), date(2026, 1, 4),
    date(2026, 1, 5), date(2026, 1, 6), date(2026, 1, 7), date(2026, 1, 8), date(2026, 1, 9),
    date(2026, 2, 23), date(2026, 3, 8), date(2026, 5, 1), date(2026, 5, 9),
    date(2026, 6, 12), date(2026, 11, 4),
}


def is_day_off(d: date) -> bool:
    return d.weekday() >= 5 or d in HOLIDAYS_2026


def call(method: str, payload: dict | None = None):
    r = requests.post(f"{WEBHOOK}{method}.json", json=payload or {}, timeout=30)
    r.raise_for_status()
    data = r.json()
    if "error" in data:
        raise RuntimeError(f"{method} error: {data.get('error_description', data['error'])}")
    return data["result"]


def get_all_items():
    """Тянем все записи смарт-процесса «Учёт времени» (с пагинацией)."""
    items = []
    start = 0
    while True:
        res = call("crm.item.list", {
            "entityTypeId": ENTITY_TYPE_ID,
            "select": ["id", "assignedById", "title", "ufCrm37Hours", "ufCrm37Object", "ufCrm37Place"],
            "start": start,
        })
        batch = res["items"]
        items.extend(batch)
        if len(batch) < 50:
            break
        start += 50
    return items


# ВАЖНО: поле begindate в этом смарт-процессе Bitrix НЕЛЬЗЯ использовать как
# дату записи — платформа принудительно перезаписывает его текущей датой при
# любом crm.item.add/update (подтверждённый баг/ограничение, не лечится).
# Поэтому реальная дата записи (в т.ч. задним/передним числом) хранится
# прямо в начале поля title в формате "ДД.ММ.ГГГГ | ...".
DATE_RE = re.compile(r"^(\d{2})\.(\d{2})\.(\d{4})")
# Для отпуска в title дальше по строке диапазон: "... | Отпуск | ДД.ММ.ГГГГ–ДД.ММ.ГГГГ"
RANGE_RE = re.compile(r"(\d{2})\.(\d{2})\.(\d{4})–(\d{2})\.(\d{2})\.(\d{4})")


def record_date(title: str):
    m = DATE_RE.match(title or "")
    if not m:
        return None
    d, mo, y = m.groups()
    return date(int(y), int(mo), int(d))


def vacation_range(title: str):
    m = RANGE_RE.search(title or "")
    if not m:
        return None
    d1, m1, y1, d2, m2, y2 = m.groups()
    return date(int(y1), int(m1), int(d1)), date(int(y2), int(m2), int(d2))


def vacationing_users_on(target: date, items):
    result = set()
    for it in items:
        if it.get("ufCrm37Object") == "Отпуск":
            rng = vacation_range(it.get("title", ""))
            if rng and rng[0] <= target <= rng[1]:
                result.add(it["assignedById"])
    return result


def _encode_url_path(url: str) -> str:
    """DETAIL_URL от Bitrix содержит непроэкранированные пробелы и кириллицу
    в пути (папка называется "Отчеты Учет времени") — мессенджер Bitrix
    обрывает автолинк на первом же пробеле, и в чате ссылка ведёт на
    несуществующий обрезанный путь (404). Кодируем путь целиком, оставляя
    только "/" как safe-символ, чтобы ссылка стала одним словом без пробелов."""
    parts = urlsplit(url)
    encoded_path = quote(parts.path, safe='/')
    return urlunsplit((parts.scheme, parts.netloc, encoded_path, parts.query, parts.fragment))


def upload_to_disk(filename: str, content_bytes: bytes) -> str:
    b64 = base64.b64encode(content_bytes).decode()
    res = call("disk.folder.uploadfile", {
        "id": REPORTS_FOLDER_ID,
        "data": {"NAME": filename},
        "fileContent": [filename, b64],
    })
    return _encode_url_path(res["DETAIL_URL"])


def send_message(dialog_id: str, message: str):
    call("imbot.message.add", {
        "BOT_ID": BOT_ID,
        "CLIENT_ID": CLIENT_ID,
        "DIALOG_ID": dialog_id,
        "MESSAGE": message,
    })


def send_personal(user_id: int, message: str):
    send_message(str(user_id), message)


def fmt_num(h) -> str:
    h = float(h)
    return str(int(h)) if h.is_integer() else str(round(h, 1)).replace(".", ",")
