#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
1C UT11 OData — Створення контрагента з TXT / Excel / JSON.
Підтримує: одиночний запис і пакетну обробку (кілька рядків).

Використання:
  python создать_контрагента_odata.py input.txt   --db Vlada2
  python создать_контрагента_odata.py batch.xlsx  --db Vlada
  python создать_контрагента_odata.py input.json  --db Vlada2
  python создать_контрагента_odata.py --sample    --db Vlada2
  python создать_контрагента_odata.py --check-only --db Vlada
"""

# ── auto-install deps ────────────────────────────────────────────────────────
import subprocess, sys

def _ensure(pkg, import_as=None):
    try:
        __import__(import_as or pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

_ensure("requests")
_ensure("openpyxl")
# ────────────────────────────────────────────────────────────────────────────

import requests
import json
import logging
import os
import re
import uuid
import base64
import csv
from pathlib import Path
from typing import Optional, Dict, Any, List

# ─────────────────────────── Logging ─────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("integration_1c.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)

# ─────────────────────────── Configuration ───────────────────────────────────
BASE_URL  = "http://localhost/Vlada2/odata/standard.odata"
TIMEOUT   = 60
ZERO_GUID = "00000000-0000-0000-0000-000000000000"


def _make_auth_header(user: str, password: str) -> str:
    token = base64.b64encode(f"{user}:{password}".encode("utf-8")).decode("ascii")
    return f"Basic {token}"


HEADERS = {
    "Content-Type": "application/json",
    "Accept":       "application/json",
    "Authorization": _make_auth_header("Админ", "77556670"),
}


# ─────────────────────────── HTTP helpers ────────────────────────────────────

def _get(endpoint: str, params: Optional[Dict] = None) -> Dict:
    url = f"{BASE_URL}/{endpoint}"
    logger.debug("GET %s params=%s", url, params)
    resp = requests.get(url, headers=HEADERS, params=params, timeout=TIMEOUT)
    resp.raise_for_status()
    return resp.json()


def _post(endpoint: str, payload: Dict) -> Dict:
    url = f"{BASE_URL}/{endpoint}"
    logger.debug("POST %s", url)
    resp = requests.post(
        url,
        headers=HEADERS,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        timeout=TIMEOUT,
    )
    if resp.status_code not in (200, 201):
        logger.error("POST %s -> %s: %s", url, resp.status_code, resp.text[:1000])
    resp.raise_for_status()
    return resp.json()


def _extract_value(response: Dict) -> Any:
    if "d" in response:
        data = response["d"]
        return data.get("results", data)
    return response.get("value", response)


def _extract_single(response: Dict) -> Optional[Dict]:
    val = _extract_value(response)
    if isinstance(val, list):
        return val[0] if val else None
    return val


# ─────────────────────────── Reference lookups ───────────────────────────────

def get_admin_ref() -> str:
    try:
        skip = 0
        while True:
            resp = _get("Catalog_Пользователи", params={
                "$select": "Ref_Key,Description", "$top": "100",
                "$skip": str(skip), "$format": "json",
            })
            items = _extract_value(resp)
            if not isinstance(items, list):
                break
            for item in items:
                if item.get("Description", "").strip() == "Админ":
                    return item["Ref_Key"]
            if len(items) < 100:
                break
            skip += 100
    except Exception as exc:
        logger.warning("Admin scan failed: %s", exc)
    # fallback: first non-empty user
    try:
        resp = _get("Catalog_Пользователи", params={
            "$select": "Ref_Key,Description", "$top": "10", "$format": "json",
        })
        for item in _extract_value(resp) or []:
            if item.get("Description", "") not in ("", "<Не указан>"):
                return item["Ref_Key"]
    except Exception:
        pass
    return ZERO_GUID


def get_currency_ref(code: str = "980") -> str:
    try:
        skip = 0
        while True:
            resp = _get("Catalog_Валюты", params={
                "$select": "Ref_Key,Code", "$top": "200",
                "$skip": str(skip), "$format": "json",
            })
            items = _extract_value(resp)
            if not isinstance(items, list):
                break
            for item in items:
                if str(item.get("Code", "")).strip() == code:
                    return item["Ref_Key"]
            if len(items) < 200:
                break
            skip += 200
    except Exception as exc:
        logger.warning("Currency lookup failed: %s", exc)
    return ZERO_GUID


def get_organization_ref() -> str:
    try:
        resp = _get("Catalog_Организации", params={
            "$select": "Ref_Key,Description", "$top": "1", "$format": "json",
        })
        item = _extract_single(resp)
        if item and "Ref_Key" in item:
            return item["Ref_Key"]
    except Exception as exc:
        logger.warning("Organization lookup failed: %s", exc)
    return ZERO_GUID


def get_bank_klassif_ref(mfo: str) -> str:
    if not mfo:
        return ZERO_GUID
    mfo = str(mfo).strip()
    skip = 0
    while True:
        resp = _get("Catalog_КлассификаторБанков", params={
            "$top": "100", "$skip": str(skip),
            "$format": "json", "$select": "Ref_Key,Code,DeletionMark",
        })
        items = resp.get("value", resp) if isinstance(resp, dict) else resp
        if not items:
            break
        for item in items:
            if item.get("DeletionMark"):
                continue
            if str(item.get("Code", "")).strip() == mfo:
                return item["Ref_Key"]
        if len(items) < 100:
            break
        skip += 100
    logger.warning("Bank MFO %s not found — will use text fields", mfo)
    return ZERO_GUID


# ─────────────────────────── Input parsers ───────────────────────────────────

def _open_text(filepath: str):
    """Open text file with BOM-based encoding detection (utf-8, utf-16, cp1251)."""
    with open(filepath, "rb") as f:
        bom = f.read(4)
    if bom[:2] in (b"\xff\xfe", b"\xfe\xff"):
        enc = "utf-16"
    elif bom[:3] == b"\xef\xbb\xbf":
        enc = "utf-8-sig"
    else:
        enc = "utf-8-sig"
    return open(filepath, encoding=enc, errors="replace", newline="")


FIELD_ALIASES = {
    "назва": "name",        "name": "name",      "найменування": "name",
    "єдрпоу": "edrpou",    "edrpou": "edrpou",  "код": "edrpou",
    "телефон": "phone",     "phone": "phone",
    "тип": "type",          "type": "type",
    "iban": "iban",         "рахунок": "iban",
    "мфо": "mfo",           "mfo": "mfo",
    "банк": "bank",         "bank": "bank",
    "інн": "inn",           "инн": "inn",         "inn": "inn",
    "коротка_назва": "short_name", "short_name": "short_name",
    "адреса": "address",    "адрес": "address",   "address": "address",
}


def _normalize(raw: Dict) -> Dict[str, str]:
    out = {}
    for k, v in raw.items():
        key = FIELD_ALIASES.get(str(k).strip().lower(), str(k).strip().lower())
        out[key] = str(v).strip() if v is not None else ""
    return out


def _validate(data: Dict[str, str]) -> None:
    missing = [f for f in ("name", "edrpou") if not data.get(f)]
    if missing:
        raise ValueError(f"Відсутні обов'язкові поля: {missing}")
    edrpou = data.get("edrpou", "")
    if not re.match(r"^\d{8,10}$", edrpou):
        raise ValueError(f"Невірний формат ЄДРПОУ: '{edrpou}' (8-10 цифр)")
    iban = data.get("iban", "")
    if iban and not re.match(r"^UA\d{27}$", iban.replace(" ", "")):
        logger.warning("IBAN може бути некоректним: '%s'", iban)


def parse_single_txt(filepath: str) -> Dict[str, str]:
    """Один контрагент — формат 'Поле: Значення' або JSON."""
    with _open_text(filepath) as fh:
        content = fh.read().strip()

    if content.startswith("{"):
        try:
            data = _normalize(json.loads(content))
            _validate(data)
            return data
        except json.JSONDecodeError:
            pass

    lines = [l.strip() for l in content.splitlines() if l.strip()]
    data: Dict[str, str] = {}
    colon_found = any(":" in line for line in lines)

    if colon_found and len(lines) >= 2:
        for line in lines:
            if ":" in line:
                key, _, value = line.partition(":")
                mapped = FIELD_ALIASES.get(key.strip().lower())
                if mapped:
                    data[mapped] = value.strip()
    else:
        ordered = ["name", "edrpou", "phone", "type", "iban", "mfo", "bank"]
        if len(lines) == 1:
            parts = lines[0].split("\t") if "\t" in lines[0] else lines[0].split(",")
            for i, key in enumerate(ordered):
                if i < len(parts):
                    data[key] = parts[i].strip()
        else:
            for i, key in enumerate(ordered):
                if i < len(lines):
                    val = lines[i]
                    data[key] = val.split(":", 1)[1].strip() if ":" in val else val

    _validate(data)
    return data


def parse_batch_txt(filepath: str) -> List[Dict[str, str]]:
    """
    Пакетний TXT/CSV/TSV — перший рядок заголовки.
    Роздільник визначається автоматично: TAB або ;
    Пропускає порожні рядки і рядки з #.
    """
    with _open_text(filepath) as fh:
        sample = fh.read(2048)

    delimiter = "\t" if sample.count("\t") >= sample.count(";") else ";"

    records = []
    with _open_text(filepath) as fh:
        reader = csv.DictReader(fh, delimiter=delimiter)
        for i, row in enumerate(reader, start=2):
            if not any(row.values()):
                continue
            raw = {k: (v or "") for k, v in row.items() if k}
            if next(iter(raw.values()), "").startswith("#"):
                continue
            try:
                data = _normalize(raw)
                _validate(data)
                records.append(data)
            except ValueError as exc:
                logger.warning("Рядок %d пропущено: %s", i, exc)
    return records


def parse_excel(filepath: str) -> List[Dict[str, str]]:
    """
    Excel .xlsx — перший рядок заголовки, решта — дані.
    Порожні рядки пропускаються.
    """
    import openpyxl
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Excel файл порожній")

    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    records = []
    for i, row in enumerate(rows[1:], start=2):
        raw = {headers[j]: (str(v).strip() if v is not None else "") for j, v in enumerate(row) if j < len(headers)}
        if not any(raw.values()):
            continue
        try:
            data = _normalize(raw)
            _validate(data)
            records.append(data)
        except ValueError as exc:
            logger.warning("Excel рядок %d пропущено: %s", i, exc)
    wb.close()
    return records


def load_records(filepath: str) -> List[Dict[str, str]]:
    """Визначає формат і повертає список записів (1+)."""
    ext = Path(filepath).suffix.lower()
    if ext in (".xlsx", ".xls"):
        return parse_excel(filepath)
    # TXT/CSV/TSV: спочатку пробуємо пакетний формат
    try:
        records = parse_batch_txt(filepath)
        if records:
            return records
    except Exception:
        pass
    # fallback: одиночний запис
    return [parse_single_txt(filepath)]


# ─────────────────────────── 1C integration steps ───────────────────────────

def find_contractor_by_edrpou(edrpou: str) -> Optional[Dict]:
    logger.info("Перевірка наявності контрагента ЄДРПОУ=%s", edrpou)
    skip, page = 0, 500
    while True:
        try:
            resp = _get("Catalog_Контрагенты", params={
                "$select": "Ref_Key,Description,КодПоЕДРПОУ,Партнер_Key",
                "$top": str(page), "$skip": str(skip), "$format": "json",
            })
        except Exception as exc:
            logger.error("Помилка запиту контрагентів: %s", exc)
            return None
        items = _extract_value(resp)
        if not isinstance(items, list):
            break
        for item in items:
            if str(item.get("КодПоЕДРПОУ", "")).strip() == str(edrpou).strip():
                logger.warning("Контрагент ЄДРПОУ=%s вже існує: %s", edrpou, item.get("Ref_Key"))
                return item
        if len(items) < page:
            break
        skip += page
    return None


def create_partner(data: Dict[str, str]) -> str:
    entity_type = data.get("type", "ТОВ").upper()
    yur_fiz = "ЧастноеЛицо" if entity_type in ("ФОП", "FOP") else "Компания"

    _XML_NS = ('xmlns="http://www.v8.1c.ru/ssl/contactinfo" '
               'xmlns:xs="http://www.w3.org/2001/XMLSchema" '
               'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"')
    contact_info = []
    phone = data.get("phone", "")
    if phone:
        contact_info = [{
            "LineNumber": 1,
            "Тип": "Телефон",
            "Вид_Key": "647f283a-0681-11e8-96fb-2c4d545a248c",
            "Представление": phone,
            "НомерТелефона": phone,
            "ДоступенВРабочееВремя": False,
            "ЗначенияПолей": (
                f'<КонтактнаяИнформация {_XML_NS} Представление="{phone}">'
                f'<Состав xsi:type="НомерТелефона" КодСтраны="" КодГорода="" '
                f'Номер="{phone}" Добавочный=""/></КонтактнаяИнформация>'
            ),
        }]

    short_name = data.get("short_name") or data["name"]
    payload: Dict[str, Any] = {
        "Description": short_name,
        "НаименованиеПолное": data["name"],
        "ЮрФизЛицо": yur_fiz,
        "ЭтоГруппа": False,
        "Поставщик": True,
    }
    if contact_info:
        payload["КонтактнаяИнформация"] = contact_info

    resp = _post("Catalog_Партнеры", payload)
    result = _extract_single(resp) or resp
    ref_key = result.get("Ref_Key") or result.get("d", {}).get("Ref_Key")
    if not ref_key:
        raise ValueError(f"Немає Ref_Key у відповіді створення партнера: {result}")
    logger.info("Партнер створено: %s", ref_key)
    return ref_key


def create_contractor(data: Dict[str, str], partner_ref: str) -> str:
    entity_type = data.get("type", "ТОВ").upper()
    yur_fiz = "ИндивидуальныйПредприниматель" if entity_type in ("ФОП", "FOP") else "ЮрЛицо"
    yur_fiz_legal = "ФизическоеЛицо" if yur_fiz == "ИндивидуальныйПредприниматель" else "ЮридическоеЛицо"

    XML_NS = ('xmlns="http://www.v8.1c.ru/ssl/contactinfo" '
              'xmlns:xs="http://www.w3.org/2001/XMLSchema" '
              'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"')
    contact_info = []
    line_num = 1
    phone = data.get("phone", "")
    if phone:
        contact_info.append({
            "LineNumber": line_num,
            "Тип": "Телефон",
            "Вид_Key": "647f282b-0681-11e8-96fb-2c4d545a248c",
            "Представление": phone,
            "НомерТелефона": phone,
            "ДоступенВРабочееВремя": False,
            "ЗначенияПолей": (
                f'<КонтактнаяИнформация {XML_NS} Представление="{phone}">'
                f'<Состав xsi:type="НомерТелефона" КодСтраны="" КодГорода="" '
                f'Номер="{phone}" Добавочный=""/></КонтактнаяИнформация>'
            ),
        })
        line_num += 1
    address = data.get("address", "")
    if address:
        contact_info.append({
            "LineNumber": line_num,
            "Тип": "Адрес",
            "Вид_Key": "647f282f-0681-11e8-96fb-2c4d545a248c",
            "Представление": address,
            "ЗначенияПолей": (
                f'<КонтактнаяИнформация {XML_NS} Представление="{address}">'
                f'<Состав xsi:type="Адрес"/></КонтактнаяИнформация>'
            ),
        })

    short_name = data.get("short_name") or data["name"]
    payload: Dict[str, Any] = {
        "Description": short_name,
        "НаименованиеПолное": data["name"],
        "КодПоЕДРПОУ": data["edrpou"],
        "ЮрФизЛицо": yur_fiz,
        "ЮридическоеФизическоеЛицо": yur_fiz_legal,
        "Партнер_Key": partner_ref,
        "ЭтоГруппа": False,
    }
    inn = data.get("inn", "").strip()
    if inn:
        payload["ПлательщикНДС"] = True
        payload["ИННПлательщикаНДС"] = inn
    if contact_info:
        payload["КонтактнаяИнформация"] = contact_info

    resp = _post("Catalog_Контрагенты", payload)
    result = _extract_single(resp) or resp
    ref_key = result.get("Ref_Key") or result.get("d", {}).get("Ref_Key")
    if not ref_key:
        raise ValueError(f"Немає Ref_Key у відповіді створення контрагента: {result}")
    logger.info("Контрагент створено: %s", ref_key)
    return ref_key


def create_bank_account(data: Dict[str, str], contractor_ref: str) -> str:
    iban      = data.get("iban", "").replace(" ", "")
    mfo       = data.get("mfo", "")
    bank_name = data.get("bank", "")

    bank_ref     = get_bank_klassif_ref(mfo)
    use_bank_ref = bank_ref != ZERO_GUID
    currency_ref = get_currency_ref("980")

    payload: Dict[str, Any] = {
        "Owner":      contractor_ref,
        "Owner_Type": "StandardODATA.Catalog_Контрагенты",
        "Description": iban or f"Рахунок {contractor_ref[:8]}",
        "НомерСчета": iban,
        "ВидСчета": "Текущий",
        "ВалютаДенежныхСредств_Key": currency_ref,
        "РучноеИзменениеРеквизитовБанка": not use_bank_ref,
    }
    if use_bank_ref:
        payload["Банк_Key"] = bank_ref
    else:
        payload["НаименованиеБанка"] = bank_name
        payload["КодБанка"] = mfo

    resp = _post("Catalog_БанковскиеСчетаКонтрагентов", payload)
    result = _extract_single(resp) or resp
    ref_key = result.get("Ref_Key") or result.get("d", {}).get("Ref_Key")
    if not ref_key:
        raise ValueError(f"Немає Ref_Key у відповіді створення банк.рахунку: {result}")
    logger.info("Банківський рахунок створено: %s", ref_key)
    return ref_key


def create_agreement(data: Dict[str, str], contractor_ref: str,
                     partner_ref: str, admin_ref: str, org_ref: str) -> str:
    agreement_name = f"Угода з {data['name']}"
    currency_ref = get_currency_ref("980")
    payload: Dict[str, Any] = {
        "Description": agreement_name,
        "Контрагент_Key": contractor_ref,
        "Партнер_Key": partner_ref,
        "Организация_Key": org_ref,
        "Валюта_Key": currency_ref,
        "Склад_Key": "4f84bc34-0681-11e8-96fb-2c4d545a248c",
        "СтатьяДвиженияДенежныхСредств_Key": "647f29b7-0681-11e8-96fb-2c4d545a248c",
        "Менеджер_Key": admin_ref,
        "Статус": "Действует",
        "Согласован": True,
        "ХозяйственнаяОперация": "ЗакупкаУПоставщика",
        "ДоступноДляЗакупки": True,
        "РегистрироватьЦеныПоставщика": True,
        "ДоступноДляПродажиКлиентам": False,
        "ПорядокОплаты": "РасчетыВГривнахОплатаВГривнах",
        "ПорядокРасчетов": "ПоНакладным",
        "ЭтоГруппа": False,
    }
    resp = _post("Catalog_СоглашенияСПоставщиками", payload)
    result = _extract_single(resp) or resp
    ref_key = result.get("Ref_Key") or result.get("d", {}).get("Ref_Key")
    if not ref_key:
        raise ValueError(f"Немає Ref_Key у відповіді створення угоди: {result}")
    logger.info("Угода створена: %s", ref_key)
    return ref_key


def check_connectivity() -> bool:
    check_hdrs = {"Authorization": HEADERS["Authorization"], "Accept": "application/json"}
    for endpoint in (f"{BASE_URL}/", BASE_URL):
        try:
            resp = requests.get(endpoint, headers=check_hdrs, timeout=15)
            logger.info("OData GET %s -> HTTP %s", endpoint, resp.status_code)
            if resp.status_code == 200:
                return True
            if resp.status_code in (401, 403):
                logger.error("Auth failed (HTTP %s)", resp.status_code)
                return False
        except requests.ConnectionError as exc:
            logger.error("Немає зв'язку з 1С OData: %s", exc)
            return False
        except Exception as exc:
            logger.error("Перевірка зв'язку: %s", exc)
    return False


# ─────────────────────────── Core integration ────────────────────────────────

def integrate_data(data: Dict[str, str]) -> Dict[str, str]:
    """Обробляє один словник з реквізитами контрагента."""
    results: Dict[str, str] = {
        "ContractorName": data.get("name", ""),
        "EDRPOU": data.get("edrpou", ""),
    }

    existing = find_contractor_by_edrpou(data["edrpou"])
    if existing:
        results["status"] = "already_exists"
        results["Contractor_Ref_Key"] = existing.get("Ref_Key", "") + " (EXISTING)"
        results["Partner_Ref_Key"] = existing.get("Партнер_Key", ZERO_GUID) + " (EXISTING)"
        return results

    partner_ref = create_partner(data)
    results["Partner_Ref_Key"] = partner_ref

    contractor_ref = create_contractor(data, partner_ref)
    results["Contractor_Ref_Key"] = contractor_ref

    if data.get("iban") or data.get("mfo"):
        results["BankAccount_Ref_Key"] = create_bank_account(data, contractor_ref)
    else:
        results["BankAccount_Ref_Key"] = "SKIPPED (no IBAN/MFO)"

    admin_ref = get_admin_ref()
    org_ref   = get_organization_ref()
    results["Agreement_Ref_Key"] = create_agreement(
        data, contractor_ref, partner_ref, admin_ref, org_ref
    )
    results["status"] = "created"
    return results


def integrate_file(filepath: str) -> List[Dict[str, str]]:
    """Завантажує файл і обробляє всі записи."""
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"Файл не знайдено: {filepath}")

    if not check_connectivity():
        raise ConnectionError(f"Немає зв'язку з 1С OData: {BASE_URL}")

    records = load_records(filepath)
    logger.info("Завантажено %d записів з %s", len(records), filepath)

    all_results = []
    for i, data in enumerate(records, start=1):
        logger.info("── Запис %d/%d: %s ──", i, len(records), data.get("name"))
        try:
            result = integrate_data(data)
        except Exception as exc:
            logger.error("Помилка запису %d (%s): %s", i, data.get("name"), exc)
            result = {
                "ContractorName": data.get("name", ""),
                "EDRPOU": data.get("edrpou", ""),
                "status": f"ERROR: {exc}",
            }
        all_results.append(result)
    return all_results


# ─────────────────────────── Output ──────────────────────────────────────────

def print_summary(all_results: List[Dict[str, str]]) -> None:
    print("\n" + "=" * 60)
    print(f"  РЕЗУЛЬТАТ: {len(all_results)} запис(ів)")
    print("=" * 60)
    for i, results in enumerate(all_results, start=1):
        status = results.get("status", "?")
        name   = results.get("ContractorName", "")
        edrpou = results.get("EDRPOU", "")
        icon   = "✓" if status in ("created", "already_exists") else "✗"
        print(f"\n  [{i}] {icon} {name} | ЄДРПОУ={edrpou} | {status}")
        for key, value in results.items():
            if key not in ("ContractorName", "EDRPOU", "status"):
                print(f"      {key:<28}: {value}")
    print("=" * 60 + "\n")


# ─────────────────────────── Sample files ────────────────────────────────────

SAMPLE_TXT = """\
назва: ТОВ Ромашка Плюс
ЄДРПОУ: 12345678
телефон: +380441234567
тип: ТОВ
IBAN: UA213996220000026007233566001
МФО: 399622
банк: АТ КБ ПРИВАТБАНК
"""

SAMPLE_BATCH = """\
назва\tєдрпоу\tтелефон\tтип\tIBAN\tМФО\tбанк
ТОВ Ромашка Плюс\t12345678\t+380441234567\tТОВ\tUA213996220000026007233566001\t399622\tАТ КБ ПРИВАТБАНК
ФОП Іваненко І.І.\t1234567890\t+380501234567\tФОП\t\t\t
"""


# ─────────────────────────── CLI ─────────────────────────────────────────────

def main() -> None:
    import argparse
    global BASE_URL

    parser = argparse.ArgumentParser(
        description="1C UT11 OData — Створення контрагента з TXT / Excel / JSON",
    )
    parser.add_argument("input_file", nargs="?",
                        help="TXT / CSV / TSV / XLSX / JSON файл з реквізитами")
    parser.add_argument("--db", default="",
                        help="Назва бази 1С (Vlada або Vlada2). Альтернатива --base-url")
    parser.add_argument("--base-url", default="",
                        help="Повний OData URL (якщо не використовується --db)")
    parser.add_argument("--user",     default="Админ",    help="Логін 1С")
    parser.add_argument("--password", default="77556670", help="Пароль 1С")
    parser.add_argument("--sample",   action="store_true",
                        help="Створити зразок файлів і вийти")
    parser.add_argument("--check-only", action="store_true",
                        help="Тільки перевірити підключення до 1С")
    parser.add_argument("--debug",    action="store_true", help="Детальний лог")

    args = parser.parse_args()

    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)

    # Визначити BASE_URL
    if args.base_url:
        BASE_URL = args.base_url.rstrip("/")
    elif args.db:
        BASE_URL = f"http://localhost/{args.db}/odata/standard.odata"
    # else залишається дефолтний Vlada2

    HEADERS["Authorization"] = _make_auth_header(args.user, args.password)

    if args.sample:
        Path("sample_single.txt").write_text(SAMPLE_TXT, encoding="utf-8")
        Path("sample_batch.txt").write_text(SAMPLE_BATCH, encoding="utf-8")
        print("Створено: sample_single.txt, sample_batch.txt")
        print(f"Запуск: python {Path(__file__).name} sample_single.txt --db Vlada2")
        sys.exit(0)

    if args.check_only:
        ok = check_connectivity()
        print(f"1С OData ({BASE_URL}): {'OK' if ok else 'НЕДОСТУПНО'}")
        sys.exit(0 if ok else 1)

    if not args.input_file:
        parser.print_help()
        print(f"\nПриклад: python {Path(__file__).name} input.xlsx --db Vlada2\n")
        sys.exit(1)

    try:
        all_results = integrate_file(args.input_file)
        print_summary(all_results)
        errors = [r for r in all_results if r.get("status", "").startswith("ERROR")]
        sys.exit(1 if errors else 0)
    except FileNotFoundError as exc:
        logger.error("Файл не знайдено: %s", exc)
        sys.exit(2)
    except ValueError as exc:
        logger.error("Помилка даних: %s", exc)
        sys.exit(3)
    except ConnectionError as exc:
        logger.error("Помилка підключення: %s", exc)
        sys.exit(4)
    except requests.HTTPError as exc:
        logger.error("HTTP помилка: %s", exc)
        sys.exit(5)
    except Exception as exc:
        logger.exception("Несподівана помилка: %s", exc)
        sys.exit(99)


if __name__ == "__main__":
    main()
