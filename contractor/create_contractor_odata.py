#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
1C UT11 OData — Створення контрагента з TXT / Excel / JSON / зображень / PDF.
Підтримує: одиночний запис, пакетну обробку, GPT-4o vision для JPEG/PNG/PDF/DOCX.

Використання:
  python create_contractor_odata.py invoice.jpg  --db Vlada2
  python create_contractor_odata.py batch.xlsx   --db Vlada
  python create_contractor_odata.py input.txt    --db Vlada2
  python create_contractor_odata.py --check-only --db Vlada

API ключ: D:\\Project\\tools\\.env  →  OPENAI_API_KEY=sk-...
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
_ensure("openai")
_ensure("PyMuPDF", "fitz")
_ensure("python-docx", "docx")
_ensure("Pillow", "PIL")
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
_fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
_console = logging.StreamHandler(sys.stdout)
_console.setLevel(logging.WARNING)
_console.setFormatter(_fmt)
_file = logging.FileHandler("integration_1c.log", encoding="utf-8")
_file.setLevel(logging.INFO)
_file.setFormatter(_fmt)
logging.basicConfig(level=logging.INFO, handlers=[_console, _file])
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


def _get_catalog_ref_by_name(catalog: str, name: str) -> str:
    """Find Ref_Key in any catalog by Description (case-insensitive)."""
    name_lower = name.strip().lower()
    skip = 0
    while True:
        resp = _get(catalog, params={
            "$top": "200", "$skip": str(skip),
            "$format": "json", "$select": "Ref_Key,Description,DeletionMark",
        })
        items = _extract_value(resp)
        if not isinstance(items, list) or not items:
            break
        for item in items:
            if item.get("DeletionMark"):
                continue
            if str(item.get("Description", "")).strip().lower() == name_lower:
                return item["Ref_Key"]
        if len(items) < 200:
            break
        skip += 200
    logger.warning("%s '%s' not found", catalog, name)
    return ZERO_GUID


def get_warehouse_ref(name: str = "Торговый зал") -> str:
    return _get_catalog_ref_by_name("Catalog_Склады", name)


def get_cash_flow_article_ref(name: str = "Оплата поставщику") -> str:
    return _get_catalog_ref_by_name(
        "Catalog_СтатьиДвиженияДенежныхСредств", name
    )


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
        items = _extract_value(resp)
        if not isinstance(items, list) or not items:
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

# ─────────────────────────── GPT-4o vision extractor ────────────────────────

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".tiff", ".tif", ".bmp", ".webp"}
DOC_EXTS   = {".pdf", ".docx", ".doc"}
VISION_EXTS = IMAGE_EXTS | DOC_EXTS

_GPT_PROMPT = """Ти — асистент для витягування реквізитів українських підприємств з документів.
Витягни реквізити і поверни ТІЛЬКИ валідний JSON без markdown і без пояснень:
{
  "назва": "повна юридична назва підприємства",
  "тип": "ТОВ або ФОП або ПП або null",
  "єдрпоу": "РІВНО 8 цифр для ТОВ/ПП/АТ, або РІВНО 10 цифр для ФОП. Може бути підписано: ЄДРПОУ, КОД, ІПН, РНОКПП. Без пробілів. Якщо не 8 і не 10 — null",
  "інн": "ІПН платника ПДВ 10-12 цифр або null",
  "телефон": "+380... або null",
  "адреса": "юридична адреса або null",
  "рахунки": [
    {"iban": "рівно 29 символів: UA + 27 цифр. Ігноруй будь-який префікс перед UA (N, Ne, P/p тощо). Переліч 27 цифр після UA уважно — не пропускай жодної. або null", "мфо": "рівно 6 цифр або null", "банк": "назва банку або null"}
  ]
}
Якщо рахунків кілька — додай всі в масив рахунки. Якщо рахунку немає — рахунки: [].
ВАЖЛИВО: IBAN = рівно UA + 27 цифр = 29 символів. Якщо бачиш NeUA або NUA — бери тільки з UA.
Якщо поле не знайдено — null. Тільки JSON."""


def _load_api_key(name: str) -> str:
    env_file = Path(r"D:\Project\tools\.env")
    if env_file.exists():
        for line in env_file.read_text(encoding="utf-8").splitlines():
            if line.startswith(f"{name}="):
                return line.split("=", 1)[1].strip()
    return os.environ.get(name, "")


def _load_openai_key() -> str:
    return _load_api_key("OPENAI_API_KEY")


def _gpt_parse(raw: str) -> dict:
    raw = raw.strip()
    if raw.startswith("```"):
        parts = raw.split("```")
        raw = parts[1] if len(parts) > 1 else raw
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


def _gpt_text(text: str, client) -> dict:
    import openai
    resp = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": _GPT_PROMPT},
            {"role": "user", "content": f"Документ:\n{text[:8000]}"},
        ],
        temperature=0, max_tokens=1500,
    )
    return _gpt_parse(resp.choices[0].message.content)


def _gpt_vision(img_bytes: bytes, client, mime: str = "image/png") -> dict:
    import openai, base64 as _b64
    b64 = _b64.b64encode(img_bytes).decode()
    resp = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": _GPT_PROMPT},
            {"role": "user", "content": [
                {"type": "text", "text": "Витягни реквізити з цього документу:"},
                {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
            ]},
        ],
        temperature=0, max_tokens=1500,
    )
    return _gpt_parse(resp.choices[0].message.content)


def _inject_regex_ibans(raw: dict, text: str) -> dict:
    """Replace GPT IBAN values with regex-found ones from raw text (more reliable for digit accuracy)."""
    text_clean = re.sub(r'[\s\-]', '', text)
    found = re.findall(r'UA\d{27}', text_clean)
    if not found:
        return raw
    existing = raw.get("рахунки") or []
    raw["рахунки"] = [
        {"iban": iban,
         "мфо":  existing[i].get("мфо", "") if i < len(existing) else "",
         "банк": existing[i].get("банк", "") if i < len(existing) else ""}
        for i, iban in enumerate(found)
    ]
    return raw


def _claude_or_vision(img_bytes: bytes, mime: str) -> dict:
    """Claude claude-sonnet via OpenRouter — better for long digit strings."""
    or_key = _load_api_key("OPENROUTER_API_KEY")
    if not or_key:
        return {}
    import base64 as _b64
    b64 = _b64.b64encode(img_bytes).decode()
    try:
        resp = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={"Authorization": f"Bearer {or_key}", "Content-Type": "application/json"},
            json={
                "model": "anthropic/claude-sonnet-4-6",
                "messages": [
                    {"role": "system", "content": _GPT_PROMPT},
                    {"role": "user", "content": [
                        {"type": "text", "text": "Витягни реквізити з цього документу:"},
                        {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
                    ]},
                ],
                "max_tokens": 1500, "temperature": 0,
            },
            timeout=60,
        )
        if resp.status_code != 200:
            logger.warning("Claude OR vision HTTP %s: %s", resp.status_code, resp.text[:300])
            return {}
        if not resp.text.strip():
            logger.warning("Claude OR vision: empty response body, HTTP %s", resp.status_code)
            return {}
        data = resp.json()
        content = (data.get("choices") or [{}])[0].get("message", {}).get("content", "")
        if not content:
            logger.warning("Claude OR vision: empty content. Response: %s", str(data)[:300])
            return {}
        try:
            return _gpt_parse(content)
        except Exception as parse_exc:
            logger.warning("Claude OR vision: JSON parse failed (%s). Content: %s", parse_exc, repr(content[:300]))
            return {}
    except Exception as exc:
        logger.warning("Claude OR vision failed: %s", exc)
        return {}


def _gemini_vision(img_bytes: bytes, mime: str) -> dict:
    """Gemini 2.0 Flash — fast fallback for IBAN extraction."""
    gem_key = _load_api_key("GEMINI_API_KEY")
    if not gem_key:
        return {}
    import base64 as _b64
    b64 = _b64.b64encode(img_bytes).decode()
    try:
        resp = requests.post(
            f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={gem_key}",
            headers={"Content-Type": "application/json"},
            json={
                "contents": [{"parts": [
                    {"text": _GPT_PROMPT + "\n\nВитягни реквізити з цього документу:"},
                    {"inline_data": {"mime_type": mime, "data": b64}},
                ]}],
                "generationConfig": {"temperature": 0, "maxOutputTokens": 1500},
            },
            timeout=60,
        )
        data = resp.json()
        if "error" in data:
            logger.warning("Gemini vision error: %s", data["error"])
            return {}
        candidates = data.get("candidates") or []
        if not candidates:
            logger.warning("Gemini vision: no candidates. Response: %s", str(data)[:300])
            return {}
        text = candidates[0].get("content", {}).get("parts", [{}])[0].get("text", "")
        if not text:
            return {}
        return _gpt_parse(text)
    except Exception as exc:
        logger.warning("Gemini vision failed: %s", exc)
        return {}


def _ibans_valid(рахунки: list) -> bool:
    """True if all non-empty IBANs are valid 29-char format."""
    for acc in рахунки:
        if isinstance(acc, dict):
            iban = str(acc.get("iban") or "").replace(" ", "").upper()
            if iban and not re.match(r'^UA\d{27}$', iban):
                return False
    return True


def _merge_ibans(base: dict, retry: dict) -> dict:
    """Replace invalid IBANs in base with valid ones from retry model."""
    base_рахунки = base.get("рахунки") or []
    retry_рахунки = retry.get("рахунки") or []
    if not retry_рахунки:
        return base
    merged = []
    for i, acc in enumerate(base_рахунки):
        acc = dict(acc) if isinstance(acc, dict) else {}
        base_iban = str(acc.get("iban") or "").replace(" ", "").upper()
        m = re.search(r'UA\d{27}', base_iban)
        base_iban = m.group(0) if m else base_iban
        if not re.match(r'^UA\d{27}$', base_iban) and i < len(retry_рахунки):
            r_acc = retry_рахунки[i] if isinstance(retry_рахунки[i], dict) else {}
            r_iban = str(r_acc.get("iban") or "").replace(" ", "").upper()
            rm = re.search(r'UA\d{27}', r_iban)
            r_iban = rm.group(0) if rm else r_iban
            if re.match(r'^UA\d{27}$', r_iban):
                acc["iban"] = r_iban
                logger.info("IBAN corrected by fallback model: %s", r_iban)
        merged.append(acc)
    for i in range(len(base_рахунки), len(retry_рахунки)):
        merged.append(retry_рахунки[i])
    return {**base, "рахунки": merged}


def _retry_vision_ibans(raw: dict, img_bytes: bytes, mime: str) -> dict:
    """Auto-retry with Claude then Gemini if any IBAN is invalid."""
    рахунки = raw.get("рахунки") or []
    if _ibans_valid(рахунки):
        return raw
    logger.info("Invalid IBAN — retrying with Claude claude-sonnet...")
    raw_claude = _claude_or_vision(img_bytes, mime)
    if raw_claude:
        raw = _merge_ibans(raw, raw_claude)
    if not _ibans_valid(raw.get("рахунки") or []):
        logger.info("IBAN still invalid — retrying with Gemini...")
        raw_gemini = _gemini_vision(img_bytes, mime)
        if raw_gemini:
            raw = _merge_ibans(raw, raw_gemini)
    return raw


def extract_from_file(filepath: str) -> Dict[str, str]:
    """GPT-4o витягує реквізити з JPEG/PNG/PDF/DOCX. Повертає нормалізований dict."""
    import openai, io
    from PIL import Image
    import fitz
    import docx as _docx

    api_key = _load_openai_key()
    if not api_key:
        raise ValueError(
            "OPENAI_API_KEY не знайдено. Додай у D:\\Project\\tools\\.env"
        )
    client = openai.OpenAI(api_key=api_key)

    p = Path(filepath)
    ext = p.suffix.lower()

    if ext == ".pdf":
        doc = fitz.open(str(p))
        text = "".join(page.get_text() for page in doc)
        doc.close()
        if len(text.strip()) > 100:
            raw = _gpt_text(text, client)
            raw = _inject_regex_ibans(raw, text)
        else:
            doc = fitz.open(str(p))
            pix = doc[0].get_pixmap(matrix=fitz.Matrix(2, 2))
            img_bytes = pix.tobytes("png")
            doc.close()
            raw = _gpt_vision(img_bytes, client, "image/png")
            raw = _retry_vision_ibans(raw, img_bytes, "image/png")

    elif ext in (".docx", ".doc"):
        document = _docx.Document(str(p))
        text = "\n".join(para.text for para in document.paragraphs)
        for table in document.tables:
            for row in table.rows:
                text += "\n" + " | ".join(cell.text for cell in row.cells)
        raw = _gpt_text(text, client)
        raw = _inject_regex_ibans(raw, text)

    else:
        mime_map = {
            ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
            ".png": "image/png", ".tiff": "image/tiff",
            ".tif": "image/tiff", ".bmp": "image/bmp", ".webp": "image/webp",
        }
        mime = mime_map.get(ext, "image/png")
        img = Image.open(str(p))
        if max(img.size) > 2048:
            img.thumbnail((2048, 2048), Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        img_bytes = buf.getvalue()
        raw = _gpt_vision(img_bytes, client, "image/png")
        raw = _retry_vision_ibans(raw, img_bytes, "image/png")

    # Extract base fields (without accounts)
    base_fields = {k: (v or "") for k, v in raw.items()
                   if v is not None and k not in ("рахунки", "iban", "мфо", "банк")}
    base = _normalize(base_fields)

    # Parse рахунки list (new format) or fall back to single iban/мфо/банк
    рахунки = raw.get("рахунки") or []
    if not рахунки:
        # backward compat: old single-field format
        рахунки = [{"iban": raw.get("iban"), "мфо": raw.get("мфо"), "банк": raw.get("банк")}]

    records = []
    for acc in рахунки:
        if not isinstance(acc, dict):
            continue
        rec = base.copy()
        raw_iban = str(acc.get("iban") or "").replace(" ", "").upper()
        # Extract clean UA+27 from noisy OCR (e.g. "NeUA663...346" → "UA663...346")
        m = re.search(r'UA\d{27}', raw_iban)
        rec["iban"] = m.group(0) if m else raw_iban
        rec["mfo"]  = str(acc.get("мфо")  or acc.get("mfo")  or "")
        rec["bank"] = str(acc.get("банк") or acc.get("bank") or "")
        records.append(rec)

    if not records:
        records = [base]

    logger.info("GPT-4o extracted: name=%s edrpou=%s accounts=%d",
                base.get("name"), base.get("edrpou"), len(records))
    return records


def _open_text(filepath: str):
    """Open text file with BOM-based encoding detection (utf-8, utf-16, cp1251)."""
    with open(filepath, "rb") as f:
        bom = f.read(4)
    if bom[:2] in (b"\xff\xfe", b"\xfe\xff"):
        enc = "utf-16"
    elif bom[:3] == b"\xef\xbb\xbf":
        enc = "utf-8-sig"
    else:
        try:
            with open(filepath, encoding="utf-8", errors="strict") as _tf:
                _tf.read()
            enc = "utf-8"
        except (UnicodeDecodeError, ValueError):
            enc = "cp1251"
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
    # ЄДРПОУ must be exactly 8 or 10 digits — clear if wrong length (OCR artifact)
    edrpou = re.sub(r'\D', '', out.get("edrpou", ""))
    out["edrpou"] = edrpou if len(edrpou) in (8, 10) else ""
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
    if ext in VISION_EXTS:
        return extract_from_file(filepath)
    # TXT/CSV/TSV: спочатку пробуємо структурні формати
    try:
        records = parse_batch_txt(filepath)
        if records:
            return records
    except Exception:
        pass
    try:
        return [parse_single_txt(filepath)]
    except Exception:
        pass
    # Fallback: вільний текст → GPT (повідомлення, рахунки-фактури тощо)
    logger.info("Structured parsers failed, using GPT for free-form text: %s", Path(filepath).name)
    import openai as _openai
    _api_key = _load_openai_key()
    if not _api_key:
        raise ValueError("OPENAI_API_KEY не знайдено. Додай у D:\\Project\\tools\\.env")
    _client = _openai.OpenAI(api_key=_api_key)
    with _open_text(filepath) as _fh:
        _text = _fh.read()
    _raw = _gpt_text(_text, _client)
    _raw = _inject_regex_ibans(_raw, _text)
    # Same as extract_from_file: normalize base fields WITHOUT рахунки
    _base_fields = {k: (v or "") for k, v in _raw.items()
                    if v is not None and k not in ("рахунки", "iban", "мфо", "банк")}
    _base = _normalize(_base_fields)
    # Regex rescue: fill name/edrpou if GPT missed them (free-form messages)
    if not _base.get("name"):
        _m = re.search(r'ФОП\s+([\wЀ-ӿ\'\-]+\s+[\wЀ-ӿ\'\-]+\s+[\wЀ-ӿ\'\-]+)', _text)
        if _m:
            _base["name"] = "ФОП " + _m.group(1).strip()
            _base.setdefault("type", "ФОП")
    if not _base.get("edrpou"):
        _m = re.search(r'(?:ЄДРПОУ|ІПН|РНОКПП|КОД)[:/\s]+(\d{8,10})', _text, re.IGNORECASE)
        if _m:
            _candidate = _m.group(1).strip()
            if len(_candidate) in (8, 10):
                _base["edrpou"] = _candidate
    _validate(_base)
    # Build one record per account (same logic as extract_from_file lines 493-513)
    _рахунки = _raw.get("рахунки") or []
    if not _рахунки:
        _рахунки = [{"iban": _raw.get("iban"), "мфо": _raw.get("мфо"), "банк": _raw.get("банк")}]
    _records = []
    for _acc in _рахунки:
        if not isinstance(_acc, dict):
            continue
        _rec = _base.copy()
        _raw_iban = str(_acc.get("iban") or "").replace(" ", "").upper()
        _mi = re.search(r'UA\d{27}', _raw_iban)
        _rec["iban"] = _mi.group(0) if _mi else _raw_iban
        _rec["mfo"]  = str(_acc.get("мфо") or _acc.get("mfo") or "")
        _rec["bank"] = str(_acc.get("банк") or _acc.get("bank") or "")
        _records.append(_rec)
    if not _records:
        _records = [_base]
    return _records


SUPPORTED_EXTS = {".xlsx", ".xls", ".txt", ".csv", ".tsv", ".json"} | VISION_EXTS


def load_records_from_directory(dirpath: str) -> List[Dict[str, str]]:
    """Читає всі підтримувані файли з каталогу і повертає плоский список записів."""
    all_records = []
    files = sorted(Path(dirpath).iterdir())
    for f in files:
        if not f.is_file() or f.suffix.lower() not in SUPPORTED_EXTS:
            continue
        try:
            recs = load_records(str(f))
            all_records.extend(recs)
            logger.info("Loaded %d record(s) from %s", len(recs), f.name)
        except Exception as exc:
            logger.warning("Skipped %s: %s", f.name, exc)
    return all_records


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


def find_bank_accounts_by_contractor(contractor_ref: str) -> List[str]:
    """Returns normalized IBANs already in 1C for this contractor (dedup guard)."""
    try:
        # Try server-side filter first; some 1C instances return HTTP 500 for $filter
        try:
            resp = _get("Catalog_БанковскиеСчетаКонтрагентов", params={
                "$format": "json", "$top": "200",
                "$select": "НомерСчета,Owner_Key",
                "$filter": f"Owner_Key eq guid'{contractor_ref}'",
            })
        except Exception:
            resp = _get("Catalog_БанковскиеСчетаКонтрагентов", params={
                "$format": "json", "$top": "1000",
                "$select": "НомерСчета,Owner_Key",
            })
        items = _extract_value(resp) or []
        if not isinstance(items, list):
            return []
        return [
            str(item.get("НомерСчета", "")).replace(" ", "").upper()
            for item in items
            if item.get("Owner_Key", "").lower() == contractor_ref.lower()
            and item.get("НомерСчета")
        ]
    except Exception as exc:
        logger.warning("Could not fetch existing bank accounts: %s", exc)
        return []


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
                     partner_ref: str, admin_ref: str) -> str:
    agreement_name = f"Угода з {data['name']}"
    currency_ref      = get_currency_ref("980")
    warehouse_ref     = get_warehouse_ref("Торговый зал")
    cash_flow_ref     = get_cash_flow_article_ref("Оплата поставщику")
    payload: Dict[str, Any] = {
        "Description": agreement_name,
        "Контрагент_Key": contractor_ref,
        "Партнер_Key": partner_ref,
        "Валюта_Key": currency_ref,
        "Склад_Key": warehouse_ref,
        "СтатьяДвиженияДенежныхСредств_Key": cash_flow_ref,
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
    results["Agreement_Ref_Key"] = create_agreement(
        data, contractor_ref, partner_ref, admin_ref
    )
    results["status"] = "created"
    return results


# ─────────────────────────── Interactive prompts ─────────────────────────────

def _prompt_select(prompt: str, count: int):
    """Returns list of 0-based indices, or None if cancelled (input 0)."""
    while True:
        raw = input(prompt).strip()
        if raw == "0":
            return None
        if raw == "":
            return list(range(count))
        try:
            chosen = [int(p.strip()) - 1 for p in raw.split(",") if p.strip()]
            valid  = [i for i in chosen if 0 <= i < count]
            if valid:
                return valid
        except ValueError:
            pass
        print(f"  Введіть числа 1-{count} через кому, Enter=всі, 0=скасувати")


def _ask_contractors_table(new_items: list, skip_items: list, yes: bool,
                           add_items: list = []) -> list:
    total = len(new_items) + len(skip_items) + len(add_items)
    LINE = "=" * 68
    print(f"\n{LINE}")
    print(f"  ЗНАЙДЕНО: {total} запис(ів)")
    print(LINE)
    for idx, c in enumerate(new_items, 1):
        print(f"  [{idx}] {c.get('name','')[:36]:<36} | ЄДРПОУ={c.get('edrpou',''):<12} | → буде створено")
    for c in add_items:
        print(f"  [+] {c.get('name','')[:36]:<36} | ЄДРПОУ={c.get('edrpou',''):<12} | → додати рахунок")
    if (new_items or add_items) and skip_items:
        print(f"  {'- ' * 34}")
    for c in skip_items:
        reason = "немає ЄДРПОУ" if not c.get("edrpou", "").strip() else "вже існує"
        print(f"  [*] {c.get('name','')[:36]:<36} | ЄДРПОУ={c.get('edrpou',''):<12} | {reason}")
    print(LINE)
    print(f"  Буде створено: {len(new_items)}  |  Додати рахунок: {len(add_items)}  |  Пропущено: {len(skip_items)}")
    print(LINE)

    if not new_items:
        if add_items:
            print("  Нових контрагентів немає.\n")
        else:
            print("  Всі записи вже існують.\n")
        return []

    if yes:
        print("  --yes: створюємо всі нові записи автоматично.\n")
        return new_items

    idx = _prompt_select(f"Створити: Enter=всі / 1-{len(new_items)} / 0=скасувати\n> ", len(new_items))
    return [] if idx is None else [new_items[i] for i in idx]


def _iban_ok(iban: str) -> bool:
    return bool(re.match(r'^UA\d{27}$', iban)) if iban else True


def _ask_accounts(name: str, ibans: list) -> list:
    print()
    print(f"  Рахунки {name}:")
    for i, (iban, mfo, bank) in enumerate(ibans, 1):
        warn = f"  ⚠ {len(iban)} символів, очікується 29" if iban and not _iban_ok(iban) else ""
        print(f"    [{i}] {iban or '-':<34} МФО {mfo or '-':<8} {(bank or '-')[:20]}{warn}")
    print()
    idx = _prompt_select(f"Вибрати: Enter=всі / 1-{len(ibans)} / 0=пропустити рахунки\n> ", len(ibans))
    if idx is None:
        return []

    result = []
    for i in idx:
        iban, mfo, bank = ibans[i]
        if iban and not _iban_ok(iban):
            print(f"\n  ⚠ IBAN [{i + 1}] має {len(iban)} символів (очікується 29)")
            raw = input("  Введи правильний IBAN або Enter щоб залишити: ").strip()
            if raw:
                cleaned = raw.replace(" ", "").upper()
                m = re.search(r'UA\d{27}', cleaned)
                iban = m.group(0) if m else cleaned
        result.append((iban, mfo, bank))
    return result


def _ask_create_without_bank(name: str) -> bool:
    print()
    print(f"  Розрахунковий рахунок не знайдено. Створити {name[:40]} без рахунку?")
    print("  Enter=так / 0=пропустити")
    raw = input("> ").strip()
    return raw != "0"


def integrate_file(filepath: str, yes: bool = False) -> List[Dict[str, str]]:
    """Load file or directory, precheck 1C, show table, ask selection, process."""
    if os.path.isdir(filepath):
        records = load_records_from_directory(filepath)
        if not records:
            print(f"  No supported files found in: {filepath}")
            return []
    elif os.path.isfile(filepath):
        records = load_records(filepath)
    else:
        raise FileNotFoundError(f"File or directory not found: {filepath}")

    if not check_connectivity():
        raise ConnectionError(f"No connection to 1C OData: {BASE_URL}")

    logger.info("Loaded %d total records", len(records))

    # Group by EDRPOU — same EDRPOU = same contractor, possibly multiple IBANs
    groups: Dict[str, list] = {}
    for rec in records:
        key = rec.get("edrpou", "").strip()
        groups.setdefault(key, []).append(rec)

    contractors = []
    for edrpou, recs in groups.items():
        # Normalize IBANs: strip spaces, upper-case
        for r in recs:
            r["iban"] = r.get("iban", "").replace(" ", "").upper()
        base = recs[0].copy()
        # Deduplicate IBANs within group (same IBAN from multiple files)
        seen: set = set()
        unique_ibans = []
        for r in recs:
            iban = r.get("iban", "")
            if iban:
                if iban not in seen:
                    seen.add(iban)
                    unique_ibans.append((iban, r.get("mfo", ""), r.get("bank", "")))
            else:
                # No IBAN — always include (two accounts without IBAN may be different)
                unique_ibans.append((iban, r.get("mfo", ""), r.get("bank", "")))
        base["_ibans"] = unique_ibans
        contractors.append(base)

    # Precheck 1C: split into new / add_account (exists+iban) / skip (exists, no iban)
    new_items:  list = []
    skip_items: list = []
    add_items:  list = []
    for c in contractors:
        edrpou = c.get("edrpou", "").strip()
        if not edrpou:
            logger.warning("Skipping record with no EDRPOU: %s", c.get("name", "?"))
            skip_items.append(c)
            continue
        try:
            existing = find_contractor_by_edrpou(edrpou)
        except Exception:
            existing = None
        if not existing:
            new_items.append(c)
        else:
            ibans = c.get("_ibans", [])
            has_iban = any(ib or mf for ib, mf, _ in ibans)
            if has_iban:
                c["_existing_ref"] = existing.get("Ref_Key", "")
                add_items.append(c)
            else:
                skip_items.append(c)

    # Show table + ask which to create
    to_create = _ask_contractors_table(new_items, skip_items, yes=yes, add_items=add_items)

    all_results = [
        {"ContractorName": c.get("name", ""), "EDRPOU": c.get("edrpou", ""), "status": "already_exists"}
        for c in skip_items
    ]

    # Add accounts to existing contractors (ask if 2+ IBANs)
    for c in add_items:
        cref  = c.pop("_existing_ref", "")
        ibans = c.pop("_ibans", [])
        ibans_valid = [(ib, mf, bk) for ib, mf, bk in ibans if ib or mf]

        if not ibans_valid:
            all_results.append({
                "ContractorName": c.get("name", ""),
                "EDRPOU": c.get("edrpou", ""),
                "status": "already_exists",
                "Contractor_Ref_Key": cref + " (EXISTING)",
            })
            continue

        if yes or len(ibans_valid) == 1:
            to_add = ibans_valid
        else:
            to_add = _ask_accounts(c.get("name", ""), ibans_valid)
            if not to_add:
                all_results.append({
                    "ContractorName": c.get("name", ""),
                    "EDRPOU": c.get("edrpou", ""),
                    "status": "already_exists",
                    "Contractor_Ref_Key": cref + " (EXISTING)",
                    "BankAccount_Ref_Key": "SKIPPED by user",
                })
                continue

        existing_ibans = find_bank_accounts_by_contractor(cref)
        added = []
        for ib, mf, bk in to_add:
            if ib and ib.upper() in existing_ibans:
                logger.info("IBAN %s already exists for %s — skipping", ib, c.get("name"))
                continue
            try:
                ref = create_bank_account(
                    {"iban": ib, "mfo": mf, "bank": bk,
                     "name": c.get("name", ""), "edrpou": c.get("edrpou", "")}, cref
                )
                added.append(ref)
                logger.info("Account added to existing %s: %s", c.get("name"), ref)
            except Exception as exc:
                logger.error("Failed to add account to %s: %s", c.get("name"), exc)
        all_results.append({
            "ContractorName": c.get("name", ""),
            "EDRPOU": c.get("edrpou", ""),
            "status": "account_added" if added else "already_exists",
            "Contractor_Ref_Key": cref + " (EXISTING)",
            "BankAccount_Ref_Key": ", ".join(added) if added else "SKIPPED",
        })

    for c in to_create:
        ibans       = c.pop("_ibans", [])
        ibans_valid = [(iban, mfo, bank) for iban, mfo, bank in ibans if iban or mfo]
        extra_ibans: list = []

        if not ibans_valid:
            if not yes and not _ask_create_without_bank(c.get("name", "")):
                continue
            for key in ("iban", "mfo", "bank"):
                c.pop(key, None)
        else:
            if yes:
                c["iban"], c["mfo"], c["bank"] = ibans_valid[0]
                extra_ibans = ibans_valid[1:]
            else:
                chosen = _ask_accounts(c.get("name", ""), ibans_valid)
                if not chosen:
                    for key in ("iban", "mfo", "bank"):
                        c.pop(key, None)
                else:
                    c["iban"], c["mfo"], c["bank"] = chosen[0]
                    extra_ibans = chosen[1:]

        logger.info("── %s ──", c.get("name"))
        try:
            result = integrate_data(c)
            if extra_ibans and not result.get("Contractor_Ref_Key", "").endswith("(EXISTING)"):
                cref = result["Contractor_Ref_Key"]
                result["ExtraBankAccounts"] = [
                    create_bank_account(
                        {"iban": ib, "mfo": mf, "bank": bk,
                         "name": c["name"], "edrpou": c["edrpou"]}, cref
                    )
                    for ib, mf, bk in extra_ibans
                ]
        except Exception as exc:
            logger.error("Error %s: %s", c.get("name"), exc)
            result = {
                "ContractorName": c.get("name", ""),
                "EDRPOU": c.get("edrpou", ""),
                "status": f"ERROR: {exc}",
            }
        all_results.append(result)

    return all_results


# ─────────────────────────── Output ──────────────────────────────────────────

_STATUS_LABEL = {
    "created":        "Створено",
    "already_exists": "Вже існує",
    "account_added":  "Рахунок додано",
}

def _count_accounts(r: Dict) -> int:
    bank = r.get("BankAccount_Ref_Key", "")
    if not bank or "SKIP" in bank:
        return 0
    primary = len([x for x in bank.split(",") if x.strip()])
    extra = len(r.get("ExtraBankAccounts") or [])
    return primary + extra


def print_summary(all_results: List[Dict[str, str]]) -> None:
    created  = sum(1 for r in all_results if r.get("status") == "created")
    added    = sum(1 for r in all_results if r.get("status") == "account_added")
    skipped  = sum(1 for r in all_results if r.get("status") == "already_exists")
    errors   = sum(1 for r in all_results if str(r.get("status","")).startswith("ERROR"))

    LINE = "=" * 68
    print(f"\n{LINE}")
    print(f"  ГОТОВО: {len(all_results)}  |  Створено: {created}  |  "
          f"Рахунків додано: {added}  |  Пропущено: {skipped}  |  Помилок: {errors}")
    print(LINE)

    for r in all_results:
        status = r.get("status", "?")
        name   = r.get("ContractorName", "")[:42]
        edrpou = r.get("EDRPOU", "")
        icon   = "✓" if status in ("created", "already_exists", "account_added") else "✗"
        label  = _STATUS_LABEL.get(status, status)
        accts  = _count_accounts(r)
        acct_s = f"  ({accts} рах.)" if accts else ""
        print(f"  {icon} {name:<42} ЄДРПОУ={edrpou:<12} {label}{acct_s}")

    print(LINE + "\n")


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
    parser.add_argument("--yes", "-y", action="store_true",
                        help="Skip confirmation — create all new records automatically")
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
        all_results = integrate_file(args.input_file, yes=args.yes)
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
