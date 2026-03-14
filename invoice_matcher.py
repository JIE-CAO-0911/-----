#!/usr/bin/env python3
import argparse
import json
import math
import re
import sys
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
from pathlib import Path


DEFAULT_CONFIG = {
    "order_columns": {
        "order_id": "",
        "amount": "",
        "date": "",
        "vendor": "",
        "description": "",
        "status": "",
        "product_link": "",
    },
    "invoice_patterns": {
        "invoice_id": [
            r"invoice\s*(?:no|number)\s*[:]?\\s*([A-Za-z0-9-]{5,})",
            r"no\.?\s*[:]?\\s*([A-Za-z0-9-]{5,})",
        ],
        "amount": [
            r"(?:total|amount\s*due|amount)\s*[:]?\\s*([$0-9,]+(?:\\.[0-9]{2})?)",
            r"(?:价税合计|金额合计|合计|总金额|应付|实付|小写)\\s*[:：]?\\s*([¥￥]?[0-9,]+(?:\\.[0-9]{1,2})?)",
        ],
        "date": [
            r"(?:invoice\s*date|date)\s*[:]?\\s*([0-9]{4}[-/\\.][0-9]{1,2}[-/\\.][0-9]{1,2})",
        ],
        "vendor": [
            r"(?:seller|vendor|merchant|supplier)\s*[:]?\\s*([A-Za-z0-9 &.,-]{3,})",
        ],
    },
    "match_rules": {
        "amount_tolerance": 0.01,
        "amount_tolerance_ratio": 0.0,
        "date_days_tolerance": 7,
        "weights": {"amount": 0.6, "date": 0.2, "vendor": 0.2},
        "min_score": 0.7,
        "allow_duplicate_orders": False,
        "vendor_similarity_threshold": 0.6,
        "merge_multi_item_orders": True,
    },
}


ORDER_CANDIDATES = {
    "order_id": {"orderid", "order", "orderno", "order_no", "order_number", "id"},
    "amount": {"amount", "total", "totalamount", "price", "payamount", "paidamount"},
    "date": {"date", "orderdate", "createdate", "paydate"},
    "vendor": {"vendor", "merchant", "seller", "supplier", "company"},
    "description": {"description", "item", "details", "notes"},
    "status": {"status", "orderstatus", "order_state", "state"},
    "product_link": {
        "link",
        "url",
        "productlink",
        "product_link",
        "itemlink",
        "item_link",
        "商品链接",
        "商品链接地址",
        "商品网址",
        "商品地址",
        "商品url",
        "商品URL",
        "链接",
    },
}

MODEL_CANDIDATES = {"型号款式", "型号", "规格", "规格型号"}
QTY_CANDIDATES = {"商品数量", "数量", "qty", "quantity"}


NUM_RE = re.compile(r"\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?")
CURRENCY_AMOUNT_RE = re.compile(
    r"(?:¥|￥|RMB|CNY)\s*([0-9][0-9,]*(?:\.\d{1,2})?)"
    r"|([0-9][0-9,]*(?:\.\d{1,2})?)\s*(?:元|人民币)",
    flags=re.IGNORECASE,
)

RAISE_ON_ERROR = False


class MatchError(RuntimeError):
    pass


def die(message):
    if RAISE_ON_ERROR:
        raise MatchError(message)
    print(message, file=sys.stderr)
    sys.exit(1)


def load_json(path):
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def write_json(path, data, force=False):
    path = Path(path)
    if path.exists() and not force:
        die(f"Refusing to overwrite existing file: {path}")
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2)
        fh.write("\n")


def normalize_key(value):
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def normalize_text(value):
    if value is None:
        return ""
    try:
        import pandas as pd

        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    if text.lower() in {"nan", "none"}:
        return ""
    return text


def format_order_id(value):
    if value is None:
        return ""
    text = normalize_text(value)
    if not text:
        return ""
    if isinstance(value, (int,)):
        return str(value)
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if abs(value - round(value)) < 0.000001:
            return str(int(round(value)))
    if text.startswith("0") and re.fullmatch(r"\d+", text):
        return text
    if re.fullmatch(r"\d{15,}", text):
        return text
    try:
        dec = Decimal(text)
    except InvalidOperation:
        return text
    if dec == dec.to_integral():
        return format(dec.quantize(Decimal(1)), "f").split(".")[0]
    return format(dec.normalize(), "f").rstrip("0").rstrip(".")


def pick_column(columns, configured, candidates):
    if configured:
        return configured if configured in columns else None
    normalized = {normalize_key(c): c for c in columns}
    for cand in candidates:
        if cand in normalized:
            return normalized[cand]
    return None


def parse_amount(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "")
    text = re.sub(r"[^0-9.]", "", text)
    if text.count(".") > 1:
        parts = text.split(".")
        text = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return float(text) if text else None
    except ValueError:
        return None


def amount_key(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return round(float(value), 2)
    try:
        return round(float(value), 2)
    except (ValueError, TypeError):
        return None


def amounts_equal(left, right):
    left_key = amount_key(left)
    right_key = amount_key(right)
    if left_key is None or right_key is None:
        return False
    return left_key == right_key


def parse_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        if 59 < float(value) < 60000:
            try:
                return (date(1899, 12, 30) + timedelta(days=float(value)))
            except OverflowError:
                return None
    text = str(value).strip()
    if not text:
        return None
    text = text.replace("T", " ")
    text = re.split(r"\s+", text)[0]
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    match = re.search(r"(20\d{2})\D?(\d{1,2})\D?(\d{1,2})", str(value))
    if match:
        year, month, day = match.groups()
        try:
            return date(int(year), int(month), int(day))
        except ValueError:
            return None
    return None


def normalize_vendor(value):
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^\w\s&.,-]", " ", text, flags=re.UNICODE)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def as_list(value):
    if value is None:
        return []
    if isinstance(value, list):
        return [v for v in value if v]
    if isinstance(value, str):
        return [value]
    return []


def extract_first(text, patterns):
    for pattern in as_list(patterns):
        try:
            match = re.search(pattern, text, flags=re.IGNORECASE)
        except re.error:
            continue
        if match:
            if match.groups():
                return match.group(1).strip()
            return match.group(0).strip()
    return None


def guess_amount(text):
    values = []
    for raw in NUM_RE.findall(text):
        raw_digits = re.sub(r"\D", "", raw)
        if raw_digits and len(raw_digits) >= 8 and "." not in raw:
            continue
        amount = parse_amount(raw)
        if amount is None:
            continue
        if 0 < amount < 1e9:
            values.append(amount)
    if not values:
        return None
    return max(values)


def find_currency_amount(text):
    values = []
    for match in CURRENCY_AMOUNT_RE.finditer(text):
        raw = match.group(1) or match.group(2)
        amount = parse_amount(raw)
        if amount is not None:
            values.append(amount)
    if not values:
        return None
    return max(values)


def guess_date(text):
    match = re.search(r"(20\d{2})\D{0,3}(\d{1,2})\D{0,3}(\d{1,2})", text)
    if match:
        year, month, day = match.groups()
        try:
            return date(int(year), int(month), int(day))
        except ValueError:
            return None
    return None


def guess_vendor(text):
    for line in text.splitlines():
        if re.search(r"\b(seller|vendor|merchant|supplier)\b", line, flags=re.IGNORECASE):
            cleaned = re.split(r"[:\-]", line, maxsplit=1)
            if len(cleaned) > 1:
                return cleaned[1].strip()
            return line.strip()
    return None


def build_parser():
    parser = argparse.ArgumentParser(
        description="Match invoice PDFs to order data from an Excel file."
    )
    parser.add_argument("--orders", help="Path to orders .xlsx file")
    parser.add_argument("--pdf-dir", help="Directory with invoice PDFs")
    parser.add_argument(
        "--config",
        default="invoice_matcher_config.json",
        help="Path to config JSON",
    )
    parser.add_argument(
        "--out", default="match_result.xlsx", help="Output report path"
    )
    parser.add_argument("--recursive", action="store_true", help="Scan PDFs recursively")
    parser.add_argument(
        "--init-config",
        action="store_true",
        help="Write a default config JSON and exit",
    )
    parser.add_argument("--force", action="store_true", help="Overwrite output/config")
    return parser


def load_config(path):
    path = Path(path) if path else None
    if path and path.exists():
        return load_json(path)
    if path and path.name != "invoice_matcher_config.json":
        die(f"Config file not found: {path}")
    return DEFAULT_CONFIG


def resolve_pdf_backend():
    try:
        import pdfplumber  # noqa: F401

        return "pdfplumber"
    except Exception:
        pass
    try:
        from pypdf import PdfReader  # noqa: F401

        return "pypdf"
    except Exception:
        pass
    try:
        from PyPDF2 import PdfReader  # noqa: F401

        return "pypdf"
    except Exception:
        return None


def read_pdf_text(path, backend):
    if backend == "pdfplumber":
        import pdfplumber

        with pdfplumber.open(path) as pdf:
            return "\n".join(page.extract_text() or "" for page in pdf.pages)
    if backend == "pypdf":
        try:
            from pypdf import PdfReader
        except ImportError:
            from PyPDF2 import PdfReader

        reader = PdfReader(path)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    return ""


def find_pdfs(pdf_dir, recursive):
    root = Path(pdf_dir)
    if not root.exists():
        die(f"PDF directory not found: {root}")
    pdfs = []
    if recursive:
        candidates = root.rglob("*")
    else:
        candidates = root.iterdir()
    for path in candidates:
        if path.is_file() and path.suffix.lower() == ".pdf":
            pdfs.append(path)
    return sorted(pdfs)


def load_orders(orders_path, config):
    try:
        import pandas as pd
    except ImportError:
        die("Missing dependency: pandas. Install with `pip install pandas openpyxl`.")

    orders_path = Path(orders_path)
    if not orders_path.exists():
        die(f"Orders file not found: {orders_path}")
    try:
        preview = pd.read_excel(orders_path, nrows=0)
    except Exception as exc:
        die(f"Failed to read orders file: {exc}")

    preview.columns = [str(c).strip() for c in preview.columns]
    columns = list(preview.columns)

    order_cols = config.get("order_columns", {})
    order_id_col = pick_column(
        columns, order_cols.get("order_id"), ORDER_CANDIDATES["order_id"]
    )
    converters = None
    if order_id_col:
        converters = {
            order_id_col: lambda x: "" if x is None else str(x).strip()
        }

    try:
        df = pd.read_excel(orders_path, converters=converters)
    except Exception as exc:
        die(f"Failed to read orders file: {exc}")

    df.columns = [str(c).strip() for c in df.columns]
    columns = list(df.columns)

    amount_col = pick_column(columns, order_cols.get("amount"), ORDER_CANDIDATES["amount"])
    if amount_col is None:
        die("Amount column not found. Set order_columns.amount in config.")

    if order_id_col not in columns:
        order_id_col = pick_column(
            columns, order_cols.get("order_id"), ORDER_CANDIDATES["order_id"]
        )
    date_col = pick_column(columns, order_cols.get("date"), ORDER_CANDIDATES["date"])
    vendor_col = pick_column(columns, order_cols.get("vendor"), ORDER_CANDIDATES["vendor"])
    desc_col = pick_column(
        columns, order_cols.get("description"), ORDER_CANDIDATES["description"]
    )
    link_col = pick_column(
        columns, order_cols.get("product_link"), ORDER_CANDIDATES["product_link"]
    )
    status_col = pick_column(
        columns, order_cols.get("status"), ORDER_CANDIDATES["status"]
    )
    if status_col is None and "订单状态" in columns:
        status_col = "订单状态"

    df["_order_id"] = df[order_id_col].apply(format_order_id) if order_id_col else ""
    df["_amount"] = df[amount_col].apply(parse_amount)
    df["_date"] = df[date_col].apply(parse_date) if date_col else None
    df["_vendor"] = (
        df[vendor_col].apply(normalize_text) if vendor_col else ""
    )
    df["_description"] = (
        df[desc_col].apply(normalize_text) if desc_col else ""
    )
    df["_product_link"] = (
        df[link_col].apply(normalize_text) if link_col else ""
    )
    df["_status"] = (
        df[status_col].apply(normalize_text) if status_col else ""
    )

    df = merge_multi_item_orders(df, config)

    df["_match_invoice_file"] = None
    df["_match_invoice_id"] = None
    df["_match_score"] = None

    return df


def load_orders_multi(orders_paths, config):
    if isinstance(orders_paths, (str, Path)):
        return load_orders(orders_paths, config)

    paths = [str(p).strip() for p in orders_paths if str(p).strip()]
    if not paths:
        die("Orders file list is empty.")

    frames = [load_orders(path, config) for path in paths]
    if len(frames) == 1:
        return frames[0]

    try:
        import pandas as pd
    except ImportError:
        die("Missing dependency: pandas. Install with `pip install pandas openpyxl`.")

    combined = pd.concat(frames, ignore_index=True)
    combined = deduplicate_orders(combined)
    combined["_match_invoice_file"] = None
    combined["_match_invoice_id"] = None
    combined["_match_score"] = None
    return combined


def deduplicate_orders(df):
    if df is None or df.empty:
        return df

    try:
        import pandas as pd
    except ImportError:
        return df

    def series_or_default(column_name, default_value):
        if column_name in df.columns:
            return df[column_name]
        return pd.Series(default_value, index=df.index)

    def norm_amount(value):
        key = amount_key(value)
        if key is None:
            return ""
        return f"{key:.2f}"

    def norm_date(value):
        if isinstance(value, datetime):
            return value.date().isoformat()
        if isinstance(value, date):
            return value.isoformat()
        return ""

    def norm_text(value):
        return normalize_text(value).strip().lower()

    key_df = df.copy()
    key_df["_k_order_id"] = series_or_default("_order_id", "").apply(format_order_id)
    key_df["_k_amount"] = series_or_default("_amount", None).apply(norm_amount)
    key_df["_k_date"] = series_or_default("_date", None).apply(norm_date)
    key_df["_k_vendor"] = series_or_default("_vendor", "").apply(norm_text)
    key_df["_k_desc"] = series_or_default("_description", "").apply(norm_text)
    key_df["_k_link"] = series_or_default("_product_link", "").apply(norm_text)
    key_df["_k_status"] = series_or_default("_status", "").apply(norm_text)

    dedupe_cols = [
        "_k_order_id",
        "_k_amount",
        "_k_date",
        "_k_vendor",
        "_k_desc",
        "_k_link",
        "_k_status",
    ]
    duplicated = key_df.duplicated(subset=dedupe_cols, keep="first")
    if not duplicated.any():
        return df
    return df.loc[~duplicated].reset_index(drop=True)


def merge_multi_item_orders(df, config):
    rules = config.get("match_rules", {})
    if not rules.get("merge_multi_item_orders", True):
        return df
    if "_order_id" not in df.columns:
        return df

    group_id = df["_order_id"].replace("", None).ffill()
    if group_id.isna().all():
        return df
    import pandas as pd

    group_id = group_id.fillna(pd.Series(df.index.astype(str), index=df.index))
    df = df.copy()
    df["_group_id"] = group_id

    model_col = next((c for c in df.columns if c in MODEL_CANDIDATES), None)
    qty_col = next((c for c in df.columns if c in QTY_CANDIDATES), None)

    def build_line_item(row):
        base = normalize_text(row.get("_description", ""))
        if not base:
            return ""
        model = normalize_text(row.get(model_col)) if model_col else ""
        qty = normalize_text(row.get(qty_col)) if qty_col else ""
        if model:
            base = f"{base} ({model})"
        if qty:
            base = f"{base} x{qty}"
        return base

    df["_line_item"] = df.apply(build_line_item, axis=1)

    def first_non_empty(series):
        for value in series:
            if normalize_text(value):
                return value
        return ""

    def first_not_null(series):
        for value in series:
            if value is not None:
                return value
        return None

    def merge_lines(series):
        seen = set()
        merged = []
        for value in series:
            text = normalize_text(value)
            if not text or text in seen:
                continue
            seen.add(text)
            merged.append(text)
        return " | ".join(merged)

    grouped = df.groupby("_group_id", sort=False)
    merged = grouped.agg(
        {
            "_order_id": first_non_empty,
            "_amount": first_not_null,
            "_date": first_not_null,
            "_vendor": first_non_empty,
            "_description": merge_lines,
            "_product_link": first_non_empty,
            "_status": first_non_empty,
            "_line_item": merge_lines,
        }
    )
    merged["_description"] = merged["_line_item"].where(
        merged["_line_item"] != "", merged["_description"]
    )
    merged = merged.drop(columns=["_line_item"])
    merged = merged.reset_index(drop=True)
    return merged


def build_orders_df_from_rows(rows):
    try:
        import pandas as pd
    except ImportError:
        die("Missing dependency: pandas. Install with `pip install pandas openpyxl`.")

    df = pd.DataFrame(rows)
    if df.empty:
        df = pd.DataFrame(
            columns=[
                "order_id",
                "vendor",
                "description",
                "product_link",
                "amount",
                "order_status",
                "order_date",
            ]
        )

    def col_or_empty(name):
        if name in df.columns:
            return df[name]
        return pd.Series([""] * len(df), index=df.index)

    df["_order_id"] = col_or_empty("order_id").apply(format_order_id)
    df["_vendor"] = col_or_empty("vendor").apply(normalize_text)
    df["_description"] = col_or_empty("description").apply(normalize_text)
    df["_amount"] = col_or_empty("amount").apply(parse_amount)
    df["_product_link"] = col_or_empty("product_link").apply(normalize_text)

    if "order_status" in df.columns:
        df["_status"] = df["order_status"].apply(normalize_text)
    else:
        df["_status"] = col_or_empty("status").apply(normalize_text)

    if "order_date" in df.columns:
        df["_date"] = df["order_date"].apply(parse_date)
    elif "date" in df.columns:
        df["_date"] = df["date"].apply(parse_date)
    else:
        df["_date"] = None

    df["_match_invoice_file"] = None
    df["_match_invoice_id"] = None
    df["_match_score"] = None

    return df


def extract_invoice(path, config, backend):
    text = read_pdf_text(path, backend)
    text_norm = " ".join(text.split())
    patterns = config.get("invoice_patterns", {})

    invoice_id = extract_first(text_norm, patterns.get("invoice_id"))

    amount = None
    amount_raw = extract_first(text_norm, patterns.get("amount"))
    if amount_raw:
        amount = parse_amount(amount_raw)
    if amount is None:
        amount = find_currency_amount(text)
    if amount is None:
        amount = guess_amount(text_norm)

    invoice_date = None
    date_raw = extract_first(text_norm, patterns.get("date"))
    if date_raw:
        invoice_date = parse_date(date_raw)
    if invoice_date is None:
        invoice_date = guess_date(text_norm)

    vendor = extract_first(text_norm, patterns.get("vendor"))
    if vendor is None:
        vendor = guess_vendor(text)

    return {
        "invoice_file": str(path),
        "invoice_id": invoice_id or "",
        "amount": amount,
        "date": invoice_date,
        "vendor": vendor or "",
        "text_excerpt": text_norm[:200],
        "match_order_row": None,
        "match_order_id": "",
        "match_score": None,
    }


def score_amount(inv_amt, order_amt, tol_abs, tol_ratio):
    if inv_amt is None or order_amt is None:
        return None
    diff = abs(inv_amt - order_amt)
    if tol_abs is not None and tol_abs > 0 and diff <= tol_abs:
        return 1.0
    if tol_ratio and order_amt:
        if diff / order_amt <= tol_ratio:
            return 1.0
    if tol_abs == 0 or tol_ratio == 0:
        return 1.0 if diff == 0 else 0.0
    return 0.0


def score_date(inv_date, order_date, tol_days):
    if inv_date is None or order_date is None:
        return None
    if not isinstance(inv_date, date) or not isinstance(order_date, date):
        return None
    delta = abs((inv_date - order_date).days)
    if tol_days <= 0:
        return 1.0 if delta == 0 else 0.0
    if delta > tol_days:
        return 0.0
    return max(0.0, 1.0 - (delta / float(tol_days)))


def score_vendor(inv_vendor, order_vendor):
    if not inv_vendor or not order_vendor:
        return None
    left = normalize_vendor(inv_vendor)
    right = normalize_vendor(order_vendor)
    if not left or not right:
        return None
    return SequenceMatcher(None, left, right).ratio()


def max_vendor_similarity(order_vendor, invoice_vendor_norms):
    left = normalize_vendor(order_vendor)
    if not left:
        return 0.0
    best = 0.0
    for right in invoice_vendor_norms:
        if not right:
            continue
        score = SequenceMatcher(None, left, right).ratio()
        if score > best:
            best = score
            if best >= 0.99:
                break
    return best


def match_orders_invoices(orders_df, invoices, config):
    rules = config.get("match_rules", {})
    weights = rules.get("weights", {})
    weight_amount = float(weights.get("amount", 0.6))
    weight_date = float(weights.get("date", 0.2))
    weight_vendor = float(weights.get("vendor", 0.2))
    min_score = float(rules.get("min_score", 0.7))

    tol_abs = rules.get("amount_tolerance", 0.0)
    tol_ratio = rules.get("amount_tolerance_ratio", 0.0)
    tol_days = int(rules.get("date_days_tolerance", 0))
    allow_dupe_orders = bool(rules.get("allow_duplicate_orders", False))

    pairs = []
    for inv_idx, invoice in enumerate(invoices):
        for order_idx, order in orders_df.iterrows():
            amount_score = score_amount(
                invoice["amount"], order["_amount"], tol_abs, tol_ratio
            )
            date_score = score_date(invoice["date"], order["_date"], tol_days)
            vendor_score = score_vendor(invoice["vendor"], order["_vendor"])

            score_sum = 0.0
            weight_sum = 0.0
            if amount_score is not None:
                score_sum += weight_amount * amount_score
                weight_sum += weight_amount
            if date_score is not None:
                score_sum += weight_date * date_score
                weight_sum += weight_date
            if vendor_score is not None:
                score_sum += weight_vendor * vendor_score
                weight_sum += weight_vendor

            total_score = (score_sum / weight_sum) if weight_sum else 0.0
            if total_score >= min_score:
                pairs.append(
                    {
                        "score": total_score,
                        "invoice_index": inv_idx,
                        "order_index": order_idx,
                        "score_amount": amount_score,
                        "score_date": date_score,
                        "score_vendor": vendor_score,
                    }
                )

    pairs.sort(key=lambda x: x["score"], reverse=True)
    matched_invoices = set()
    matched_orders = set()
    matches = []

    for pair in pairs:
        inv_idx = pair["invoice_index"]
        order_idx = pair["order_index"]
        if inv_idx in matched_invoices:
            continue
        if not allow_dupe_orders and order_idx in matched_orders:
            continue
        matched_invoices.add(inv_idx)
        matched_orders.add(order_idx)
        invoice = invoices[inv_idx]
        order = orders_df.loc[order_idx]
        invoice["match_order_row"] = int(order_idx)
        invoice["match_order_id"] = str(order["_order_id"])
        invoice["match_score"] = pair["score"]
        orders_df.at[order_idx, "_match_invoice_file"] = invoice["invoice_file"]
        orders_df.at[order_idx, "_match_invoice_id"] = invoice["invoice_id"]
        orders_df.at[order_idx, "_match_score"] = pair["score"]
        matches.append(
            {
                "invoice_file": invoice["invoice_file"],
                "invoice_id": invoice["invoice_id"],
                "invoice_amount": invoice["amount"],
                "invoice_date": invoice["date"],
                "invoice_vendor": invoice["vendor"],
                "order_row": int(order_idx),
                "order_id": order["_order_id"],
                "order_amount": order["_amount"],
                "order_date": order["_date"],
                "order_vendor": order["_vendor"],
                "score": pair["score"],
                "score_amount": pair["score_amount"],
                "score_date": pair["score_date"],
                "score_vendor": pair["score_vendor"],
            }
        )

    return matches


def compute_order_match_statuses(orders_df, invoices, matches, config):
    rules = config.get("match_rules", {})
    vendor_threshold = float(rules.get("vendor_similarity_threshold", 0.6))

    invoice_amount_counts = {}
    invoice_vendor_norms = []
    for invoice in invoices:
        key = amount_key(invoice.get("amount"))
        if key is not None:
            invoice_amount_counts[key] = invoice_amount_counts.get(key, 0) + 1
        invoice_vendor_norms.append(normalize_vendor(invoice.get("vendor")))

    match_by_order = {int(item["order_row"]): item for item in matches}
    status_map = {}

    for order_idx, order in orders_df.iterrows():
        order_amount_key = amount_key(order["_amount"])
        amount_match_count = (
            invoice_amount_counts.get(order_amount_key, 0)
            if order_amount_key is not None
            else 0
        )
        matched = match_by_order.get(int(order_idx))
        matched_amount_ok = False
        if matched:
            matched_amount_ok = amounts_equal(
                order["_amount"], matched.get("invoice_amount")
            )

        vendor_match = False
        if order["_vendor"]:
            vendor_match = (
                max_vendor_similarity(order["_vendor"], invoice_vendor_norms)
                >= vendor_threshold
            )

        if matched and matched_amount_ok and amount_match_count == 1:
            status = "完美匹配"
            reason = "金额一致且唯一匹配"
            color = "green"
        elif amount_match_count > 1:
            status = "疑似匹配"
            reason = "金额对应多张发票"
            color = "orange"
        elif vendor_match and amount_match_count == 0:
            status = "疑似匹配"
            reason = "开票单位相近但金额不一致"
            color = "orange"
        elif matched:
            status = "疑似匹配"
            reason = "匹配未满足完美条件"
            color = "orange"
        else:
            status = "无匹配"
            reason = "未找到匹配发票"
            color = "red"

        status_map[order_idx] = {
            "status": status,
            "reason": reason,
            "color": color,
            "amount_match_count": amount_match_count,
        }

    return status_map


def write_output(out_path, orders_df, invoices, matches):
    try:
        import pandas as pd
    except ImportError:
        die("Missing dependency: pandas. Install with `pip install pandas openpyxl`.")

    invoices_df = pd.DataFrame(invoices)
    matches_df = pd.DataFrame(matches)

    unmatched_invoices_df = invoices_df[
        invoices_df["match_order_row"].isna()
    ].copy()
    unmatched_orders_df = orders_df[
        orders_df["_match_invoice_file"].isna()
    ].copy()

    out_path = Path(out_path)
    if out_path.suffix.lower() == ".xlsx":
        try:
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                orders_df.to_excel(writer, sheet_name="orders", index=False)
                invoices_df.to_excel(writer, sheet_name="invoices", index=False)
                matches_df.to_excel(writer, sheet_name="matches", index=False)
                unmatched_invoices_df.to_excel(
                    writer, sheet_name="unmatched_invoices", index=False
                )
                unmatched_orders_df.to_excel(
                    writer, sheet_name="unmatched_orders", index=False
                )
            return
        except Exception as exc:
            print(f"Failed to write Excel output: {exc}", file=sys.stderr)

    stem = out_path.stem
    parent = out_path.parent
    orders_df.to_csv(parent / f"{stem}_orders.csv", index=False)
    invoices_df.to_csv(parent / f"{stem}_invoices.csv", index=False)
    matches_df.to_csv(parent / f"{stem}_matches.csv", index=False)
    unmatched_invoices_df.to_csv(parent / f"{stem}_unmatched_invoices.csv", index=False)
    unmatched_orders_df.to_csv(parent / f"{stem}_unmatched_orders.csv", index=False)


def run_matching(orders_path, pdf_dir, config_path, out_path, recursive=False):
    config = load_config(config_path)
    backend = resolve_pdf_backend()
    if backend is None:
        die("No PDF backend found. Install `pdfplumber` or `pypdf` to read PDFs.")

    orders_df = load_orders(orders_path, config)
    pdfs = find_pdfs(pdf_dir, recursive)
    if not pdfs:
        die("No PDF files found in the specified directory.")

    invoices = [extract_invoice(path, config, backend) for path in pdfs]
    matches = match_orders_invoices(orders_df, invoices, config)
    status_map = compute_order_match_statuses(orders_df, invoices, matches, config)
    orders_df["_match_status"] = orders_df.index.map(
        lambda idx: status_map.get(idx, {}).get("status", "")
    )
    orders_df["_match_reason"] = orders_df.index.map(
        lambda idx: status_map.get(idx, {}).get("reason", "")
    )
    write_output(out_path, orders_df, invoices, matches)

    return {
        "orders": len(orders_df),
        "invoices": len(invoices),
        "matches": len(matches),
        "output": str(out_path),
    }


def main():
    parser = build_parser()
    args = parser.parse_args()

    if args.init_config:
        write_json(args.config, DEFAULT_CONFIG, force=args.force)
        print(f"Wrote config to {args.config}")
        return

    if not args.orders or not args.pdf_dir:
        die("Both --orders and --pdf-dir are required unless --init-config is used.")

    summary = run_matching(args.orders, args.pdf_dir, args.config, args.out, args.recursive)
    print(
        f"Processed {summary['orders']} orders, {summary['invoices']} invoices."
    )
    print(f"Matches found: {summary['matches']}")
    print(f"Output written to: {summary['output']}")


if __name__ == "__main__":
    main()
