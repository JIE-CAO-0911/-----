#!/usr/bin/env python3
import base64
import hashlib
import html
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

import invoice_matcher as matcher
import screenshot_manager as shots

try:
    from PIL import Image

    PIL_AVAILABLE = True
except Exception:
    Image = None
    PIL_AVAILABLE = False

try:
    import streamlit.runtime.runtime as st_runtime

    RUNTIME_AVAILABLE = True
except Exception:
    RUNTIME_AVAILABLE = False

try:
    from st_aggrid import (
        AgGrid,
        DataReturnMode,
        GridOptionsBuilder,
        GridUpdateMode,
    )

    try:
        from st_aggrid import JsCode

        JS_CODE_AVAILABLE = True
    except Exception:
        JsCode = None
        JS_CODE_AVAILABLE = False

    AGGRID_AVAILABLE = True
except Exception:
    AGGRID_AVAILABLE = False
    JS_CODE_AVAILABLE = False

try:
    import tkinter as tk
    from tkinter import filedialog

    TK_AVAILABLE = True
except Exception:
    TK_AVAILABLE = False


matcher.RAISE_ON_ERROR = True

APP_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG_PATH = APP_DIR / "invoice_matcher_config.json"
DEFAULT_OUTPUT_PATH = APP_DIR / "match_result.xlsx"
DEFAULT_ORDER_FILES = ["订单数据.xlsx", "订单数据 (1).xlsx"]
DATA_ROOT_DIR = APP_DIR / "data_store"
CACHE_ROOT_DIR = DATA_ROOT_DIR / "cache"
HISTORY_ROOT_DIR = DATA_ROOT_DIR / "history"
PREPROCESS_CACHE_DIR = CACHE_ROOT_DIR / "preprocess"
SCREENSHOT_CACHE_DIR = CACHE_ROOT_DIR / "screenshot"
AUTO_SHUTDOWN_IDLE_SECONDS = 1
AUTO_SHUTDOWN_CHECK_SECONDS = 0.5
AUTO_SHUTDOWN_STARTED = False
PREPROCESS_CACHE_PATH = PREPROCESS_CACHE_DIR / "state.json"
LEGACY_PREPROCESS_CACHE_PATH = APP_DIR / "preprocess_cache.json"
SCREENSHOT_CACHE_PATH = SCREENSHOT_CACHE_DIR / "state.json"
PREPROCESS_HISTORY_LIMIT = 20
GRID_CUSTOM_CSS = {
    ".ag-cell.product-link-cell": {
        "background-color": "#eef2ff",
        "color": "#1d4ed8",
        "border-radius": "4px",
        "text-align": "center",
        "cursor": "pointer",
    },
    ".ag-cell.product-link-cell:hover": {
        "background-color": "#e0e7ff",
    },
    ".ag-cell.product-link-empty": {
        "color": "#888888",
        "text-align": "center",
    },
    ".ag-row.ag-row-selected": {
        "outline": "2px solid #f59e0b",
        "outline-offset": "-2px",
    },
    ".ag-cell.screenshot-status-none": {
        "background-color": "#fee2e2",
        "color": "#991b1b",
        "border-radius": "4px",
        "text-align": "center",
        "cursor": "pointer",
        "font-weight": "600",
    },
    ".ag-cell.screenshot-status-partial": {
        "background-color": "#ffedd5",
        "color": "#9a3412",
        "border-radius": "4px",
        "text-align": "center",
        "cursor": "pointer",
        "font-weight": "600",
    },
    ".ag-cell.screenshot-status-full": {
        "background-color": "#dcfce7",
        "color": "#166534",
        "border-radius": "4px",
        "text-align": "center",
        "cursor": "pointer",
        "font-weight": "600",
    },
}
EDITOR_COLUMNS = [
    "order_id",
    "vendor",
    "description",
    "product_link",
    "amount",
    "order_status",
    "order_date",
    "_row_id",
]
REIMBURSED_ORDER_ID_CANDIDATES = [
    "order_id",
    "订单号",
    "订单编号",
    "订单ID",
    "订单id",
]
REIMBURSED_DETAIL_SHEET_CANDIDATES = [
    "details",
    "detail",
    "报销明细",
    "明细",
]
REIMBURSED_AMOUNT_CANDIDATES = [
    "amount",
    "实付",
    "实付金额",
    "金额",
    "支付金额",
]
REIMBURSED_DATE_CANDIDATES = [
    "order_date",
    "时间",
    "订单时间",
    "下单时间",
    "订单提交时间",
    "日期",
]
REIMBURSED_VENDOR_CANDIDATES = [
    "vendor",
    "店铺",
    "店铺名称",
    "商家",
    "卖家",
]
MAIN_TAB_ITEMS = [
    {
        "key": "预处理",
        "label": "订单预处理",
        "desc": "导入订单、去重筛选、清理已报销记录",
    },
    {
        "key": "匹配",
        "label": "发票匹配核对",
        "desc": "自动匹配订单与发票，集中复核疑似结果",
    },
    {
        "key": "截图",
        "label": "截图资料归档",
        "desc": "补齐凭证截图并导出报销压缩包/打印版",
    },
]
MAIN_TAB_LABELS = {item["key"]: item["label"] for item in MAIN_TAB_ITEMS}
MAIN_TAB_DESCS = {item["key"]: item["desc"] for item in MAIN_TAB_ITEMS}


def default_orders_path():
    for name in DEFAULT_ORDER_FILES:
        path = APP_DIR / name
        if path.exists():
            return str(path)
    return ""


def load_config(path):
    try:
        return matcher.load_config(path)
    except matcher.MatchError as exc:
        st.error(str(exc))
        return matcher.DEFAULT_CONFIG


def save_config(path, config):
    try:
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(config, fh, indent=2, ensure_ascii=False)
            fh.write("\n")
        st.success("已保存")
    except Exception as exc:
        st.error(f"保存失败：{exc}")


def inject_ui_styles():
    st.markdown(
        """
        <style>
        html, body, [class*="css"] {
            font-family: "MiSans", "PingFang SC", "Source Han Sans SC", "Microsoft YaHei", sans-serif;
        }
        .main .block-container {
            max-width: 1320px;
            padding-top: 1.2rem;
            padding-bottom: 2rem;
        }
        .assistant-hero {
            background: linear-gradient(120deg, #ecfeff 0%, #eef2ff 52%, #fff7ed 100%);
            border: 1px solid #d8e8ea;
            border-radius: 20px;
            padding: 18px 22px 16px 22px;
            margin-bottom: 0.8rem;
        }
        .assistant-hero h1 {
            margin: 0;
            font-size: 2rem;
            line-height: 1.15;
            color: #0f2f3a;
            letter-spacing: 0.4px;
        }
        .assistant-hero p {
            margin: 8px 0 0 0;
            color: #35515d;
            font-size: 0.95rem;
        }
        div[data-testid="stSegmentedControl"] > div {
            gap: 0.4rem;
        }
        div[data-testid="stSegmentedControl"] button {
            border-radius: 999px !important;
            border: 1px solid #cddfe2 !important;
            font-weight: 650 !important;
            letter-spacing: 0.2px;
            padding: 0.35rem 0.9rem !important;
        }
        div[data-testid="stSegmentedControl"] button[aria-pressed="true"] {
            background: linear-gradient(135deg, #0f766e 0%, #115e59 100%) !important;
            color: #ffffff !important;
            border-color: #115e59 !important;
        }
        div[data-testid="stMetric"] {
            background: #f8fbfc;
            border: 1px solid #d8e8ea;
            border-radius: 14px;
            padding: 0.45rem 0.75rem;
        }
        div[data-testid="stMetricLabel"] p {
            font-weight: 650;
        }
        .date-span-card {
            background: #f8fbfc;
            border: 1px solid #d8e8ea;
            border-radius: 14px;
            padding: 0.46rem 0.75rem 0.5rem 0.75rem;
            min-height: 96px;
        }
        .date-span-label {
            color: #2d3f47;
            font-weight: 650;
            font-size: 0.88rem;
            line-height: 1.2;
            margin-bottom: 0.3rem;
        }
        .date-span-row {
            display: flex;
            align-items: baseline;
            gap: 0.42rem;
            margin: 0.1rem 0;
        }
        .date-span-sub {
            color: #5b6d75;
            font-weight: 650;
            font-size: 0.78rem;
            letter-spacing: 0.2px;
            min-width: 2rem;
        }
        .date-span-line {
            color: #12222c;
            font-weight: 650;
            font-size: 1.08rem;
            line-height: 1.32;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .import-divider {
            width: 1px;
            min-height: 198px;
            height: 100%;
            background: linear-gradient(180deg, #dce7ea 0%, #cfdde2 100%);
            margin: 0 auto;
            border-radius: 999px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_page_hero():
    st.markdown(
        """
        <div class="assistant-hero">
          <h1>报销助手</h1>
          <p>整合订单预处理、发票匹配核对、截图资料归档，减少重复操作并提升报销效率。</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_main_navigation():
    tab_keys = [item["key"] for item in MAIN_TAB_ITEMS]
    pending_tab = st.session_state.get("pending_tab")
    if pending_tab in tab_keys:
        st.session_state.active_tab = pending_tab
        st.session_state.pending_tab = None

    current_tab = st.session_state.get("active_tab")
    if current_tab not in tab_keys:
        current_tab = tab_keys[0]
        st.session_state.active_tab = current_tab

    active_tab = st.segmented_control(
        "主要功能",
        options=tab_keys,
        format_func=lambda key: MAIN_TAB_LABELS.get(key, key),
        selection_mode="single",
        key="active_tab",
        label_visibility="collapsed",
    )
    if active_tab is None:
        active_tab = st.session_state.get("active_tab", current_tab)
    st.caption(MAIN_TAB_DESCS.get(active_tab, ""))
    return active_tab


def sync_sidebar_config_state(config_path):
    state_key = "sidebar_config_sync_path"
    target = str(config_path or "")
    if st.session_state.get(state_key) == target:
        return

    config = load_config(config_path)
    order_cols = config.get("order_columns", {})
    rules = config.get("match_rules", {})

    st.session_state["sidebar_order_id_col"] = order_cols.get("order_id", "")
    st.session_state["sidebar_amount_col"] = order_cols.get("amount", "")
    st.session_state["sidebar_date_col"] = order_cols.get("date", "")
    st.session_state["sidebar_vendor_col"] = order_cols.get("vendor", "")
    st.session_state["sidebar_desc_col"] = order_cols.get("description", "")
    st.session_state["sidebar_status_col"] = order_cols.get("status", "")
    st.session_state["sidebar_link_col"] = order_cols.get("product_link", "")
    st.session_state["sidebar_merge_items"] = bool(
        rules.get("merge_multi_item_orders", True)
    )
    st.session_state[state_key] = target


def render_sidebar_settings(config_path):
    sync_sidebar_config_state(config_path)

    st.subheader("列映射")
    col1, col2 = st.columns(2)
    with col1:
        st.text_input("订单号", key="sidebar_order_id_col")
        st.text_input("实付", key="sidebar_amount_col")
        st.text_input("时间", key="sidebar_date_col")
        st.text_input("店铺", key="sidebar_vendor_col")
    with col2:
        st.text_input("商品", key="sidebar_desc_col")
        st.text_input("状态", key="sidebar_status_col")
        st.text_input("商品链接", key="sidebar_link_col")
        st.checkbox("合并明细", key="sidebar_merge_items")

    if st.button("保存设置", key="sidebar_save_settings"):
        config = load_config(config_path)
        config["order_columns"] = {
            "order_id": str(st.session_state.get("sidebar_order_id_col", "")).strip(),
            "amount": str(st.session_state.get("sidebar_amount_col", "")).strip(),
            "date": str(st.session_state.get("sidebar_date_col", "")).strip(),
            "vendor": str(st.session_state.get("sidebar_vendor_col", "")).strip(),
            "description": str(st.session_state.get("sidebar_desc_col", "")).strip(),
            "status": str(st.session_state.get("sidebar_status_col", "")).strip(),
            "product_link": str(st.session_state.get("sidebar_link_col", "")).strip(),
        }
        config.setdefault("match_rules", {})
        config["match_rules"]["merge_multi_item_orders"] = bool(
            st.session_state.get("sidebar_merge_items", True)
        )
        config["match_rules"].pop("allow_duplicate_orders", None)
        config["match_rules"].pop("vendor_similarity_threshold", None)
        save_config(config_path, config)


def ensure_data_directories():
    for folder in (
        DATA_ROOT_DIR,
        CACHE_ROOT_DIR,
        HISTORY_ROOT_DIR,
        PREPROCESS_CACHE_DIR,
        SCREENSHOT_CACHE_DIR,
    ):
        folder.mkdir(parents=True, exist_ok=True)


def to_bool(value, default=False):
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    if isinstance(value, str):
        text = value.strip().lower()
        if text in {"1", "true", "yes", "y", "on"}:
            return True
        if text in {"0", "false", "no", "n", "off", ""}:
            return False
    return default


def ensure_session_state():
    ensure_data_directories()
    if "order_df" not in st.session_state:
        st.session_state.order_df = pd.DataFrame(columns=EDITOR_COLUMNS)
    shots.init_session_state()
    if "order_row_id_seq" not in st.session_state:
        st.session_state.order_row_id_seq = 1
    if "order_grid_version" not in st.session_state:
        st.session_state.order_grid_version = 0
    if "match_result_df" not in st.session_state:
        st.session_state.match_result_df = None
    if "match_detail" not in st.session_state:
        st.session_state.match_detail = {}
    if "summary" not in st.session_state:
        st.session_state.summary = None
    if "orders_paths_text" not in st.session_state:
        st.session_state.orders_paths_text = default_orders_path()
    if "pdf_dir" not in st.session_state:
        st.session_state.pdf_dir = r"D:\发票\报销中"
    if "recursive_scan" not in st.session_state:
        st.session_state.recursive_scan = False
    if "config_path" not in st.session_state:
        st.session_state.config_path = str(DEFAULT_CONFIG_PATH)
    if "output_path" not in st.session_state:
        st.session_state.output_path = str(DEFAULT_OUTPUT_PATH)
    if "color_perfect" not in st.session_state:
        st.session_state.color_perfect = "#d6f5d6"
    if "color_suspect" not in st.session_state:
        st.session_state.color_suspect = "#ffe9c6"
    if "color_nomatch" not in st.session_state:
        st.session_state.color_nomatch = "#ffd6d6"
    if "preprocess_notice" not in st.session_state:
        st.session_state.preprocess_notice = ""
    if "reimbursed_notice" not in st.session_state:
        st.session_state.reimbursed_notice = ""
    if "reimbursed_order_ids" not in st.session_state:
        st.session_state.reimbursed_order_ids = []
    if "reimbursed_fallback_keys" not in st.session_state:
        st.session_state.reimbursed_fallback_keys = []
    if "reimbursed_source_name" not in st.session_state:
        st.session_state.reimbursed_source_name = ""
    if "reimbursed_source_rows" not in st.session_state:
        st.session_state.reimbursed_source_rows = 0
    if "preprocess_auto_save" not in st.session_state:
        st.session_state.preprocess_auto_save = True
    st.session_state.preprocess_auto_save = to_bool(
        st.session_state.get("preprocess_auto_save"), default=True
    )
    if "preprocess_undo_stack" not in st.session_state:
        st.session_state.preprocess_undo_stack = []
    if "preprocess_redo_stack" not in st.session_state:
        st.session_state.preprocess_redo_stack = []
    if "preprocess_cache_loaded" not in st.session_state:
        st.session_state.preprocess_cache_loaded = False
    if "screenshot_orders_df" not in st.session_state:
        st.session_state.screenshot_orders_df = None
    if "screenshot_notice" not in st.session_state:
        st.session_state.screenshot_notice = ""
    if "screenshot_auto_save" not in st.session_state:
        st.session_state.screenshot_auto_save = True
    st.session_state.screenshot_auto_save = to_bool(
        st.session_state.get("screenshot_auto_save"), default=True
    )
    if "settings_migrated_v1" not in st.session_state:
        st.session_state.screenshot_auto_save = True
        st.session_state.settings_migrated_v1 = True
    if "settings_migrated_v2" not in st.session_state:
        st.session_state.screenshot_auto_save = True
        st.session_state.settings_migrated_v2 = True
    if "screenshot_cache_loaded" not in st.session_state:
        st.session_state.screenshot_cache_loaded = False
    if "screenshot_last_saved_signature" not in st.session_state:
        st.session_state.screenshot_last_saved_signature = ""
    if "active_tab" not in st.session_state:
        st.session_state.active_tab = "预处理"
    if "screenshot_prefill_df" not in st.session_state:
        st.session_state.screenshot_prefill_df = None
    if "pending_tab" not in st.session_state:
        st.session_state.pending_tab = None
    if "pending_match_auto_start" not in st.session_state:
        st.session_state.pending_match_auto_start = False
    if "pending_clear_add_order_form" not in st.session_state:
        st.session_state.pending_clear_add_order_form = False
    if not st.session_state.preprocess_cache_loaded:
        snapshot = load_preprocess_cache()
        if snapshot is not None:
            restore_preprocess_snapshot(
                snapshot,
                notice=f"已恢复暂存数据（{len(snapshot.get('rows', []))} 条）",
                reset_grid=True,
            )
            st.session_state.preprocess_undo_stack = []
            st.session_state.preprocess_redo_stack = []
        st.session_state.preprocess_cache_loaded = True
    if not st.session_state.screenshot_cache_loaded:
        snapshot = load_screenshot_cache()
        if snapshot is not None:
            restore_screenshot_snapshot(
                snapshot,
                notice=f"已恢复截图暂存（{len(snapshot.get('rows', []))} 条）",
            )
            st.session_state.screenshot_last_saved_signature = screenshot_snapshot_signature(
                snapshot
            )
        st.session_state.screenshot_cache_loaded = True


def orders_df_for_editor(orders_df):
    return pd.DataFrame(
        {
            "order_id": orders_df["_order_id"].apply(matcher.format_order_id),
            "vendor": orders_df["_vendor"],
            "description": orders_df["_description"],
            "product_link": orders_df.get("_product_link", ""),
            "amount": orders_df["_amount"],
            "order_status": orders_df["_status"],
            "order_date": orders_df["_date"],
        }
    )


def editor_df_to_rows(df):
    rows = []
    for _, row in df.iterrows():
        rows.append(
            {
                "order_id": matcher.format_order_id(row.get("order_id")),
                "vendor": str(row.get("vendor") or "").strip(),
                "description": str(row.get("description") or "").strip(),
                "product_link": str(row.get("product_link") or "").strip(),
                "amount": row.get("amount"),
                "order_status": str(row.get("order_status") or "").strip(),
                "order_date": row.get("order_date"),
            }
        )
    return rows


def style_match_rows(row, color_map):
    color = color_map.get(row.get("match_status", ""))
    if not color:
        return ["" for _ in row]
    return [f"background-color: {color}; color: #000" for _ in row]


def parse_orders_paths(text):
    if not text:
        return []
    flat = text.replace(";", "\n")
    paths = [line.strip() for line in flat.splitlines() if line.strip()]
    return paths


def find_column_by_candidates(columns, candidates):
    stripped = {str(col).strip(): col for col in columns}
    lowered = {str(col).strip().lower(): col for col in columns}
    for cand in candidates:
        if cand in stripped:
            return stripped[cand]
        cand_lower = str(cand).strip().lower()
        if cand_lower in lowered:
            return lowered[cand_lower]
    return None


def normalize_order_id_for_match(value):
    return matcher.format_order_id(value).strip()


def normalize_amount_for_match(value):
    amount = matcher.parse_amount(value)
    if amount is None:
        return ""
    return f"{round(float(amount), 2):.2f}"


def normalize_date_for_match(value):
    parsed = matcher.parse_date(value)
    if parsed:
        return parsed.isoformat()
    text = str(value or "").strip()
    return "" if text.lower() in {"", "nan", "none"} else text


def normalize_vendor_for_match(value):
    return matcher.normalize_vendor(value)


def build_amount_date_vendor_key(amount, order_date, vendor):
    amount_key = normalize_amount_for_match(amount)
    date_key = normalize_date_for_match(order_date)
    vendor_key = normalize_vendor_for_match(vendor)
    if not amount_key or not date_key or not vendor_key:
        return ""
    return "|".join([amount_key, date_key, vendor_key])


def deduplicate_editor_orders(df):
    if df is None or df.empty:
        return df, 0

    key_df = pd.DataFrame(index=df.index)
    key_df["_k_order_id"] = df["order_id"].apply(normalize_order_id_for_match)
    key_df["_k_vendor"] = df["vendor"].apply(
        lambda v: matcher.normalize_text(v).strip().lower()
    )
    key_df["_k_description"] = df["description"].apply(
        lambda v: matcher.normalize_text(v).strip().lower()
    )
    key_df["_k_link"] = df["product_link"].apply(
        lambda v: matcher.normalize_text(v).strip().lower()
    )
    key_df["_k_amount"] = df["amount"].apply(normalize_amount_for_match)
    key_df["_k_status"] = df["order_status"].apply(
        lambda v: matcher.normalize_text(v).strip().lower()
    )
    key_df["_k_date"] = df["order_date"].apply(normalize_date_for_match)

    duplicated = key_df.duplicated(keep="first")
    removed = int(duplicated.sum())
    if removed == 0:
        return df, 0
    deduped = df.loc[~duplicated].reset_index(drop=True)
    return deduped, removed


def compute_reimbursed_mask(df, reimbursed_order_ids, reimbursed_fallback_keys):
    if df is None or df.empty:
        return pd.Series([], dtype=bool)
    if not reimbursed_order_ids and not reimbursed_fallback_keys:
        return pd.Series(False, index=df.index)

    order_ids = df["order_id"].apply(normalize_order_id_for_match)
    id_mask = order_ids.apply(lambda value: bool(value) and value in reimbursed_order_ids)

    if not reimbursed_fallback_keys:
        return id_mask

    fallback_keys = df.apply(
        lambda row: build_amount_date_vendor_key(
            row.get("amount"),
            row.get("order_date"),
            row.get("vendor"),
        ),
        axis=1,
    )
    fallback_mask = order_ids.apply(lambda value: not bool(value)) & fallback_keys.apply(
        lambda key: bool(key) and key in reimbursed_fallback_keys
    )
    return id_mask | fallback_mask


def add_reimbursed_flag_column(df, reimbursed_order_ids, reimbursed_fallback_keys):
    display_df = df.copy()
    if "_reimbursed" in display_df.columns:
        display_df = display_df.drop(columns=["_reimbursed"])
    mask = compute_reimbursed_mask(
        display_df, reimbursed_order_ids, reimbursed_fallback_keys
    )
    # Keep helper column at the end so AgGrid checkbox column stays visible.
    display_df["_reimbursed"] = mask.values
    return display_df


def load_reimbursed_excel(data):
    workbook = pd.read_excel(io.BytesIO(data), sheet_name=None)
    if not workbook:
        raise ValueError("Excel 文件为空")

    for sheet in REIMBURSED_DETAIL_SHEET_CANDIDATES:
        for name, frame in workbook.items():
            if str(name).strip().lower() == sheet.lower():
                return frame

    for frame in workbook.values():
        if not frame.empty:
            return frame
    return next(iter(workbook.values()))


def load_reimbursed_orders_from_zip(data):
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        files = [name for name in zf.namelist() if not name.endswith("/")]
        if not files:
            raise ValueError("ZIP 文件为空")

        def score(name):
            base = Path(name).name.lower()
            priority = 0
            if "detail" in base or "明细" in base:
                priority -= 10
            if base.endswith(".xlsx"):
                priority -= 5
            elif base.endswith(".xls"):
                priority -= 4
            elif base.endswith(".csv"):
                priority -= 3
            return (priority, len(base))

        for member in sorted(files, key=score):
            lower = member.lower()
            if lower.endswith((".xlsx", ".xls")):
                with zf.open(member) as fh:
                    data_bytes = fh.read()
                try:
                    return load_reimbursed_excel(data_bytes)
                except Exception:
                    continue
            if lower.endswith(".csv"):
                with zf.open(member) as fh:
                    data_bytes = fh.read()
                try:
                    return pd.read_csv(io.BytesIO(data_bytes))
                except Exception:
                    continue

    raise ValueError("ZIP 中未找到可用的报销明细文件（xlsx/xls/csv）")


def load_reimbursed_orders_file(uploaded_file):
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()
    if name.endswith(".zip"):
        return load_reimbursed_orders_from_zip(data)
    if name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(data))
    if name.endswith((".xlsx", ".xls")):
        return load_reimbursed_excel(data)
    raise ValueError("不支持的文件类型")


def collect_reimbursed_matchers(df):
    if df is None or df.empty:
        return {"order_ids": [], "fallback_keys": []}

    order_col = find_column_by_candidates(df.columns, REIMBURSED_ORDER_ID_CANDIDATES)
    amount_col = find_column_by_candidates(df.columns, REIMBURSED_AMOUNT_CANDIDATES)
    date_col = find_column_by_candidates(df.columns, REIMBURSED_DATE_CANDIDATES)
    vendor_col = find_column_by_candidates(df.columns, REIMBURSED_VENDOR_CANDIDATES)

    order_ids = set()
    fallback_keys = set()

    for _, row in df.iterrows():
        order_id = normalize_order_id_for_match(row.get(order_col)) if order_col else ""
        if order_id:
            order_ids.add(order_id)
            continue
        key = build_amount_date_vendor_key(
            row.get(amount_col) if amount_col else None,
            row.get(date_col) if date_col else None,
            row.get(vendor_col) if vendor_col else None,
        )
        if key:
            fallback_keys.add(key)

    if not order_ids and not fallback_keys:
        raise ValueError("未读取到有效订单号，且无法生成金额+日期+店铺兜底键")

    return {
        "order_ids": sorted(order_ids),
        "fallback_keys": sorted(fallback_keys),
    }


def format_editor_date(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    if isinstance(value, datetime):
        return value.date().strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    parsed = matcher.parse_date(value)
    if parsed:
        return parsed.strftime("%Y-%m-%d")
    text = str(value).strip()
    return text if text.lower() not in {"nan", "none"} else ""


def ensure_editor_df(df):
    df = df.copy()
    defaults = {
        "order_id": "",
        "vendor": "",
        "description": "",
        "product_link": "",
        "amount": None,
        "order_status": "",
        "order_date": "",
        "_row_id": None,
    }
    for col, default in defaults.items():
        if col not in df.columns:
            df[col] = default
    df = df[EDITOR_COLUMNS]

    df["order_date"] = df["order_date"].apply(format_editor_date)

    df["_row_id"] = pd.to_numeric(df["_row_id"], errors="coerce")
    missing = df["_row_id"].isna()
    if missing.any():
        start = st.session_state.order_row_id_seq
        count = int(missing.sum())
        df.loc[missing, "_row_id"] = range(start, start + count)
        st.session_state.order_row_id_seq = start + count

    if not df["_row_id"].isna().all():
        max_id = int(df["_row_id"].max())
        st.session_state.order_row_id_seq = max(
            st.session_state.order_row_id_seq, max_id + 1
        )
        df["_row_id"] = df["_row_id"].astype(int)

    return df


def set_order_df(df, notice=None, reset_grid=False, reset_row_ids=False):
    if reset_row_ids:
        st.session_state.order_row_id_seq = 1
    df = ensure_editor_df(df)
    st.session_state.order_df = df.reset_index(drop=True)
    st.session_state.match_result_df = None
    st.session_state.match_detail = {}
    if reset_grid:
        st.session_state.order_grid_version += 1
    if notice:
        st.session_state.preprocess_notice = notice


def serialize_preprocess_value(value):
    if isinstance(value, (datetime, date)):
        return format_editor_date(value)
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if isinstance(value, Path):
        return str(value)
    return value


def snapshot_from_order_df(df=None, row_id_seq=None):
    if df is None:
        df = st.session_state.get("order_df")
    if not isinstance(df, pd.DataFrame):
        df = pd.DataFrame(columns=EDITOR_COLUMNS)
    working = df.copy()
    for col in EDITOR_COLUMNS:
        if col not in working.columns:
            working[col] = None
    working = working[EDITOR_COLUMNS]
    working["order_date"] = working["order_date"].apply(format_editor_date)

    rows = []
    for record in working.to_dict("records"):
        rows.append({k: serialize_preprocess_value(v) for k, v in record.items()})

    if row_id_seq is None:
        row_id_seq = st.session_state.get("order_row_id_seq", 1)
    try:
        row_id_seq_value = int(row_id_seq)
    except Exception:
        row_id_seq_value = 1

    return {"rows": rows, "row_id_seq": row_id_seq_value}


def snapshots_equal(left, right):
    if not isinstance(left, dict) or not isinstance(right, dict):
        return False
    return (
        int(left.get("row_id_seq", 1)) == int(right.get("row_id_seq", 1))
        and left.get("rows", []) == right.get("rows", [])
    )


def trim_history(stack):
    if len(stack) > PREPROCESS_HISTORY_LIMIT:
        del stack[: len(stack) - PREPROCESS_HISTORY_LIMIT]


def save_preprocess_cache(snapshot=None):
    if snapshot is None:
        snapshot = snapshot_from_order_df()
    ensure_data_directories()
    payload = {
        "version": 1,
        "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "snapshot": snapshot,
    }
    with open(PREPROCESS_CACHE_PATH, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    return PREPROCESS_CACHE_PATH


def load_preprocess_cache():
    candidate_paths = [PREPROCESS_CACHE_PATH, LEGACY_PREPROCESS_CACHE_PATH]
    payload = None
    used_path = None
    for path in candidate_paths:
        if not path.exists():
            continue
        try:
            with open(path, "r", encoding="utf-8") as fh:
                payload = json.load(fh)
            used_path = path
            break
        except Exception:
            payload = None
    if payload is None:
        return None

    snapshot = payload.get("snapshot")
    if not isinstance(snapshot, dict):
        return None
    rows = snapshot.get("rows")
    if not isinstance(rows, list):
        return None
    row_id_seq = snapshot.get("row_id_seq", 1)
    try:
        row_id_seq = int(row_id_seq)
    except Exception:
        row_id_seq = 1
    snapshot = {"rows": rows, "row_id_seq": row_id_seq}
    if used_path == LEGACY_PREPROCESS_CACHE_PATH and not PREPROCESS_CACHE_PATH.exists():
        try:
            save_preprocess_cache(snapshot)
        except Exception:
            pass
    return snapshot


def restore_preprocess_snapshot(snapshot, notice=None, reset_grid=True):
    rows = snapshot.get("rows", []) if isinstance(snapshot, dict) else []
    row_id_seq = snapshot.get("row_id_seq", 1) if isinstance(snapshot, dict) else 1
    df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=EDITOR_COLUMNS)
    set_order_df(df, notice=notice, reset_grid=reset_grid, reset_row_ids=True)
    try:
        row_id_seq_value = int(row_id_seq)
    except Exception:
        row_id_seq_value = st.session_state.get("order_row_id_seq", 1)
    st.session_state.order_row_id_seq = max(
        st.session_state.get("order_row_id_seq", 1), row_id_seq_value
    )


def maybe_auto_save_preprocess(force=False):
    if not force and not st.session_state.get("preprocess_auto_save", False):
        return
    try:
        save_preprocess_cache()
    except Exception as exc:
        st.session_state.preprocess_notice = f"自动保存失败：{exc}"


def on_preprocess_auto_save_toggle():
    if st.session_state.get("preprocess_auto_save", False):
        maybe_auto_save_preprocess(force=True)


def apply_preprocess_change(
    new_df,
    notice=None,
    reset_grid=False,
    reset_row_ids=False,
    track_history=True,
    force_save=False,
):
    before = snapshot_from_order_df()
    set_order_df(new_df, notice=notice, reset_grid=reset_grid, reset_row_ids=reset_row_ids)
    after = snapshot_from_order_df()
    changed = not snapshots_equal(before, after)

    if track_history and changed:
        undo_stack = st.session_state.get("preprocess_undo_stack", [])
        undo_stack.append(before)
        trim_history(undo_stack)
        st.session_state.preprocess_undo_stack = undo_stack
        st.session_state.preprocess_redo_stack = []

    if changed:
        maybe_auto_save_preprocess(force=force_save)
    elif force_save:
        maybe_auto_save_preprocess(force=True)

    return changed


def undo_preprocess_change():
    undo_stack = st.session_state.get("preprocess_undo_stack", [])
    if not undo_stack:
        return False

    snapshot = undo_stack.pop()
    st.session_state.preprocess_undo_stack = undo_stack

    redo_stack = st.session_state.get("preprocess_redo_stack", [])
    redo_stack.append(snapshot_from_order_df())
    trim_history(redo_stack)
    st.session_state.preprocess_redo_stack = redo_stack

    restore_preprocess_snapshot(snapshot, notice="已撤回", reset_grid=True)
    maybe_auto_save_preprocess()
    return True


def redo_preprocess_change():
    redo_stack = st.session_state.get("preprocess_redo_stack", [])
    if not redo_stack:
        return False

    snapshot = redo_stack.pop()
    st.session_state.preprocess_redo_stack = redo_stack

    undo_stack = st.session_state.get("preprocess_undo_stack", [])
    undo_stack.append(snapshot_from_order_df())
    trim_history(undo_stack)
    st.session_state.preprocess_undo_stack = undo_stack

    restore_preprocess_snapshot(snapshot, notice="已重做", reset_grid=True)
    maybe_auto_save_preprocess()
    return True


def ensure_screenshot_df_for_cache(df):
    if not isinstance(df, pd.DataFrame):
        df = pd.DataFrame()
    working = df.copy()
    defaults = {
        "order_id": "",
        "vendor": "",
        "description": "",
        "product_link": "",
        "amount": "",
        "order_status": "",
        "order_date": "",
    }
    for col, default in defaults.items():
        if col not in working.columns:
            working[col] = default
    if "_row_id" not in working.columns:
        working["_row_id"] = range(1, len(working) + 1)
    working["order_date"] = working["order_date"].apply(format_editor_date)
    return working


def serialize_screenshot_store(store):
    serialized = {}
    for order_key, entry in (store or {}).items():
        safe_key = str(order_key)
        order_images = []
        payment_images = []
        for img in entry.get(shots.CATEGORY_ORDER, []):
            if isinstance(img, (bytes, bytearray)):
                order_images.append(base64.b64encode(bytes(img)).decode("ascii"))
        for img in entry.get(shots.CATEGORY_PAYMENT, []):
            if isinstance(img, (bytes, bytearray)):
                payment_images.append(base64.b64encode(bytes(img)).decode("ascii"))
        serialized[safe_key] = {
            shots.CATEGORY_ORDER: order_images,
            shots.CATEGORY_PAYMENT: payment_images,
        }
    return serialized


def deserialize_screenshot_store(data):
    store = {}
    if not isinstance(data, dict):
        return store
    for order_key, entry in data.items():
        order_list = []
        payment_list = []
        for encoded in (entry or {}).get(shots.CATEGORY_ORDER, []):
            if not isinstance(encoded, str):
                continue
            try:
                order_list.append(base64.b64decode(encoded.encode("ascii")))
            except Exception:
                continue
        for encoded in (entry or {}).get(shots.CATEGORY_PAYMENT, []):
            if not isinstance(encoded, str):
                continue
            try:
                payment_list.append(base64.b64decode(encoded.encode("ascii")))
            except Exception:
                continue
        store[str(order_key)] = {
            shots.CATEGORY_ORDER: order_list,
            shots.CATEGORY_PAYMENT: payment_list,
        }
    return store


def snapshot_from_screenshot_state(df=None, store=None):
    if df is None:
        df = st.session_state.get("screenshot_orders_df")
    if store is None:
        store = st.session_state.get("screenshot_store", {})

    working_df = ensure_screenshot_df_for_cache(df)
    rows = []
    for record in working_df.to_dict("records"):
        rows.append({k: serialize_preprocess_value(v) for k, v in record.items()})

    return {
        "rows": rows,
        "store": serialize_screenshot_store(store),
    }


def screenshot_snapshot_signature(snapshot):
    payload = json.dumps(snapshot, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def save_screenshot_cache(snapshot=None):
    if snapshot is None:
        snapshot = snapshot_from_screenshot_state()
    ensure_data_directories()
    payload = {
        "version": 1,
        "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "snapshot": snapshot,
    }
    with open(SCREENSHOT_CACHE_PATH, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    return SCREENSHOT_CACHE_PATH


def load_screenshot_cache():
    if not SCREENSHOT_CACHE_PATH.exists():
        return None
    try:
        with open(SCREENSHOT_CACHE_PATH, "r", encoding="utf-8") as fh:
            payload = json.load(fh)
    except Exception:
        return None
    snapshot = payload.get("snapshot")
    if not isinstance(snapshot, dict):
        return None
    rows = snapshot.get("rows")
    store = snapshot.get("store")
    if not isinstance(rows, list) or not isinstance(store, dict):
        return None
    return {"rows": rows, "store": store}


def restore_screenshot_snapshot(snapshot, notice=None):
    rows = snapshot.get("rows", []) if isinstance(snapshot, dict) else []
    store_data = snapshot.get("store", {}) if isinstance(snapshot, dict) else {}
    df = pd.DataFrame(rows) if rows else pd.DataFrame()
    st.session_state.screenshot_orders_df = ensure_screenshot_df_for_cache(df)
    st.session_state.screenshot_store = deserialize_screenshot_store(store_data)
    if notice:
        st.session_state.screenshot_notice = notice


def maybe_auto_save_screenshot(force=False):
    if not force and not st.session_state.get("screenshot_auto_save", False):
        return
    try:
        snapshot = snapshot_from_screenshot_state()
        signature = screenshot_snapshot_signature(snapshot)
        if (
            not force
            and signature == st.session_state.get("screenshot_last_saved_signature", "")
        ):
            return
        save_screenshot_cache(snapshot)
        st.session_state.screenshot_last_saved_signature = signature
    except Exception as exc:
        st.session_state.screenshot_notice = f"截图自动保存失败：{exc}"


def on_screenshot_auto_save_toggle():
    if st.session_state.get("screenshot_auto_save", False):
        maybe_auto_save_screenshot(force=True)


def trigger_rerun():
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()


def build_export_df(df):
    export_df = ensure_editor_df(df).copy()
    if "_row_id" in export_df.columns:
        export_df = export_df.drop(columns=["_row_id"])
    return export_df


def build_export_df_with_config(df, config):
    export_df = build_export_df(df)
    order_cols = (config or {}).get("order_columns", {})

    def col_name(key, fallback):
        value = order_cols.get(key)
        return value.strip() if isinstance(value, str) and value.strip() else fallback

    rename_map = {
        "order_id": col_name("order_id", "order_id"),
        "amount": col_name("amount", "amount"),
        "order_date": col_name("date", "order_date"),
        "vendor": col_name("vendor", "vendor"),
        "description": col_name("description", "description"),
        "product_link": col_name("product_link", "product_link"),
        "order_status": col_name("status", "order_status"),
    }
    export_df = export_df.rename(columns=rename_map)
    ordered_cols = [
        rename_map["order_id"],
        rename_map["vendor"],
        rename_map["description"],
        rename_map["product_link"],
        rename_map["amount"],
        rename_map["order_status"],
        rename_map["order_date"],
    ]
    export_df = export_df[ordered_cols]
    return export_df


def merge_grid_updates(full_df, updated_df):
    full_df = ensure_editor_df(full_df)
    updated_df = ensure_editor_df(updated_df)
    full_indexed = full_df.set_index("_row_id")
    updated_indexed = updated_df.set_index("_row_id")
    full_indexed.update(updated_indexed)
    merged = full_indexed.reset_index()
    order = full_df["_row_id"].tolist()
    merged = merged.set_index("_row_id").loc[order].reset_index()
    return merged


def add_screenshot_status_column(df, order_id_field, row_id_field):
    display_df = df.copy()
    if "screenshot_status" in display_df.columns:
        display_df = display_df.drop(columns=["screenshot_status"])
    statuses = []
    for _, row in display_df.iterrows():
        order_key = shots.build_order_key(row.get(order_id_field), row.get(row_id_field))
        statuses.append(shots.get_status_label(order_key))
    display_df.insert(0, "screenshot_status", statuses)
    return display_df


def add_auto_unique_id_column(df, column_name="auto_unique_id", position=1):
    display_df = df.copy()
    if column_name in display_df.columns:
        display_df = display_df.drop(columns=[column_name])
    if "screenshot_status" in display_df.columns:
        insert_at = display_df.columns.get_loc("screenshot_status") + 1
    elif "order_id" in display_df.columns:
        insert_at = display_df.columns.get_loc("order_id") + 1
    else:
        insert_at = position
    insert_at = max(0, min(insert_at, len(display_df.columns)))
    display_df.insert(insert_at, column_name, range(1, len(display_df) + 1))
    return display_df


def add_focus_row_column(df, row_id_field, selected_row_id, force_focus=False):
    display_df = df.copy()
    if "_focus_row" in display_df.columns:
        display_df = display_df.drop(columns=["_focus_row"])
    if not force_focus:
        return display_df
    if row_id_field not in display_df.columns:
        display_df["_focus_row"] = False
        return display_df
    if selected_row_id is None or selected_row_id == "":
        display_df["_focus_row"] = False
        return display_df
    try:
        target = str(int(selected_row_id))
    except Exception:
        target = str(selected_row_id)
    display_df["_focus_row"] = display_df[row_id_field].astype(str) == target
    return display_df


def resolve_selected_row(selected_rows, display_df, row_id_field, state_key):
    lock_key = f"{state_key}_lock"
    saved = st.session_state.get(state_key)
    if st.session_state.get(lock_key) and saved is not None and not display_df.empty:
        match = display_df[display_df[row_id_field] == saved]
        if not match.empty:
            if (
                selected_rows
                and row_id_field in selected_rows[0]
                and selected_rows[0][row_id_field] == saved
            ):
                st.session_state[lock_key] = False
                return selected_rows[0]
            return match.iloc[0].to_dict()
    if selected_rows:
        selected = selected_rows[0]
        if row_id_field in selected:
            value = selected[row_id_field]
            try:
                value = int(value)
            except Exception:
                pass
            st.session_state[state_key] = value
        st.session_state[lock_key] = False
        return selected
    if saved is not None and not display_df.empty:
        match = display_df[display_df[row_id_field] == saved]
        if not match.empty:
            return match.iloc[0].to_dict()
    return None


def make_step_order_callback(row_ids, state_key, step):
    def _step_order():
        if not row_ids:
            return
        current = st.session_state.get(state_key)
        if current in row_ids:
            idx = row_ids.index(current)
            next_id = row_ids[(idx + step) % len(row_ids)]
        else:
            next_id = row_ids[0]
        st.session_state[state_key] = next_id
        st.session_state[f"{state_key}_lock"] = True
        trigger_rerun()

    return _step_order


def sum_amounts(values):
    total = 0.0
    for value in values:
        amount = None
        if isinstance(value, (int, float)):
            if isinstance(value, float):
                try:
                    if pd.isna(value):
                        continue
                except Exception:
                    pass
            amount = float(value)
        else:
            amount = matcher.parse_amount(value)
        if amount is None:
            continue
        total += float(amount)
    return total


def format_amount(value):
    try:
        return f"{float(value):,.2f}"
    except Exception:
        return str(value)


def compute_order_date_span(values):
    dates = []
    for value in values:
        parsed = matcher.parse_date(value)
        if parsed is not None:
            dates.append(parsed)
    if not dates:
        return "无有效日期", 0
    start_date = min(dates)
    end_date = max(dates)
    days = (end_date - start_date).days + 1
    if start_date == end_date:
        return start_date.strftime("%Y-%m-%d"), 1
    return f"{start_date:%Y-%m-%d} ~ {end_date:%Y-%m-%d}", days


def render_date_span_card(date_span_text):
    text = str(date_span_text or "").strip()
    if " ~ " in text:
        start_text, end_text = text.split(" ~ ", 1)
    elif "~" in text:
        start_text, end_text = text.split("~", 1)
        start_text = start_text.strip()
        end_text = end_text.strip()
    else:
        start_text, end_text = text, ""

    start_safe = html.escape(start_text or "-")
    end_safe = html.escape(end_text or "")
    if end_safe:
        content_html = (
            '<div class="date-span-row">'
            '<span class="date-span-sub">最早</span>'
            f'<span class="date-span-line">{start_safe}</span>'
            "</div>"
            '<div class="date-span-row">'
            '<span class="date-span-sub">最晚</span>'
            f'<span class="date-span-line">{end_safe}</span>'
            "</div>"
        )
    else:
        content_html = (
            '<div class="date-span-row">'
            f'<span class="date-span-line">{start_safe}</span>'
            "</div>"
        )
    st.markdown(
        f"""
        <div class="date-span-card">
          <div class="date-span-label">时间跨度</div>
          {content_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def open_pdf_file(path):
    try:
        if os.name == "nt":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
        return True
    except Exception as exc:
        st.warning(f"无法打开发票：{exc}")
        return False


def describe_invoice(invoice):
    name = Path(invoice.get("invoice_file", "")).name
    invoice_id = invoice.get("invoice_id")
    if invoice_id:
        return f"{name} ({invoice_id})"
    return name


def build_order_invoice_map(matches):
    return {int(item["order_row"]): item for item in matches}


def build_suspect_invoice_map(orders_df, invoices, config, matches):
    amount_map = {}
    inv_by_file = {}
    for invoice in invoices:
        inv_by_file[invoice.get("invoice_file")] = invoice
        key = matcher.amount_key(invoice.get("amount"))
        if key is None:
            continue
        amount_map.setdefault(key, []).append(invoice)

    match_map = build_order_invoice_map(matches)
    suspect_map = {}
    for order_idx, order in orders_df.iterrows():
        suspects = []
        amount_key = matcher.amount_key(order["_amount"])
        if amount_key is not None and amount_key in amount_map:
            suspects = amount_map[amount_key][:]

        if not suspects:
            matched = match_map.get(int(order_idx))
            if matched:
                invoice = inv_by_file.get(matched.get("invoice_file"))
                if invoice:
                    suspects = [invoice]

        suspect_map[int(order_idx)] = suspects

    return suspect_map


def start_auto_shutdown_monitor():
    global AUTO_SHUTDOWN_STARTED
    if AUTO_SHUTDOWN_STARTED:
        return
    if not RUNTIME_AVAILABLE:
        return
    if not st_runtime.Runtime.exists():
        return

    AUTO_SHUTDOWN_STARTED = True

    def monitor():
        idle_start = None
        while True:
            try:
                runtime = st_runtime.Runtime.instance()
                active = runtime._session_mgr.num_active_sessions()
            except Exception:
                active = None

            if active == 0:
                if idle_start is None:
                    idle_start = time.time()
                elif time.time() - idle_start >= AUTO_SHUTDOWN_IDLE_SECONDS:
                    try:
                        runtime.stop()
                    except Exception:
                        pass
                    os._exit(0)
            else:
                idle_start = None

            time.sleep(AUTO_SHUTDOWN_CHECK_SECONDS)

    thread = threading.Thread(target=monitor, daemon=True)
    thread.start()


def build_order_grid_options(
    df,
    editable,
    selected_row_id=None,
    row_id_field=None,
    force_focus=False,
    selection_mode="single",
    use_checkbox=False,
    header_checkbox=False,
):
    builder = GridOptionsBuilder.from_dataframe(df)
    builder.configure_default_column(
        editable=editable,
        resizable=True,
        sortable=True,
        filter=True,
    )
    if "_focus_row" in df.columns:
        builder.configure_column("_focus_row", hide=True, editable=False)
    if "_reimbursed" in df.columns:
        builder.configure_column("_reimbursed", hide=True, editable=False)
    if "auto_unique_id" in df.columns:
        builder.configure_column(
            "auto_unique_id", header_name="auto_unique_id", width=90, editable=False
        )
    if "screenshot_status" in df.columns:
        status_style = build_screenshot_status_style()
        builder.configure_column(
            "screenshot_status",
            header_name="截图状态",
            width=120,
            editable=False,
            cellStyle=status_style,
        )
    order_id_style = None
    if "_reimbursed" in df.columns:
        order_id_style = build_reimbursed_cell_style()
    if order_id_style:
        builder.configure_column(
            "order_id",
            header_name="订单号",
            width=140,
            cellStyle=order_id_style,
            checkboxSelection=bool(use_checkbox),
            headerCheckboxSelection=bool(use_checkbox and header_checkbox),
            headerCheckboxSelectionFilteredOnly=False,
        )
    else:
        builder.configure_column(
            "order_id",
            header_name="订单号",
            width=140,
            checkboxSelection=bool(use_checkbox),
            headerCheckboxSelection=bool(use_checkbox and header_checkbox),
            headerCheckboxSelectionFilteredOnly=False,
        )
    builder.configure_column("vendor", header_name="店铺", width=160)
    builder.configure_column("description", header_name="明细", width=260)
    link_formatter, link_class, link_tooltip = build_product_link_helpers()
    if link_formatter:
        builder.configure_column(
            "product_link",
            header_name="商品详情",
            width=110,
            editable=editable,
            valueFormatter=link_formatter,
            cellClass=link_class,
            tooltipValueGetter=link_tooltip,
        )
    else:
        builder.configure_column("product_link", header_name="商品链接", width=220)
    builder.configure_column(
        "amount",
        header_name="实付",
        width=110,
        type=["numericColumn", "numberColumnFilter"],
    )
    builder.configure_column(
        "order_status",
        header_name="状态",
        width=120,
        filter=False,
    )
    builder.configure_column("order_date", header_name="时间", width=120)
    builder.configure_column("_row_id", header_name="ID", hide=True, editable=False)
    if "order_row" in df.columns:
        builder.configure_column("order_row", hide=True, editable=False)
    if "_order_row" in df.columns:
        builder.configure_column("_order_row", hide=True, editable=False)
    builder.configure_selection(
        selection_mode,
        use_checkbox=use_checkbox,
        header_checkbox=header_checkbox,
    )
    builder.configure_grid_options(
        rowHeight=32,
        stopEditingWhenCellsLoseFocus=True,
    )
    link_click = build_product_link_click_handler()
    if link_click:
        builder.configure_grid_options(onCellClicked=link_click)
    focus_opts = build_row_focus_options(row_id_field, selected_row_id, force_focus)
    if focus_opts:
        builder.configure_grid_options(**focus_opts)
    return builder.build()


def build_product_link_helpers():
    if not JS_CODE_AVAILABLE or JsCode is None:
        return None, None, None
    try:
        formatter = JsCode(
            """
            function(params) {
                return params.value ? "商品详情" : "无";
            }
            """
        )
        cell_class = JsCode(
            """
            function(params) {
                return params.value ? "product-link-cell" : "product-link-empty";
            }
            """
        )
        tooltip = JsCode(
            """
            function(params) {
                return params.value ? String(params.value) : "";
            }
            """
        )
        return formatter, cell_class, tooltip
    except Exception:
        return None, None, None


def build_product_link_click_handler():
    if not JS_CODE_AVAILABLE or JsCode is None:
        return None
    try:
        return JsCode(
            """
            function(event) {
                if (!event || !event.colDef || event.colDef.field !== "product_link") {
                    return;
                }
                const raw = event.data ? event.data["product_link"] : event.value;
                if (!raw) {
                    return;
                }
                const text = String(raw);
                const url = encodeURI(text.split(" | ")[0]);
                window.open(url, "_blank", "noopener");
            }
            """
        )
    except Exception:
        return None


def build_screenshot_status_style():
    if not JS_CODE_AVAILABLE or JsCode is None:
        return None
    try:
        return JsCode(
            """
            function(params) {
                if (params.value === '无截图') {
                    return {backgroundColor:'#fee2e2', color:'#991b1b', fontWeight:'600',
                            textAlign:'center', cursor:'pointer', borderRadius:'4px'};
                }
                if (params.value === '截图不完全') {
                    return {backgroundColor:'#ffedd5', color:'#9a3412', fontWeight:'600',
                            textAlign:'center', cursor:'pointer', borderRadius:'4px'};
                }
                if (params.value === '有截图') {
                    return {backgroundColor:'#dcfce7', color:'#166534', fontWeight:'600',
                            textAlign:'center', cursor:'pointer', borderRadius:'4px'};
                }
                return {};
            }
            """
        )
    except Exception:
        return None


def build_reimbursed_cell_style():
    if not JS_CODE_AVAILABLE or JsCode is None:
        return None
    try:
        return JsCode(
            """
            function(params) {
                if (!params || !params.data || !params.data._reimbursed) {
                    return {};
                }
                return {
                    backgroundColor: '#ffedd5',
                    color: '#9a3412',
                    fontWeight: '700',
                    borderRadius: '4px'
                };
            }
            """
        )
    except Exception:
        return None


def build_row_focus_options(row_id_field, selected_row_id, force_focus=False):
    if not JS_CODE_AVAILABLE or JsCode is None or not row_id_field:
        return {}

    def_js = {}
    def_js["getRowId"] = JsCode(
        f"function(params) {{ return String(params.data.{row_id_field}); }}"
    )
    if not force_focus:
        return def_js
    selected_value = None
    if selected_row_id not in (None, ""):
        selected_value = str(selected_row_id)
    def_js["context"] = {"selectedRowId": selected_value}
    focus_js = JsCode(
        f"""
        function(params) {{
            const ctx = params.context || {{}};
            const target = ctx.selectedRowId;
            if (target === undefined || target === null || target === '') {{
                return;
            }}
            const targetStr = String(target);
            const node = params.api.getRowNode(targetStr);
            if (node) {{
                node.setSelected(true);
                params.api.ensureNodeVisible(node, 'middle');
                return;
            }}
            const count = params.api.getDisplayedRowCount();
            for (let i = 0; i < count; i++) {{
                const rowNode = params.api.getDisplayedRowAtIndex(i);
                if (rowNode && rowNode.data && String(rowNode.data.{row_id_field}) === targetStr) {{
                    rowNode.setSelected(true);
                    params.api.ensureIndexVisible(i, 'middle');
                    break;
                }}
            }}
        }}
        """
    )
    def_js["onFirstDataRendered"] = focus_js
    def_js["onGridReady"] = focus_js
    def_js["onRowDataUpdated"] = focus_js
    return def_js


def build_match_grid_options(
    df, color_map, selected_row_id=None, row_id_field=None, force_focus=False
):
    builder = GridOptionsBuilder.from_dataframe(df)
    builder.configure_default_column(
        editable=False,
        resizable=True,
        sortable=True,
        filter=True,
    )
    builder.configure_column("_order_row", header_name="ROW", hide=True)
    if "order_row" in df.columns:
        builder.configure_column("order_row", hide=True, editable=False)
    if "_focus_row" in df.columns:
        builder.configure_column("_focus_row", hide=True, editable=False)
    if "auto_unique_id" in df.columns:
        builder.configure_column(
            "auto_unique_id", header_name="auto_unique_id", width=90, editable=False
        )
    if "screenshot_status" in df.columns:
        status_style = build_screenshot_status_style()
        builder.configure_column(
            "screenshot_status",
            header_name="截图状态",
            width=120,
            editable=False,
            cellStyle=status_style,
        )
    if "order_date" in df.columns:
        builder.configure_column("order_date", header_name="时间", width=120)
    link_formatter, link_class, link_tooltip = build_product_link_helpers()
    if link_formatter:
        builder.configure_column(
            "product_link",
            header_name="商品详情",
            width=110,
            valueFormatter=link_formatter,
            cellClass=link_class,
            tooltipValueGetter=link_tooltip,
        )
    elif "product_link" in df.columns:
        builder.configure_column("product_link", header_name="商品链接", width=220)
    if JS_CODE_AVAILABLE and JsCode is not None:
        try:
            color_js = JsCode(
                f"""
                function(params) {{
                    if (params.value === '完美匹配') {{
                        return {{'backgroundColor': '{color_map.get("完美匹配", "#d6f5d6")}', 'color': '#000'}};
                    }}
                    if (params.value === '疑似匹配') {{
                        return {{'backgroundColor': '{color_map.get("疑似匹配", "#ffe9c6")}', 'color': '#000'}};
                    }}
                    if (params.value === '无匹配') {{
                        return {{'backgroundColor': '{color_map.get("无匹配", "#ffd6d6")}', 'color': '#000'}};
                    }}
                    return {{}};
                }}
                """
            )
            builder.configure_column("match_status", cellStyle=color_js)
        except Exception:
            pass
    builder.configure_selection("single", use_checkbox=False)
    builder.configure_grid_options(rowHeight=32)
    link_click = build_product_link_click_handler()
    if link_click:
        builder.configure_grid_options(onCellClicked=link_click)
    focus_opts = build_row_focus_options(row_id_field, selected_row_id, force_focus)
    if focus_opts:
        builder.configure_grid_options(**focus_opts)
    return builder.build()


def select_order_files():
    if not TK_AVAILABLE:
        st.warning("无法打开选择窗口，请手动输入路径。")
        return []
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    paths = filedialog.askopenfilenames(
        title="选择订单文件",
        filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
    )
    root.destroy()
    return list(paths) if paths else []


def select_pdf_folder():
    if not TK_AVAILABLE:
        st.warning("无法打开选择窗口，请手动输入路径。")
        return None
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askdirectory(title="选择发票文件夹")
    root.destroy()
    return path or None


def select_export_path(default_name):
    if not TK_AVAILABLE:
        st.warning("无法打开保存窗口，请手动下载或检查 Tk 环境。")
        return None
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.asksaveasfilename(
        title="保存导出文件",
        initialfile=default_name,
        defaultextension=".xlsx",
        filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
    )
    root.destroy()
    return path or None


def select_export_directory(title):
    if not TK_AVAILABLE:
        st.warning("无法打开目录选择窗口，请检查 Tk 环境。")
        return None
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askdirectory(title=title)
    root.destroy()
    return path or None


def copy_to_clipboard(text):
    if text is None:
        return False, "内容为空"
    value = str(text).strip()
    if not value:
        return False, "内容为空"
    if os.name == "nt":
        try:
            subprocess.run(["clip"], input=value, text=True, check=True)
            return True, None
        except Exception:
            pass
    if TK_AVAILABLE:
        try:
            root = tk.Tk()
            root.withdraw()
            root.clipboard_clear()
            root.clipboard_append(value)
            root.update()
            root.destroy()
            return True, None
        except Exception as exc:
            return False, str(exc)
    try:
        import pyperclip

        pyperclip.copy(value)
        return True, None
    except Exception:
        return False, "无法复制，请检查 Tk 环境或安装 pyperclip"


def sanitize_filename(value, default="未命名"):
    text = str(value or "").strip()
    if text.lower() in {"nan", "none"}:
        text = ""
    if not text:
        text = default
    text = re.sub(r'[<>:"/\\\\|?*]', "_", text)
    text = text.replace("\n", " ").replace("\r", " ")
    return text.strip() or default


def parse_amount_value(value):
    amount = matcher.parse_amount(value)
    if amount is not None:
        return float(amount)
    try:
        return float(value)
    except Exception:
        return None


def build_export_basename(total_amount):
    try:
        amount_value = float(total_amount)
    except Exception:
        amount_value = 0.0
    amount_text = f"{amount_value:.2f}"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"报销明细{amount_text}+{timestamp}", timestamp, amount_text


def build_print_basename(total_amount):
    try:
        amount_value = float(total_amount)
    except Exception:
        amount_value = 0.0
    amount_text = f"{amount_value:.2f}"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"报销明细打印 {amount_text} {timestamp}", timestamp, amount_text


def build_reimbursement_details_df(df):
    base_cols = [
        "order_id",
        "vendor",
        "description",
        "product_link",
        "amount",
        "order_status",
        "order_date",
        "invoice_id",
        "invoice_file",
    ]
    cols = [col for col in base_cols if col in df.columns]
    details_df = df[cols].copy() if cols else df.copy()
    order_counts = []
    payment_counts = []
    store = st.session_state.get("screenshot_store", {})
    for idx, row in df.iterrows():
        fallback_id = row.get("_row_id", idx)
        order_key = shots.build_order_key(row.get("order_id"), fallback_id)
        entry = store.get(order_key, {})
        order_counts.append(len(entry.get(shots.CATEGORY_ORDER, [])))
        payment_counts.append(len(entry.get(shots.CATEGORY_PAYMENT, [])))
    details_df["订单截图数"] = order_counts
    details_df["支付截图数"] = payment_counts
    return details_df


def export_reimbursement_bundle(df, export_dir):
    total_amount = sum_amounts(df.get("amount", []))
    base_name, timestamp, amount_text = build_export_basename(total_amount)
    export_dir = Path(export_dir)
    export_dir.mkdir(parents=True, exist_ok=True)
    missing_invoices = []

    with tempfile.TemporaryDirectory() as temp_dir:
        root_dir = Path(temp_dir) / base_name
        root_dir.mkdir(parents=True, exist_ok=True)

        details_df = build_reimbursement_details_df(df)
        summary_df = pd.DataFrame(
            [
                {
                    "订单数": len(df),
                    "总金额": format_amount(total_amount),
                    "导出时间": timestamp,
                }
            ]
        )
        excel_path = root_dir / f"{base_name}.xlsx"
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            details_df.to_excel(writer, index=False, sheet_name="details")
            summary_df.to_excel(writer, index=False, sheet_name="summary")

        store = st.session_state.get("screenshot_store", {})
        folder_counts = {}
        for idx, row in df.iterrows():
            amount_value = parse_amount_value(row.get("amount"))
            amount_name = f"{amount_value:.2f}" if amount_value is not None else "0.00"
            vendor_name = sanitize_filename(row.get("vendor"), default="未知店名")
            base_folder = f"{amount_name}+{vendor_name}"
            suffix = folder_counts.get(base_folder, 0)
            folder_counts[base_folder] = suffix + 1
            folder_name = (
                f"{base_folder}_{suffix + 1}" if suffix else base_folder
            )
            order_dir = root_dir / folder_name
            order_dir.mkdir(parents=True, exist_ok=True)

            invoice_path = row.get("invoice_file")
            if invoice_path is not None:
                try:
                    if pd.isna(invoice_path):
                        invoice_path = None
                except Exception:
                    pass
            if invoice_path:
                invoice_path = Path(str(invoice_path))
                if invoice_path.exists():
                    shutil.copy2(invoice_path, order_dir / invoice_path.name)
                else:
                    missing_invoices.append(str(invoice_path))

            fallback_id = row.get("_row_id", idx)
            order_key = shots.build_order_key(row.get("order_id"), fallback_id)
            entry = store.get(order_key, {})
            for i, img_bytes in enumerate(entry.get(shots.CATEGORY_ORDER, []), 1):
                (order_dir / f"订单截图_{i}.png").write_bytes(img_bytes)
            for i, img_bytes in enumerate(entry.get(shots.CATEGORY_PAYMENT, []), 1):
                (order_dir / f"支付截图_{i}.png").write_bytes(img_bytes)

        zip_path = export_dir / f"{base_name}.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for path in root_dir.rglob("*"):
                arcname = path.relative_to(root_dir.parent)
                zf.write(path, arcname)

    return zip_path, missing_invoices


def get_pdf_backend():
    try:
        from pypdf import PdfReader, PdfWriter

        return PdfReader, PdfWriter
    except Exception:
        try:
            from PyPDF2 import PdfReader, PdfWriter

            return PdfReader, PdfWriter
        except Exception:
            return None, None


def append_pdf_file(writer, reader_cls, path):
    reader = reader_cls(str(path))
    for page in reader.pages:
        writer.add_page(page)


def append_image_bytes(writer, reader_cls, img_bytes):
    if not PIL_AVAILABLE or Image is None:
        raise RuntimeError("缺少图片处理依赖，请安装 Pillow")
    image = Image.open(io.BytesIO(img_bytes))
    if image.mode in ("RGBA", "P"):
        image = image.convert("RGB")
    buffer = io.BytesIO()
    image.save(buffer, format="PDF")
    buffer.seek(0)
    reader = reader_cls(buffer)
    for page in reader.pages:
        writer.add_page(page)


def export_reimbursement_print_pdf(df, export_dir):
    total_amount = sum_amounts(df.get("amount", []))
    base_name, timestamp, amount_text = build_print_basename(total_amount)
    export_dir = Path(export_dir)
    export_dir.mkdir(parents=True, exist_ok=True)

    reader_cls, writer_cls = get_pdf_backend()
    if reader_cls is None:
        raise RuntimeError("缺少 PDF 合并依赖，请安装 pypdf 或 PyPDF2")

    writer = writer_cls()
    store = st.session_state.get("screenshot_store", {})
    missing_invoices = []
    image_errors = 0
    added_pages = 0

    for idx, row in df.iterrows():
        invoice_path = row.get("invoice_file")
        if invoice_path is not None:
            try:
                if pd.isna(invoice_path):
                    invoice_path = None
            except Exception:
                pass
        if invoice_path:
            invoice_path = Path(str(invoice_path))
            if invoice_path.exists():
                try:
                    before = len(writer.pages)
                    append_pdf_file(writer, reader_cls, invoice_path)
                    added_pages += len(writer.pages) - before
                except Exception:
                    missing_invoices.append(str(invoice_path))
            else:
                missing_invoices.append(str(invoice_path))

        fallback_id = row.get("_row_id", idx)
        order_key = shots.build_order_key(row.get("order_id"), fallback_id)
        entry = store.get(order_key, {})
        for img_bytes in entry.get(shots.CATEGORY_ORDER, []):
            try:
                before = len(writer.pages)
                append_image_bytes(writer, reader_cls, img_bytes)
                added_pages += len(writer.pages) - before
            except Exception:
                image_errors += 1
        for img_bytes in entry.get(shots.CATEGORY_PAYMENT, []):
            try:
                before = len(writer.pages)
                append_image_bytes(writer, reader_cls, img_bytes)
                added_pages += len(writer.pages) - before
            except Exception:
                image_errors += 1

    if added_pages == 0:
        raise RuntimeError("没有可导出的内容")

    pdf_path = export_dir / f"{base_name}.pdf"
    with open(pdf_path, "wb") as fh:
        writer.write(fh)

    return pdf_path, missing_invoices, image_errors


def render_preprocess_tab(orders_paths, config_path):
    st.subheader("订单预处理工作台")
    st.caption("先读取订单并完成去重/筛选，再进入发票匹配。")
    if st.session_state.preprocess_notice:
        st.success(st.session_state.preprocess_notice)
        st.session_state.preprocess_notice = ""
    if st.session_state.reimbursed_notice:
        st.success(st.session_state.reimbursed_notice)
        st.session_state.reimbursed_notice = ""

    if not AGGRID_AVAILABLE:
        st.error("未安装可编辑表格组件：streamlit-aggrid")
        st.caption("安装命令：pip install streamlit-aggrid")
        return

    current_df = ensure_editor_df(st.session_state.order_df)
    st.session_state.order_df = current_df

    reimbursed_ids = set(st.session_state.get("reimbursed_order_ids", []))
    reimbursed_fallback_keys = set(st.session_state.get("reimbursed_fallback_keys", []))
    reimbursed_mask_all = compute_reimbursed_mask(
        current_df, reimbursed_ids, reimbursed_fallback_keys
    )
    reimbursed_match_count = int(reimbursed_mask_all.sum())
    reimbursed_total = len(reimbursed_ids) + len(reimbursed_fallback_keys)

    total_amount = sum_amounts(current_df.get("amount", []))
    date_span_text, _span_days = compute_order_date_span(current_df.get("order_date", []))
    vendor_count = int(
        current_df["vendor"]
        .fillna("")
        .astype(str)
        .map(lambda value: value.strip())
        .replace("", pd.NA)
        .dropna()
        .nunique()
    )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("订单总数", f"{len(current_df)}")
    with m2:
        render_date_span_card(date_span_text)
    m3.metric("订单总金额", format_amount(total_amount))
    m4.metric("已报销命中", f"{reimbursed_match_count}")
    st.caption(f"涉及店铺：{vendor_count} 家")

    st.markdown("#### 1. 数据导入")
    st.caption("左侧导入订单/发票并解析，右侧单独导入已报销明细。")
    import_left_col, import_divider_col, import_right_col = st.columns([1.35, 0.05, 1.1])
    with import_left_col:
        with st.container(border=True):
            st.markdown("##### 订单与发票")
            select_col1, select_col2 = st.columns([1, 3])
            with select_col1:
                if st.button("选订单", key="preprocess_pick_orders"):
                    chosen = select_order_files()
                    if chosen:
                        st.session_state.orders_paths_text = "\n".join(chosen)
            with select_col2:
                st.text_area("订单文件(多选)", key="orders_paths_text", height=86)

            select_col3, select_col4 = st.columns([1, 3])
            with select_col3:
                if st.button("选发票", key="preprocess_pick_pdfs"):
                    chosen = select_pdf_folder()
                    if chosen:
                        st.session_state.pdf_dir = chosen
            with select_col4:
                st.text_input("发票文件夹", key="pdf_dir")

            st.checkbox("递归扫描发票目录", key="recursive_scan")

            current_orders_paths = parse_orders_paths(st.session_state.orders_paths_text)
            if st.button("读取并解析订单", key="preprocess_parse_orders", use_container_width=True):
                try:
                    config = matcher.load_config(config_path)
                    if not current_orders_paths:
                        st.error("请选择订单文件")
                        return
                    orders_df = matcher.load_orders_multi(current_orders_paths, config)
                    editor_df = orders_df_for_editor(orders_df)
                    editor_df, removed_count = deduplicate_editor_orders(editor_df)
                    notice = f"已解析 {len(editor_df)} 条"
                    if removed_count:
                        notice += f"，自动去重 {removed_count} 条"
                    apply_preprocess_change(
                        editor_df,
                        notice=notice,
                        reset_grid=True,
                        reset_row_ids=True,
                        track_history=True,
                    )
                except Exception as exc:
                    st.error(f"解析失败：{exc}")
    with import_divider_col:
        st.markdown('<div class="import-divider"></div>', unsafe_allow_html=True)
    with import_right_col:
        with st.container(border=True):
            st.markdown("##### 导入已报销订单")
            st.caption("支持 xlsx/xls/csv/zip，zip 内会自动提取报销明细表。")
            reimbursed_upload = st.file_uploader(
                "选择报销明细文件",
                type=["xlsx", "xls", "csv", "zip"],
                key="preprocess_reimbursed_upload",
            )
            if st.button(
                "读取已报销订单",
                key="preprocess_read_reimbursed",
                disabled=reimbursed_upload is None,
                use_container_width=True,
            ):
                try:
                    reimbursed_df = load_reimbursed_orders_file(reimbursed_upload)
                    reimbursed_matchers = collect_reimbursed_matchers(reimbursed_df)
                    reimbursed_ids = reimbursed_matchers["order_ids"]
                    reimbursed_fallback_keys = reimbursed_matchers["fallback_keys"]
                    st.session_state.reimbursed_order_ids = reimbursed_ids
                    st.session_state.reimbursed_fallback_keys = reimbursed_fallback_keys
                    st.session_state.reimbursed_source_name = reimbursed_upload.name
                    st.session_state.reimbursed_source_rows = len(reimbursed_df)
                    total_reimbursed = len(reimbursed_ids) + len(reimbursed_fallback_keys)
                    st.session_state.reimbursed_notice = (
                        f"已导入已报销订单：{total_reimbursed} 个"
                    )
                    trigger_rerun()
                except Exception as exc:
                    st.error(f"导入失败：{exc}")
            if reimbursed_upload is not None:
                st.caption(f"已选择：{reimbursed_upload.name}")

    if reimbursed_total:
        source_name = st.session_state.get("reimbursed_source_name") or "已导入文件"
        st.info(
            f"已报销订单：{reimbursed_total} 个（当前预处理命中 {reimbursed_match_count} 个）"
        )
        st.caption(f"来源：{source_name}")
        if reimbursed_fallback_keys:
            st.caption(f"其中无订单号兜底匹配：{len(reimbursed_fallback_keys)} 个")
        st.caption("已报销订单会在“订单号”列高亮显示。")

    st.markdown("#### 2. 编辑与缓存")
    st.caption("支持自动保存与最多 20 步撤回/重做，确保预处理过程可回退。")
    undo_count = len(st.session_state.get("preprocess_undo_stack", []))
    redo_count = len(st.session_state.get("preprocess_redo_stack", []))
    tools_col1, tools_col2, tools_col3, tools_col4 = st.columns([1.2, 1, 1, 1.2])
    with tools_col1:
        if st.button(
            "暂存",
            key="preprocess_manual_cache",
            use_container_width=True,
            disabled=bool(st.session_state.get("preprocess_auto_save")),
        ):
            try:
                save_preprocess_cache()
                st.session_state.preprocess_notice = "暂存成功"
                trigger_rerun()
            except Exception as exc:
                st.error(f"暂存失败：{exc}")
    with tools_col2:
        if st.button(
            "撤回",
            key="preprocess_undo_action",
            use_container_width=True,
            disabled=undo_count == 0,
        ):
            if undo_preprocess_change():
                trigger_rerun()
    with tools_col3:
        if st.button(
            "重做",
            key="preprocess_redo_action",
            use_container_width=True,
            disabled=redo_count == 0,
        ):
            if redo_preprocess_change():
                trigger_rerun()
    with tools_col4:
        st.checkbox(
            "自动保存",
            key="preprocess_auto_save",
            on_change=on_preprocess_auto_save_toggle,
        )
    st.caption(
        f"撤回/重做最多保留 {PREPROCESS_HISTORY_LIMIT} 步：可撤回 {undo_count}，可重做 {redo_count}"
    )
    if st.session_state.get("preprocess_auto_save"):
        st.caption("自动保存已开启，手动“暂存”按钮已禁用。")

    st.markdown("#### 3. 预处理订单表格")
    filter_col, hint_col = st.columns([2.4, 1.6])
    status_options = [
        value
        for value in sorted(
            {
                str(v).strip()
                for v in current_df["order_status"].dropna().tolist()
                if str(v).strip()
            }
        )
    ]
    with filter_col:
        if status_options:
            selected_statuses = st.multiselect(
                "状态筛选",
                options=status_options,
                default=status_options,
                key="order_status_filter",
            )
        else:
            selected_statuses = []
            st.caption("当前无状态字段内容，默认展示全部订单。")
    with hint_col:
        st.caption("双击单元格可编辑；勾选行后可批量删除。")

    if status_options and selected_statuses:
        display_base_df = current_df[current_df["order_status"].isin(selected_statuses)]
    elif status_options and not selected_statuses:
        display_base_df = current_df.iloc[0:0]
    else:
        display_base_df = current_df

    selected_row_id = st.session_state.get("preprocess_selected_row_id")
    display_base_df = add_reimbursed_flag_column(
        display_base_df, reimbursed_ids, reimbursed_fallback_keys
    )
    display_df = add_auto_unique_id_column(display_base_df)

    grid_key = f"order_grid_{st.session_state.order_grid_version}"
    grid_response = AgGrid(
        display_df,
        gridOptions=build_order_grid_options(
            display_df,
            True,
            selected_row_id,
            "_row_id",
            selection_mode="multiple",
            use_checkbox=True,
            header_checkbox=True,
        ),
        data_return_mode=DataReturnMode.AS_INPUT,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        update_on=["cellValueChanged", "selectionChanged"],
        fit_columns_on_grid_load=True,
        theme="balham",
        key=grid_key,
        allow_unsafe_jscode=True,
        custom_css=GRID_CUSTOM_CSS,
    )

    updated_data = grid_response.get("data")
    if updated_data is not None:
        if isinstance(updated_data, list):
            updated_df = pd.DataFrame(updated_data)
        else:
            updated_df = updated_data
        updated_clean = ensure_editor_df(updated_df)
        display_clean = ensure_editor_df(display_base_df)
        if not updated_clean.equals(display_clean):
            if display_base_df.shape[0] != current_df.shape[0]:
                next_df = merge_grid_updates(current_df, updated_clean)
            else:
                next_df = updated_clean
            apply_preprocess_change(
                next_df,
                reset_grid=False,
                reset_row_ids=False,
                track_history=True,
            )

    selected_rows = grid_response.get("selected_rows")
    if selected_rows is None:
        selected_rows = []
    elif isinstance(selected_rows, pd.DataFrame):
        selected_rows = selected_rows.to_dict("records")
    selected_ids = set()
    for row in selected_rows:
        if "_row_id" not in row:
            continue
        value = row.get("_row_id")
        try:
            value = int(value)
        except Exception:
            pass
        selected_ids.add(value)

    action_col1, action_col2, action_col3, action_col4 = st.columns(
        [1, 1, 1.3, 1.8]
    )
    with action_col1:
        if st.button("删除选中", key="preprocess_delete_selected", use_container_width=True):
            if not selected_ids:
                st.warning("请先勾选要删除的订单")
            else:
                new_df = st.session_state.order_df[
                    ~st.session_state.order_df["_row_id"].isin(selected_ids)
                ]
                apply_preprocess_change(
                    new_df,
                    notice=f"已删除 {len(selected_ids)} 行",
                    reset_grid=True,
                    track_history=True,
                )
                trigger_rerun()
    with action_col2:
        if st.button("清空", key="preprocess_clear_all", use_container_width=True):
            empty_df = pd.DataFrame(columns=EDITOR_COLUMNS)
            apply_preprocess_change(
                empty_df,
                notice="已清空",
                reset_grid=True,
                reset_row_ids=True,
                track_history=True,
            )
            trigger_rerun()
    with action_col3:
        if st.button(
            "移除已报销订单",
            key="preprocess_remove_reimbursed",
            use_container_width=True,
            disabled=reimbursed_match_count == 0,
        ):
            new_df = st.session_state.order_df.loc[~reimbursed_mask_all].copy()
            apply_preprocess_change(
                new_df,
                notice=f"已移除 {reimbursed_match_count} 条已报销订单",
                reset_grid=True,
                track_history=True,
            )
            trigger_rerun()
    with action_col4:
        if st.button(
            "开始发票匹配",
            key="preprocess_start_match",
            type="primary",
            use_container_width=True,
            disabled=st.session_state.order_df is None or st.session_state.order_df.empty,
        ):
            st.session_state.pending_tab = "匹配"
            st.session_state.pending_match_auto_start = True
            trigger_rerun()

    tool_tab_export, tool_tab_add = st.tabs(["导出预处理结果", "快速新增订单"])
    with tool_tab_export:
        export_config = load_config(config_path)
        export_df = build_export_df_with_config(st.session_state.order_df, export_config)
        export_col1, export_col2, export_col3 = st.columns([1, 1, 2])
        with export_col1:
            csv_data = export_df.to_csv(index=False, encoding="utf-8-sig")
            st.download_button(
                "导出 CSV",
                data=csv_data,
                file_name="订单预筛选.csv",
                mime="text/csv",
                disabled=export_df.empty,
                use_container_width=True,
            )
        with export_col2:
            if st.button("导出 Excel(另存为)", key="preprocess_export_excel_dialog", use_container_width=True):
                if export_df.empty:
                    st.warning("没有可导出的数据")
                else:
                    save_path = select_export_path("订单预筛选.xlsx")
                    if save_path:
                        if not save_path.lower().endswith(".xlsx"):
                            save_path += ".xlsx"
                        try:
                            export_df.to_excel(save_path, index=False, sheet_name="orders")
                            st.success(f"已导出：{save_path}")
                        except Exception as exc:
                            st.warning(f"导出 Excel 失败：{exc}")
            if not TK_AVAILABLE:
                if export_df.empty:
                    st.download_button(
                        "导出 Excel",
                        data=b"",
                        file_name="订单预筛选.xlsx",
                        disabled=True,
                        use_container_width=True,
                    )
                else:
                    try:
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                            export_df.to_excel(writer, index=False, sheet_name="orders")
                        st.download_button(
                            "导出 Excel",
                            data=buffer.getvalue(),
                            file_name="订单预筛选.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                    except Exception as exc:
                        st.warning(f"导出 Excel 失败：{exc}")
        with export_col3:
            st.caption(f"当前 {len(export_df)} 行")
            total_export_amount = sum_amounts(st.session_state.order_df.get("amount", []))
            st.caption(f"订单总金额：{format_amount(total_export_amount)}")

    with tool_tab_add:
        st.caption("手工补录单条订单信息，提交后会进入预处理表格。")
        if st.session_state.get("pending_clear_add_order_form", False):
            st.session_state["new_order_id"] = ""
            st.session_state["new_vendor"] = ""
            st.session_state["new_desc"] = ""
            st.session_state["new_amount"] = ""
            st.session_state["new_date"] = ""
            st.session_state["new_link"] = ""
            st.session_state["pending_clear_add_order_form"] = False
        with st.form("preprocess_add_order_form", clear_on_submit=False):
            new_cols = st.columns(3)
            with new_cols[0]:
                new_order_id = st.text_input("订单号", key="new_order_id")
                new_vendor = st.text_input("店铺", key="new_vendor")
            with new_cols[1]:
                new_desc = st.text_input("明细", key="new_desc")
                new_amount = st.text_input("实付", key="new_amount")
            with new_cols[2]:
                new_date = st.text_input("时间", key="new_date")
                new_link = st.text_input("商品链接", key="new_link")
                st.caption("交易状态默认：交易成功")

            submitted = st.form_submit_button(
                "添加订单", key="preprocess_add_order", type="secondary"
            )

        if submitted:
            formatted_order_id = matcher.format_order_id(new_order_id)
            if not formatted_order_id:
                st.warning("请填写订单号后再添加，避免被误判为已报销订单。")
            else:
                amount_value = matcher.parse_amount(new_amount) if new_amount else None
                next_row_id = st.session_state.order_row_id_seq
                new_row = {
                    "order_id": formatted_order_id,
                    "vendor": new_vendor.strip(),
                    "description": new_desc.strip(),
                    "product_link": new_link.strip(),
                    "amount": amount_value if amount_value is not None else new_amount,
                    "order_status": "交易成功",
                    "order_date": new_date.strip(),
                    "_row_id": next_row_id,
                }
                new_df = pd.concat(
                    [st.session_state.order_df, pd.DataFrame([new_row])],
                    ignore_index=True,
                )
                apply_preprocess_change(
                    new_df,
                    notice="已添加",
                    reset_grid=True,
                    reset_row_ids=False,
                    track_history=True,
                )
                st.session_state["pending_clear_add_order_form"] = True
                trigger_rerun()


def render_match_tab(orders_paths, pdf_dir, config_path, output_path, recursive, color_map):
    st.subheader("发票匹配与复核")
    orders_hint = f"{len(orders_paths)} 个文件" if orders_paths else "未选择"
    st.caption(f"订单：{orders_hint} ｜ 发票：{pdf_dir}")
    matched_orders_for_screenshots = pd.DataFrame()

    auto_start_match = bool(st.session_state.pop("pending_match_auto_start", False))
    if st.button("开始匹配") or auto_start_match:
        config = load_config(config_path)
        try:
            backend = matcher.resolve_pdf_backend()
            if backend is None:
                st.error("缺少 PDF 解析库")
                return

            if st.session_state.order_df is not None and not st.session_state.order_df.empty:
                order_rows = editor_df_to_rows(st.session_state.order_df)
                orders_df = matcher.build_orders_df_from_rows(order_rows)
            else:
                if not orders_paths:
                    st.error("请选择订单文件")
                    return
                orders_df = matcher.load_orders_multi(orders_paths, config)

            pdf_root = Path(pdf_dir)
            if not pdf_root.exists():
                st.error("发票路径无效")
                return

            pdfs = matcher.find_pdfs(pdf_dir, recursive)
            if not pdfs:
                st.error("未找到 PDF")
                return

            progress = st.progress(0)
            invoices = []
            for idx, path in enumerate(pdfs, 1):
                invoices.append(matcher.extract_invoice(path, config, backend))
                progress.progress(idx / len(pdfs))

            matches = matcher.match_orders_invoices(orders_df, invoices, config)
            status_map = matcher.compute_order_match_statuses(
                orders_df, invoices, matches, config
            )
            orders_df["_match_status"] = orders_df.index.map(
                lambda idx: status_map.get(idx, {}).get("status", "")
            )
            orders_df["_match_reason"] = orders_df.index.map(
                lambda idx: status_map.get(idx, {}).get("reason", "")
            )
            matcher.write_output(output_path, orders_df, invoices, matches)

            result_df = pd.DataFrame(
                {
                    "_order_row": orders_df.index.astype(int),
                    "order_id": orders_df["_order_id"].apply(matcher.format_order_id),
                    "vendor": orders_df["_vendor"],
                    "description": orders_df["_description"],
                    "product_link": orders_df.get("_product_link", ""),
                    "amount": orders_df["_amount"],
                    "order_status": orders_df["_status"],
                    "order_date": orders_df["_date"],
                    "match_status": orders_df["_match_status"],
                    "match_reason": orders_df["_match_reason"],
                }
            )
            result_df["order_date"] = result_df["order_date"].apply(format_editor_date)

            st.session_state.match_result_df = result_df
            st.session_state.match_detail = {
                "suspect_map": build_suspect_invoice_map(
                    orders_df, invoices, config, matches
                ),
                "order_invoice_map": build_order_invoice_map(matches),
                "unmatched_invoices": [
                    inv for inv in invoices if inv.get("match_order_row") is None
                ],
            }
            order_total = sum_amounts(orders_df["_amount"])
            perfect_total = sum_amounts(
                orders_df.loc[orders_df["_match_status"] == "完美匹配", "_amount"]
            )
            st.session_state.summary = {
                "orders": len(orders_df),
                "invoices": len(invoices),
                "matches": len(matches),
                "output": output_path,
                "order_total": order_total,
                "perfect_total": perfect_total,
            }
            st.success("完成")
        except Exception as exc:
            st.error(f"失败：{exc}")

    if st.session_state.summary:
        summary = st.session_state.summary
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("订单", summary["orders"])
        m2.metric("发票", summary["invoices"])
        m3.metric("匹配", summary["matches"])
        m4.metric("订单总金额", format_amount(summary.get("order_total", 0.0)))
        m5.metric("完美匹配金额", format_amount(summary.get("perfect_total", 0.0)))
        st.caption(f"输出：{Path(summary['output']).name}")

    if st.session_state.match_result_df is not None:
        st.markdown("### 结果")
        st.caption("点击“疑似匹配”行可查看疑似发票列表。")
        status_filter = st.selectbox(
            "状态筛选",
            options=["全部"]
            + sorted(
                {
                    value
                    for value in st.session_state.match_result_df["order_status"].dropna().unique()
                    if str(value).strip()
                }
            ),
        )
        view_df = st.session_state.match_result_df.copy()
        if status_filter != "全部":
            view_df = view_df[view_df["order_status"] == status_filter]
        view_df = add_auto_unique_id_column(view_df)

        if AGGRID_AVAILABLE:
            selected_row_id = st.session_state.get("match_selected_row_id")
            grid_response = AgGrid(
                view_df,
                gridOptions=build_match_grid_options(
                    view_df, color_map, selected_row_id, "_order_row"
                ),
                data_return_mode=DataReturnMode.AS_INPUT,
                update_mode=GridUpdateMode.SELECTION_CHANGED,
                fit_columns_on_grid_load=True,
                theme="balham",
                key="match_result_grid",
                height=520,
                allow_unsafe_jscode=True,
                custom_css=GRID_CUSTOM_CSS,
            )
            selected_rows = grid_response.get("selected_rows")
            if selected_rows is None:
                selected_rows = []
            elif isinstance(selected_rows, pd.DataFrame):
                selected_rows = selected_rows.to_dict("records")

            selected = resolve_selected_row(
                selected_rows, view_df, "_order_row", "match_selected_row_id"
            )
            if selected:
                if selected.get("match_status") == "疑似匹配":
                    detail = st.session_state.get("match_detail", {})
                    suspect_map = detail.get("suspect_map", {})
                    order_row = selected.get("_order_row")
                    suspects = suspect_map.get(int(order_row), []) if order_row is not None else []
                    with st.expander("疑似发票", expanded=True):
                        if not suspects:
                            st.caption("未找到疑似发票")
                        else:
                            for idx, invoice in enumerate(suspects):
                                label = describe_invoice(invoice)
                                if st.button(
                                    label,
                                    key=f"suspect_open_{order_row}_{idx}",
                                ):
                                    open_pdf_file(invoice.get("invoice_file", ""))

        else:
            st.dataframe(
                view_df.style.apply(lambda r: style_match_rows(r, color_map), axis=1),
                use_container_width=True,
                height=520,
            )

        matched_orders = view_df[view_df["match_status"] == "完美匹配"].copy()

        if not matched_orders.empty:
            order_invoice_map = st.session_state.match_detail.get("order_invoice_map", {})
            matched_orders["invoice_file"] = matched_orders["_order_row"].map(
                lambda idx: order_invoice_map.get(int(idx), {}).get("invoice_file", "")
            )
            matched_orders["invoice_id"] = matched_orders["_order_row"].map(
                lambda idx: order_invoice_map.get(int(idx), {}).get("invoice_id", "")
            )
        matched_orders_for_screenshots = matched_orders.copy()
        export_col1, export_col2, export_col3 = st.columns([1, 1, 2])
        with export_col1:
            csv_data = matched_orders.to_csv(index=False, encoding="utf-8-sig")
            st.download_button(
                "导出已匹配订单 CSV",
                data=csv_data,
                file_name="已匹配订单.csv",
                mime="text/csv",
                disabled=matched_orders.empty,
            )
        with export_col2:
            if matched_orders.empty:
                st.download_button(
                    "导出已匹配订单 Excel",
                    data=b"",
                    file_name="已匹配订单.xlsx",
                    disabled=True,
                )
            else:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    matched_orders.to_excel(
                        writer, index=False, sheet_name="matched_orders"
                    )
                st.download_button(
                    "导出已匹配订单 Excel",
                    data=buffer.getvalue(),
                    file_name="已匹配订单.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        with export_col3:
            st.caption(f"已匹配 {len(matched_orders)} 行")

    if st.session_state.match_result_df is not None:
        detail = st.session_state.get("match_detail", {})
        unmatched_invoices = detail.get("unmatched_invoices", [])
        if unmatched_invoices:
            with st.expander(f"未匹配发票 ({len(unmatched_invoices)})", expanded=False):
                for idx, invoice in enumerate(unmatched_invoices):
                    label = describe_invoice(invoice)
                    if st.button(label, key=f"unmatched_open_{idx}"):
                        open_pdf_file(invoice.get("invoice_file", ""))

    st.divider()
    jump_disabled = (
        st.session_state.match_result_df is None
        or matched_orders_for_screenshots.empty
    )
    if st.button("进入截图管理", disabled=jump_disabled):
        st.session_state.screenshot_prefill_df = matched_orders_for_screenshots
        st.session_state.pending_tab = "截图"
        trigger_rerun()
    if st.session_state.match_result_df is None:
        st.caption("请先完成匹配后进入截图管理。")
    elif matched_orders_for_screenshots.empty:
        st.caption("没有可进入截图管理的完美匹配订单。")


def load_matched_orders_file(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)


def prepare_screenshot_orders_df(df):
    return ensure_screenshot_df_for_cache(df)


def render_screenshot_tab():
    st.subheader("截图资料归档")
    if st.session_state.get("screenshot_notice"):
        st.success(st.session_state.screenshot_notice)
        st.session_state.screenshot_notice = ""

    if not AGGRID_AVAILABLE:
        st.error("未安装可编辑表格组件：streamlit-aggrid")
        st.caption("安装命令：pip install streamlit-aggrid")
        return

    uploaded = st.file_uploader("导入已匹配订单文件", type=["xlsx", "csv"])
    prefill_df = st.session_state.get("screenshot_prefill_df")
    cached_df = st.session_state.get("screenshot_orders_df")
    if uploaded:
        try:
            df = load_matched_orders_file(uploaded)
            st.session_state.screenshot_prefill_df = None
        except Exception as exc:
            st.error(f"导入失败：{exc}")
            return
    elif prefill_df is not None:
        df = prefill_df.copy()
        st.session_state.screenshot_prefill_df = None
        if df.empty:
            st.warning("匹配结果为空，请先完成匹配或手动导入")
            return
        st.caption("已从匹配结果加载，可继续或上传文件替换。")
    elif isinstance(cached_df, pd.DataFrame) and not cached_df.empty:
        df = cached_df.copy()
        st.caption("已从本地截图暂存加载，可继续或上传文件替换。")
    else:
        st.caption("请上传已匹配订单文件（由匹配页导出）。")
        return

    df = prepare_screenshot_orders_df(df)
    st.session_state.screenshot_orders_df = df.copy()

    st.caption(f"已导入 {len(df)} 行")
    st.caption("点击“截图状态”可进入截图管理。")
    tools_col1, tools_col2 = st.columns([1, 1.2])
    with tools_col1:
        if st.button(
            "暂存",
            key="screenshot_manual_save",
            disabled=bool(st.session_state.get("screenshot_auto_save")),
        ):
            try:
                snapshot = snapshot_from_screenshot_state(df=df)
                save_screenshot_cache(snapshot)
                st.session_state.screenshot_last_saved_signature = (
                    screenshot_snapshot_signature(snapshot)
                )
                st.session_state.screenshot_notice = "暂存成功"
                trigger_rerun()
            except Exception as exc:
                st.error(f"暂存失败：{exc}")
    with tools_col2:
        st.checkbox(
            "自动保存",
            key="screenshot_auto_save",
            on_change=on_screenshot_auto_save_toggle,
        )
    if st.session_state.get("screenshot_auto_save"):
        st.caption("自动保存已开启，已禁用手动“暂存”按钮。")

    display_df = add_screenshot_status_column(df, "order_id", "_row_id")
    display_df = add_auto_unique_id_column(display_df)

    selected_row_id = st.session_state.get("screenshot_selected_row_id")
    force_focus = bool(st.session_state.get("screenshot_selected_row_id_lock"))
    display_df = add_focus_row_column(
        display_df, "_row_id", selected_row_id, force_focus
    )
    grid_response = AgGrid(
        display_df,
        gridOptions=build_order_grid_options(
            display_df,
            editable=False,
            selected_row_id=selected_row_id,
            row_id_field="_row_id",
            force_focus=force_focus,
        ),
        data_return_mode=DataReturnMode.AS_INPUT,
        update_mode=GridUpdateMode.SELECTION_CHANGED,
        fit_columns_on_grid_load=True,
        theme="balham",
        key="screenshot_orders_grid",
        height=520,
        allow_unsafe_jscode=True,
        custom_css=GRID_CUSTOM_CSS,
    )

    selected_rows = grid_response.get("selected_rows")
    if selected_rows is None:
        selected_rows = []
    elif isinstance(selected_rows, pd.DataFrame):
        selected_rows = selected_rows.to_dict("records")

    selected = resolve_selected_row(
        selected_rows, display_df, "_row_id", "screenshot_selected_row_id"
    )

    with st.expander("截图管理", expanded=selected is not None):
        if not selected:
            st.info("请先在表格中点击一条订单的“截图状态”。")
        else:
            order_id = selected.get("order_id", "")
            order_row = selected.get("_row_id")
            order_key = shots.build_order_key(order_id, order_row)
            header_col1, header_col2 = st.columns([4, 1], gap="small")
            with header_col1:
                label = (
                    f"订单号：{order_id}" if str(order_id).strip() else f"订单行：{order_row}"
                )
                st.markdown(
                    f"<div style='font-size:20px;font-weight:600'>{label}</div>",
                    unsafe_allow_html=True,
                )
            with header_col2:
                copy_disabled = not bool(str(order_id).strip())
                if st.button(
                    "复制",
                    key=f"copy_order_id_{order_row}",
                    disabled=copy_disabled,
                ):
                    ok, err = copy_to_clipboard(order_id)
                    if ok:
                        st.success("已复制订单号")
                    else:
                        st.warning(f"复制失败：{err}")
            row_ids = display_df["_row_id"].tolist()
            next_action = make_step_order_callback(row_ids, "screenshot_selected_row_id", 1)
            prev_action = make_step_order_callback(row_ids, "screenshot_selected_row_id", -1)
            shots.render_screenshot_manager(
                order_key, on_next_order=next_action, on_prev_order=prev_action
            )

    st.session_state.screenshot_orders_df = df.copy()
    maybe_auto_save_screenshot()

    st.subheader("导出")
    export_col1, export_col2, export_col3 = st.columns([1, 1, 3])
    with export_col1:
        if st.button("全部导出"):
            export_dir = select_export_directory("选择导出目录")
            if export_dir:
                try:
                    zip_path, missing = export_reimbursement_bundle(df, export_dir)
                    st.success(f"已导出：{zip_path}")
                    if missing:
                        st.warning(f"未找到发票文件：{len(missing)} 个")
                except Exception as exc:
                    st.error(f"导出失败：{exc}")
    with export_col2:
        if st.button("打印导出"):
            export_dir = select_export_directory("选择导出目录")
            if export_dir:
                try:
                    pdf_path, missing, image_errors = export_reimbursement_print_pdf(
                        df, export_dir
                    )
                    st.success(f"已导出：{pdf_path}")
                    if missing:
                        st.warning(f"未找到发票文件：{len(missing)} 个")
                    if image_errors:
                        st.warning(f"有 {image_errors} 张截图无法合并")
                except Exception as exc:
                    st.error(f"导出失败：{exc}")
    with export_col3:
        st.caption("压缩包与打印 PDF 的名称都包含总金额与当前时间。")


def main():
    st.set_page_config(
        page_title="报销助手",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    ensure_session_state()
    start_auto_shutdown_monitor()
    inject_ui_styles()
    render_page_hero()

    with st.sidebar:
        st.header("设置")
        with st.expander("全局设置", expanded=True):
            st.text_input("配置文件", key="config_path")
            render_sidebar_settings(st.session_state.config_path)
            st.subheader("颜色")
            st.color_picker("完美", key="color_perfect")
            st.color_picker("疑似", key="color_suspect")
            st.color_picker("无匹配", key="color_nomatch")

    orders_paths = parse_orders_paths(st.session_state.orders_paths_text)
    pdf_dir = st.session_state.pdf_dir
    recursive = bool(st.session_state.get("recursive_scan", False))
    config_path = st.session_state.config_path
    output_path = st.session_state.output_path

    color_map = {
        "完美匹配": st.session_state.color_perfect,
        "疑似匹配": st.session_state.color_suspect,
        "无匹配": st.session_state.color_nomatch,
    }

    active_tab = render_main_navigation()

    if active_tab == "预处理":
        render_preprocess_tab(orders_paths, config_path)
    elif active_tab == "匹配":
        render_match_tab(
            orders_paths,
            pdf_dir,
            config_path,
            output_path,
            recursive,
            color_map,
        )
    else:
        render_screenshot_tab()

if __name__ == "__main__":
    main()
