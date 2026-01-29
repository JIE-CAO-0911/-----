#!/usr/bin/env python3
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
AUTO_SHUTDOWN_IDLE_SECONDS = 1
AUTO_SHUTDOWN_CHECK_SECONDS = 0.5
AUTO_SHUTDOWN_STARTED = False
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


def ensure_session_state():
    if "order_df" not in st.session_state:
        st.session_state.order_df = pd.DataFrame(
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
    if "config_path" not in st.session_state:
        st.session_state.config_path = str(DEFAULT_CONFIG_PATH)
    if "output_path" not in st.session_state:
        st.session_state.output_path = str(DEFAULT_OUTPUT_PATH)
    if "suspect_threshold" not in st.session_state:
        st.session_state.suspect_threshold = 0.6
    if "color_perfect" not in st.session_state:
        st.session_state.color_perfect = "#d6f5d6"
    if "color_suspect" not in st.session_state:
        st.session_state.color_suspect = "#ffe9c6"
    if "color_nomatch" not in st.session_state:
        st.session_state.color_nomatch = "#ffd6d6"
    if "preprocess_notice" not in st.session_state:
        st.session_state.preprocess_notice = ""
    if "active_tab" not in st.session_state:
        st.session_state.active_tab = "预处理"
    if "screenshot_prefill_df" not in st.session_state:
        st.session_state.screenshot_prefill_df = None
    if "pending_tab" not in st.session_state:
        st.session_state.pending_tab = None


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
    rules = (config or {}).get("match_rules", {})
    vendor_threshold = float(rules.get("vendor_similarity_threshold", 0.6))

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
        else:
            order_vendor = order.get("_vendor")
            if order_vendor:
                for invoice in invoices:
                    score = matcher.score_vendor(invoice.get("vendor"), order_vendor)
                    if score is not None and score >= vendor_threshold:
                        suspects.append(invoice)

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
    builder.configure_column("order_id", header_name="订单号", width=140)
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


def render_config_tab(config_path):
    config = load_config(config_path)

    st.subheader("列映射")
    order_cols = config.get("order_columns", {})
    col1, col2 = st.columns(2)
    with col1:
        order_id = st.text_input("订单号", value=order_cols.get("order_id", ""))
        amount = st.text_input("实付", value=order_cols.get("amount", ""))
        date_col = st.text_input("时间", value=order_cols.get("date", ""))
    with col2:
        vendor = st.text_input("店铺", value=order_cols.get("vendor", ""))
        description = st.text_input("商品", value=order_cols.get("description", ""))
        status = st.text_input("状态", value=order_cols.get("status", ""))
        product_link = st.text_input("商品链接", value=order_cols.get("product_link", ""))

    st.subheader("规则")
    rules = config.get("match_rules", {})
    r1, r2, r3 = st.columns(3)
    with r1:
        amount_tol = st.number_input(
            "金额容差", value=float(rules.get("amount_tolerance", 0.01)), step=0.01
        )
        min_score = st.number_input(
            "最低分", value=float(rules.get("min_score", 0.7)), step=0.05
        )
    with r2:
        date_tol = st.number_input(
            "日期(天)", value=int(rules.get("date_days_tolerance", 7)), step=1
        )
        vendor_thresh = st.number_input(
            "相似度",
            value=float(rules.get("vendor_similarity_threshold", 0.6)),
            min_value=0.0,
            max_value=1.0,
            step=0.05,
        )
    with r3:
        allow_dupe = st.checkbox(
            "允许多票",
            value=bool(rules.get("allow_duplicate_orders", False)),
        )
        merge_items = st.checkbox(
            "合并明细",
            value=bool(rules.get("merge_multi_item_orders", True)),
        )

    if st.button("保存配置"):
        config["order_columns"] = {
            "order_id": order_id.strip(),
            "amount": amount.strip(),
            "date": date_col.strip(),
            "vendor": vendor.strip(),
            "description": description.strip(),
            "status": status.strip(),
            "product_link": product_link.strip(),
        }
        config.setdefault("match_rules", {})
        config["match_rules"]["amount_tolerance"] = amount_tol
        config["match_rules"]["date_days_tolerance"] = int(date_tol)
        config["match_rules"]["min_score"] = min_score
        config["match_rules"]["allow_duplicate_orders"] = allow_dupe
        config["match_rules"]["vendor_similarity_threshold"] = vendor_thresh
        config["match_rules"]["merge_multi_item_orders"] = merge_items
        save_config(config_path, config)


def render_preprocess_tab(orders_paths, config_path):
    st.subheader("预处理")
    if st.session_state.preprocess_notice:
        st.success(st.session_state.preprocess_notice)
        st.session_state.preprocess_notice = ""

    if not AGGRID_AVAILABLE:
        st.error("未安装可编辑表格组件：streamlit-aggrid")
        st.caption("安装命令：pip install streamlit-aggrid")
        return

    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("解析"):
            try:
                config = matcher.load_config(config_path)
                if not orders_paths:
                    st.error("请选择订单文件")
                    return
                orders_df = matcher.load_orders_multi(orders_paths, config)
                editor_df = orders_df_for_editor(orders_df)
                set_order_df(
                    editor_df,
                    notice=f"已解析 {len(editor_df)} 条",
                    reset_grid=True,
                    reset_row_ids=True,
                )
            except Exception as exc:
                st.error(f"解析失败：{exc}")
    with col2:
        allow_edit = st.checkbox("允许编辑", value=False)

    if allow_edit:
        st.caption("双击单元格编辑，点击行后可删除当前行。")
    else:
        st.caption("当前为只读模式，如需修改请勾选“允许编辑”。")

    current_df = ensure_editor_df(st.session_state.order_df)
    st.session_state.order_df = current_df
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
    if status_options:
        selected_statuses = st.multiselect(
            "状态筛选",
            options=status_options,
            default=status_options,
            key="order_status_filter",
        )
    else:
        selected_statuses = []

    if status_options and selected_statuses:
        display_base_df = current_df[current_df["order_status"].isin(selected_statuses)]
    elif status_options and not selected_statuses:
        display_base_df = current_df.iloc[0:0]
    else:
        display_base_df = current_df

    selected_row_id = st.session_state.get("preprocess_selected_row_id")
    display_df = add_auto_unique_id_column(display_base_df)

    grid_key = f"order_grid_{st.session_state.order_grid_version}"
    grid_response = AgGrid(
        display_df,
        gridOptions=build_order_grid_options(
            display_df,
            allow_edit,
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
    if allow_edit and updated_data is not None:
        if isinstance(updated_data, list):
            updated_df = pd.DataFrame(updated_data)
        else:
            updated_df = updated_data
        updated_clean = ensure_editor_df(updated_df)
        if not updated_clean.equals(display_base_df):
            st.session_state.match_result_df = None
            st.session_state.match_detail = {}
        if display_base_df.shape[0] != current_df.shape[0]:
            st.session_state.order_df = merge_grid_updates(current_df, updated_clean)
        else:
            st.session_state.order_df = updated_clean

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

    action_col1, action_col2, action_col3 = st.columns([1, 1, 1])
    with action_col1:
        if st.button("新增空行"):
            new_row = {
                "order_id": "",
                "vendor": "",
                "description": "",
                "product_link": "",
                "amount": None,
                "order_status": "",
                "order_date": "",
                "_row_id": st.session_state.order_row_id_seq,
            }
            st.session_state.order_row_id_seq += 1
            new_df = pd.concat(
                [st.session_state.order_df, pd.DataFrame([new_row])],
                ignore_index=True,
            )
            set_order_df(new_df, notice="已新增空行", reset_grid=True)
            trigger_rerun()
    with action_col2:
        if st.button("删除选中"):
            if not selected_ids:
                st.warning("请先勾选要删除的订单")
            else:
                new_df = st.session_state.order_df[
                    ~st.session_state.order_df["_row_id"].isin(selected_ids)
                ]
                set_order_df(
                    new_df,
                    notice=f"已删除 {len(selected_ids)} 行",
                    reset_grid=True,
                )
                trigger_rerun()
    with action_col3:
        if st.button("清空"):
            empty_df = pd.DataFrame(columns=EDITOR_COLUMNS)
            set_order_df(empty_df, notice="已清空", reset_grid=True, reset_row_ids=True)
            trigger_rerun()
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
        )
    with export_col2:
        if st.button("导出 Excel(另存为)"):
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
                st.download_button("导出 Excel", data=b"", file_name="订单预筛选.xlsx", disabled=True)
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
                    )
                except Exception as exc:
                    st.warning(f"导出 Excel 失败：{exc}")
    with export_col3:
        st.caption(f"当前 {len(export_df)} 行")
        total_amount = sum_amounts(st.session_state.order_df.get("amount", []))
        st.caption(f"订单总金额：{format_amount(total_amount)}")

    with st.expander("新增"):
        new_cols = st.columns(3)
        with new_cols[0]:
            new_order_id = st.text_input("订单号", key="new_order_id")
            new_vendor = st.text_input("店铺", key="new_vendor")
        with new_cols[1]:
            new_desc = st.text_input("明细", key="new_desc")
            new_amount = st.text_input("实付", key="new_amount")
        with new_cols[2]:
            new_status = st.text_input("状态", key="new_status")
            new_date = st.text_input("时间", key="new_date")
            new_link = st.text_input("商品链接", key="new_link")

        if st.button("添加"):
            amount_value = matcher.parse_amount(new_amount) if new_amount else None
            new_row = {
                "order_id": matcher.format_order_id(new_order_id),
                "vendor": new_vendor.strip(),
                "description": new_desc.strip(),
                "product_link": new_link.strip(),
                "amount": amount_value if amount_value is not None else new_amount,
                "order_status": new_status.strip(),
                "order_date": new_date.strip(),
                "_row_id": st.session_state.order_row_id_seq,
            }
            st.session_state.order_row_id_seq += 1
            new_df = pd.concat(
                [st.session_state.order_df, pd.DataFrame([new_row])],
                ignore_index=True,
            )
            set_order_df(new_df, notice="已添加", reset_grid=True)
            trigger_rerun()


def render_match_tab(orders_paths, pdf_dir, config_path, output_path, recursive, color_map):
    st.subheader("匹配")
    orders_hint = f"{len(orders_paths)} 个文件" if orders_paths else "未选择"
    st.caption(f"订单：{orders_hint} ｜ 发票：{pdf_dir}")
    matched_orders_for_screenshots = pd.DataFrame()

    if st.button("开始"):
        config = load_config(config_path)
        config.setdefault("match_rules", {})
        config["match_rules"]["vendor_similarity_threshold"] = float(
            st.session_state.suspect_threshold
        )
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
    df = df.copy()
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
        if col not in df.columns:
            df[col] = default
    if "_row_id" not in df.columns:
        df["_row_id"] = range(1, len(df) + 1)
    df["order_date"] = df["order_date"].apply(format_editor_date)
    return df


def render_screenshot_tab():
    st.subheader("截图管理")
    if not AGGRID_AVAILABLE:
        st.error("未安装可编辑表格组件：streamlit-aggrid")
        st.caption("安装命令：pip install streamlit-aggrid")
        return
    uploaded = st.file_uploader("导入已匹配订单文件", type=["xlsx", "csv"])
    prefill_df = st.session_state.get("screenshot_prefill_df")
    if uploaded:
        try:
            df = load_matched_orders_file(uploaded)
            st.session_state.screenshot_prefill_df = None
        except Exception as exc:
            st.error(f"导入失败：{exc}")
            return
    elif prefill_df is not None:
        df = prefill_df.copy()
        if df.empty:
            st.warning("匹配结果为空，请先完成匹配或手动导入")
            return
        st.caption("已从匹配结果加载，可继续或上传文件替换。")
    else:
        st.caption("请上传已匹配订单文件（由匹配页导出）。")
        return

    df = prepare_screenshot_orders_df(df)
    st.caption(f"已导入 {len(df)} 行")
    st.caption("点击“截图状态”可进入截图管理。")
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
    st.set_page_config(page_title="发票匹配", layout="wide")
    ensure_session_state()
    start_auto_shutdown_monitor()

    st.title("发票匹配")

    with st.sidebar:
        st.header("文件")
        select_col1, select_col2 = st.columns([1, 3])
        with select_col1:
            if st.button("选订单"):
                chosen = select_order_files()
                if chosen:
                    st.session_state.orders_paths_text = "\n".join(chosen)
        with select_col2:
            st.text_area("订单文件(多选)", key="orders_paths_text", height=80)

        select_col3, select_col4 = st.columns([1, 3])
        with select_col3:
            if st.button("选发票"):
                chosen = select_pdf_folder()
                if chosen:
                    st.session_state.pdf_dir = chosen
        with select_col4:
            st.text_input("发票文件夹", key="pdf_dir")

        recursive = st.checkbox("递归扫描", value=False)

        with st.expander("设置", expanded=True):
            st.text_input("配置文件", key="config_path")
            st.text_input("输出文件", key="output_path")
            st.subheader("疑似阈值")
            st.slider(
                "相似度",
                min_value=0.0,
                max_value=1.0,
                step=0.05,
                key="suspect_threshold",
            )
            st.subheader("颜色")
            st.color_picker("完美", key="color_perfect")
            st.color_picker("疑似", key="color_suspect")
            st.color_picker("无匹配", key="color_nomatch")

    orders_paths = parse_orders_paths(st.session_state.orders_paths_text)
    pdf_dir = st.session_state.pdf_dir
    config_path = st.session_state.config_path
    output_path = st.session_state.output_path

    color_map = {
        "完美匹配": st.session_state.color_perfect,
        "疑似匹配": st.session_state.color_suspect,
        "无匹配": st.session_state.color_nomatch,
    }

    tab_labels = ["预处理", "匹配", "配置", "截图"]
    pending_tab = st.session_state.get("pending_tab")
    if pending_tab in tab_labels:
        st.session_state.active_tab = pending_tab
        st.session_state.pending_tab = None
    current_tab = st.session_state.get("active_tab", tab_labels[0])
    if current_tab not in tab_labels:
        current_tab = tab_labels[0]
    active_tab = st.radio(
        "导航",
        tab_labels,
        index=tab_labels.index(current_tab),
        horizontal=True,
        label_visibility="collapsed",
        key="active_tab",
    )

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
    elif active_tab == "配置":
        render_config_tab(config_path)
    else:
        render_screenshot_tab()

if __name__ == "__main__":
    main()
