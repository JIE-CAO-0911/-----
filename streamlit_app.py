#!/usr/bin/env python3
import io
import json
import os
import re
import subprocess
import sys
import threading
import time
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

import invoice_matcher as matcher

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


def build_order_grid_options(df, editable):
    builder = GridOptionsBuilder.from_dataframe(df)
    builder.configure_default_column(
        editable=editable,
        resizable=True,
        sortable=True,
        filter=True,
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
    builder.configure_selection("multiple", use_checkbox=True)
    builder.configure_grid_options(
        rowHeight=32,
        stopEditingWhenCellsLoseFocus=True,
    )
    link_click = build_product_link_click_handler()
    if link_click:
        builder.configure_grid_options(onCellClicked=link_click)
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


def build_match_grid_options(df, color_map):
    builder = GridOptionsBuilder.from_dataframe(df)
    builder.configure_default_column(
        editable=False,
        resizable=True,
        sortable=True,
        filter=True,
    )
    builder.configure_column("_order_row", header_name="ROW", hide=True)
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
        st.caption("双击单元格编辑，勾选行后点击“删除选中”。")
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
        display_df = current_df[current_df["order_status"].isin(selected_statuses)]
    elif status_options and not selected_statuses:
        display_df = current_df.iloc[0:0]
    else:
        display_df = current_df

    grid_key = f"order_grid_{st.session_state.order_grid_version}"
    grid_response = AgGrid(
        display_df,
        gridOptions=build_order_grid_options(display_df, allow_edit),
        data_return_mode=DataReturnMode.AS_INPUT,
        update_mode=GridUpdateMode.MODEL_CHANGED,
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
        updated_df = ensure_editor_df(updated_df)
        if not updated_df.equals(display_df):
            st.session_state.match_result_df = None
            st.session_state.match_detail = {}
        if display_df.shape[0] != current_df.shape[0]:
            st.session_state.order_df = merge_grid_updates(current_df, updated_df)
        else:
            st.session_state.order_df = updated_df

    selected_rows = grid_response.get("selected_rows")
    if selected_rows is None:
        selected_rows = []
    elif isinstance(selected_rows, pd.DataFrame):
        selected_rows = selected_rows.to_dict("records")
    selected_ids = {
        int(row["_row_id"])
        for row in selected_rows
        if row and row.get("_row_id") is not None
    }

    action_col1, action_col2, action_col3, action_col4 = st.columns([1, 1, 1, 2])
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
                st.warning("未选择行")
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
    with action_col4:
        st.caption(f"已选 {len(selected_ids)} 行")

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

        if AGGRID_AVAILABLE:
            grid_response = AgGrid(
                view_df,
                gridOptions=build_match_grid_options(view_df, color_map),
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

            if selected_rows:
                selected = selected_rows[0]
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

    tab1, tab2, tab3 = st.tabs(["预处理", "匹配", "配置"])
    color_map = {
        "完美匹配": st.session_state.color_perfect,
        "疑似匹配": st.session_state.color_suspect,
        "无匹配": st.session_state.color_nomatch,
    }

    with tab1:
        render_preprocess_tab(orders_paths, config_path)
    with tab2:
        render_match_tab(
            orders_paths,
            pdf_dir,
            config_path,
            output_path,
            recursive,
            color_map,
        )
    with tab3:
        render_config_tab(config_path)


if __name__ == "__main__":
    main()
