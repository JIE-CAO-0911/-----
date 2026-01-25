#!/usr/bin/env python3
import json
from pathlib import Path

import pandas as pd
import streamlit as st

import invoice_matcher as matcher

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
        st.success("配置已保存。")
    except Exception as exc:
        st.error(f"保存配置失败：{exc}")


def ensure_session_state():
    if "order_df" not in st.session_state:
        st.session_state.order_df = pd.DataFrame(
            columns=[
                "order_id",
                "vendor",
                "description",
                "amount",
                "order_status",
                "order_date",
            ]
        )
    if "match_result_df" not in st.session_state:
        st.session_state.match_result_df = None
    if "summary" not in st.session_state:
        st.session_state.summary = None
    if "orders_path" not in st.session_state:
        st.session_state.orders_path = default_orders_path()
    if "pdf_dir" not in st.session_state:
        st.session_state.pdf_dir = str(APP_DIR / "发票PDF")
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


def orders_df_for_editor(orders_df):
    return pd.DataFrame(
        {
            "order_id": orders_df["_order_id"].apply(matcher.format_order_id),
            "vendor": orders_df["_vendor"],
            "description": orders_df["_description"],
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

def select_order_file():
    if not TK_AVAILABLE:
        st.warning("当前环境无法打开文件选择窗口，请手动输入路径。")
        return None
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title="选择订单文件",
        filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
    )
    root.destroy()
    return path or None


def select_pdf_folder():
    if not TK_AVAILABLE:
        st.warning("当前环境无法打开文件夹选择窗口，请手动输入路径。")
        return None
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askdirectory(title="选择发票 PDF 文件夹")
    root.destroy()
    return path or None


def render_config_tab(config_path):
    config = load_config(config_path)

    st.subheader("订单列映射")
    order_cols = config.get("order_columns", {})
    col1, col2 = st.columns(2)
    with col1:
        order_id = st.text_input("订单号", value=order_cols.get("order_id", ""))
        amount = st.text_input("实付金额", value=order_cols.get("amount", ""))
        date_col = st.text_input("订单时间", value=order_cols.get("date", ""))
    with col2:
        vendor = st.text_input("店铺名称", value=order_cols.get("vendor", ""))
        description = st.text_input("商品名称", value=order_cols.get("description", ""))
        status = st.text_input("订单状态", value=order_cols.get("status", ""))

    st.subheader("匹配规则")
    rules = config.get("match_rules", {})
    r1, r2, r3 = st.columns(3)
    with r1:
        amount_tol = st.number_input(
            "金额容差", value=float(rules.get("amount_tolerance", 0.01)), step=0.01
        )
        min_score = st.number_input(
            "最低匹配分", value=float(rules.get("min_score", 0.7)), step=0.05
        )
    with r2:
        date_tol = st.number_input(
            "日期容差(天)", value=int(rules.get("date_days_tolerance", 7)), step=1
        )
        vendor_thresh = st.number_input(
            "店铺相似度阈值",
            value=float(rules.get("vendor_similarity_threshold", 0.6)),
            min_value=0.0,
            max_value=1.0,
            step=0.05,
        )
    with r3:
        allow_dupe = st.checkbox(
            "允许一个订单匹配多张发票",
            value=bool(rules.get("allow_duplicate_orders", False)),
        )
        merge_items = st.checkbox(
            "合并多行商品",
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
        }
        config.setdefault("match_rules", {})
        config["match_rules"]["amount_tolerance"] = amount_tol
        config["match_rules"]["date_days_tolerance"] = int(date_tol)
        config["match_rules"]["min_score"] = min_score
        config["match_rules"]["allow_duplicate_orders"] = allow_dupe
        config["match_rules"]["vendor_similarity_threshold"] = vendor_thresh
        config["match_rules"]["merge_multi_item_orders"] = merge_items
        save_config(config_path, config)


def render_preprocess_tab(orders_path, config_path):
    st.subheader("订单预处理")
    st.caption("先解析订单，编辑或删除后再进行匹配。")

    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("解析订单"):
            try:
                config = matcher.load_config(config_path)
                orders_df = matcher.load_orders(orders_path, config)
                st.session_state.order_df = orders_df_for_editor(orders_df)
                st.session_state.match_result_df = None
                st.success(f"已解析 {len(st.session_state.order_df)} 条订单。")
            except Exception as exc:
                st.error(f"解析失败：{exc}")

    editor_df = st.data_editor(
        st.session_state.order_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=False,
        key="order_editor",
    )

    action_col1, action_col2, action_col3 = st.columns([1, 1, 2])
    with action_col1:
        if st.button("应用编辑"):
            st.session_state.order_df = editor_df.reset_index(drop=True)
            st.session_state.match_result_df = None
            st.success("订单列表已更新。")
    with action_col2:
        delete_targets = st.multiselect(
            "删除行(按序号)",
            options=list(editor_df.index),
            format_func=lambda idx: f"{idx} | {editor_df.loc[idx, 'order_id']}",
        )
        if st.button("删除选中"):
            st.session_state.order_df = editor_df.drop(delete_targets).reset_index(drop=True)
            st.session_state.match_result_df = None
            st.success("已删除选中订单。")
    with action_col3:
        if st.button("清空列表"):
            st.session_state.order_df = st.session_state.order_df.iloc[0:0]
            st.session_state.match_result_df = None
            st.info("订单列表已清空。")

    with st.expander("新增订单"):
        new_cols = st.columns(3)
        with new_cols[0]:
            new_order_id = st.text_input("订单号(新增)", key="new_order_id")
            new_vendor = st.text_input("店铺名称(新增)", key="new_vendor")
        with new_cols[1]:
            new_desc = st.text_input("明细(新增)", key="new_desc")
            new_amount = st.text_input("实付金额(新增)", key="new_amount")
        with new_cols[2]:
            new_status = st.text_input("订单状态(新增)", key="new_status")
            new_date = st.text_input("订单时间(新增)", key="new_date")

        if st.button("添加订单"):
            amount_value = matcher.parse_amount(new_amount) if new_amount else None
            new_row = {
                "order_id": matcher.format_order_id(new_order_id),
                "vendor": new_vendor.strip(),
                "description": new_desc.strip(),
                "amount": amount_value if amount_value is not None else new_amount,
                "order_status": new_status.strip(),
                "order_date": new_date.strip(),
            }
            st.session_state.order_df = pd.concat(
                [st.session_state.order_df, pd.DataFrame([new_row])],
                ignore_index=True,
            )
            st.session_state.match_result_df = None
            st.success("已添加订单。")


def render_match_tab(orders_path, pdf_dir, config_path, output_path, recursive, color_map):
    st.subheader("发票匹配")
    st.caption(f"订单文件：{orders_path} ｜ 发票文件夹：{pdf_dir}")

    if st.button("开始匹配"):
        config = load_config(config_path)
        config.setdefault("match_rules", {})
        config["match_rules"]["vendor_similarity_threshold"] = float(
            st.session_state.suspect_threshold
        )
        try:
            backend = matcher.resolve_pdf_backend()
            if backend is None:
                st.error("缺少 PDF 解析库，请安装 pdfplumber 或 pypdf。")
                return

            if st.session_state.order_df is not None and not st.session_state.order_df.empty:
                order_rows = editor_df_to_rows(st.session_state.order_df)
                orders_df = matcher.build_orders_df_from_rows(order_rows)
            else:
                orders_df = matcher.load_orders(orders_path, config)

            pdf_root = Path(pdf_dir)
            if not pdf_root.exists():
                st.error("发票文件夹路径无效。")
                return

            pdfs = matcher.find_pdfs(pdf_dir, recursive)
            if not pdfs:
                st.error("未找到任何 PDF 文件。")
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
                    "order_id": orders_df["_order_id"].apply(matcher.format_order_id),
                    "vendor": orders_df["_vendor"],
                    "description": orders_df["_description"],
                    "amount": orders_df["_amount"],
                    "order_status": orders_df["_status"],
                    "match_status": orders_df["_match_status"],
                    "match_reason": orders_df["_match_reason"],
                }
            )

            st.session_state.match_result_df = result_df
            st.session_state.summary = {
                "orders": len(orders_df),
                "invoices": len(invoices),
                "matches": len(matches),
                "output": output_path,
            }
            st.success("匹配完成。")
        except Exception as exc:
            st.error(f"匹配失败：{exc}")

    if st.session_state.summary:
        summary = st.session_state.summary
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("订单数", summary["orders"])
        m2.metric("发票数", summary["invoices"])
        m3.metric("匹配数", summary["matches"])
        m4.metric("输出文件", Path(summary["output"]).name)

    if st.session_state.match_result_df is not None:
        st.markdown("### 匹配结果")
        status_filter = st.selectbox(
            "订单状态筛选",
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

        st.dataframe(
            view_df.style.apply(lambda r: style_match_rows(r, color_map), axis=1),
            use_container_width=True,
            height=520,
        )


def main():
    st.set_page_config(page_title="发票/订单匹配助手", layout="wide")
    ensure_session_state()

    st.title("发票/订单匹配助手（Streamlit 版）")

    with st.sidebar:
        st.header("文件选择")
        select_col1, select_col2 = st.columns([1, 3])
        with select_col1:
            if st.button("选择订单文件"):
                chosen = select_order_file()
                if chosen:
                    st.session_state.orders_path = chosen
        with select_col2:
            st.text_input("订单文件 (.xlsx)", key="orders_path")

        select_col3, select_col4 = st.columns([1, 3])
        with select_col3:
            if st.button("选择发票文件夹"):
                chosen = select_pdf_folder()
                if chosen:
                    st.session_state.pdf_dir = chosen
        with select_col4:
            st.text_input("发票 PDF 文件夹", key="pdf_dir")

        recursive = st.checkbox("递归扫描子文件夹", value=False)

        with st.expander("设置面板", expanded=True):
            st.text_input("配置文件 (.json)", key="config_path")
            st.text_input("输出结果", key="output_path")
            st.subheader("疑似匹配阈值")
            st.slider(
                "店铺相似度阈值",
                min_value=0.0,
                max_value=1.0,
                step=0.05,
                key="suspect_threshold",
            )
            st.subheader("颜色规则")
            st.color_picker("完美匹配", key="color_perfect")
            st.color_picker("疑似匹配", key="color_suspect")
            st.color_picker("无匹配", key="color_nomatch")

    orders_path = st.session_state.orders_path
    pdf_dir = st.session_state.pdf_dir
    config_path = st.session_state.config_path
    output_path = st.session_state.output_path

    tab1, tab2, tab3 = st.tabs(["订单预处理", "匹配", "配置"])
    color_map = {
        "完美匹配": st.session_state.color_perfect,
        "疑似匹配": st.session_state.color_suspect,
        "无匹配": st.session_state.color_nomatch,
    }

    with tab1:
        render_preprocess_tab(orders_path, config_path)
    with tab2:
        render_match_tab(
            orders_path,
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
