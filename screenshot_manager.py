#!/usr/bin/env python3
import io

import streamlit as st
from PIL import Image, ImageGrab


CATEGORY_ORDER = "order"
CATEGORY_PAYMENT = "payment"
STATUS_NONE = "无截图"
STATUS_PARTIAL = "截图不完全"
STATUS_FULL = "有截图"


def init_session_state():
    if "screenshot_store" not in st.session_state:
        st.session_state.screenshot_store = {}


def build_order_key(order_id, fallback_id=None):
    order_id = str(order_id or "").strip()
    if order_id:
        return order_id
    if fallback_id is None:
        return "row-unknown"
    return f"row-{fallback_id}"


def ensure_entry(order_key):
    store = st.session_state.screenshot_store
    if order_key not in store:
        store[order_key] = {CATEGORY_ORDER: [], CATEGORY_PAYMENT: []}
    return store[order_key]


def read_clipboard_image():
    image = ImageGrab.grabclipboard()
    if image is None:
        return None
    if isinstance(image, list):
        for item in image:
            try:
                return Image.open(item)
            except Exception:
                continue
        return None
    return image


def image_to_bytes(image):
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


def add_clipboard_image(order_key, category):
    image = read_clipboard_image()
    if image is None:
        st.warning("剪贴板里没有图片。")
        return False
    entry = ensure_entry(order_key)
    entry[category].append(image_to_bytes(image))
    return True


def clear_images(order_key, category):
    entry = ensure_entry(order_key)
    entry[category] = []


def remove_image(order_key, category, index):
    entry = ensure_entry(order_key)
    if 0 <= index < len(entry[category]):
        del entry[category][index]


def get_status_label(order_key):
    entry = ensure_entry(order_key)
    has_order = len(entry[CATEGORY_ORDER]) > 0
    has_payment = len(entry[CATEGORY_PAYMENT]) > 0
    if has_order and has_payment:
        return STATUS_FULL
    if has_order or has_payment:
        return STATUS_PARTIAL
    return STATUS_NONE


def render_screenshot_manager(order_key, order_label=None, on_next_order=None, on_prev_order=None):
    entry = ensure_entry(order_key)
    if order_label:
        st.caption(order_label)

    render_category_block(order_key, entry, CATEGORY_ORDER, "订单截图")
    render_category_block(order_key, entry, CATEGORY_PAYMENT, "支付截图")
    if on_next_order or on_prev_order:
        st.divider()
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            if on_prev_order and st.button("上一个订单", key=f"{order_key}_prev"):
                on_prev_order()
        with col2:
            if on_next_order and st.button("下一个订单", key=f"{order_key}_next"):
                on_next_order()


def render_category_block(order_key, entry, category, title):
    st.markdown(f"**{title}**")
    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        if st.button("从剪贴板添加", key=f"{order_key}_{category}_paste"):
            if add_clipboard_image(order_key, category):
                trigger_rerun()
    with col2:
        if st.button("清空", key=f"{order_key}_{category}_clear"):
            clear_images(order_key, category)
            trigger_rerun()
    with col3:
        st.caption(f"已添加 {len(entry[category])} 张")

    if not entry[category]:
        st.caption("暂无截图")
        return

    remove_idx = None
    cols = st.columns(3)
    for idx, img_bytes in enumerate(entry[category]):
        with cols[idx % 3]:
            st.image(img_bytes, use_container_width=True)
            if st.button("删除", key=f"{order_key}_{category}_del_{idx}"):
                remove_idx = idx
    if remove_idx is not None:
        remove_image(order_key, category, remove_idx)
        trigger_rerun()


def trigger_rerun():
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()
