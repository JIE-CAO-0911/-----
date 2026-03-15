#!/usr/bin/env python3
import base64
import io

import streamlit as st
from PIL import Image, ImageGrab


CATEGORY_ORDER = "order"
CATEGORY_PAYMENT = "payment"
B64_CHARS = set(
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
)
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


def _iter_legacy_container(items):
    if isinstance(items, dict):

        def sort_key(item):
            key = item[0]
            text = str(key).strip()
            if text.isdigit():
                return (0, int(text))
            return (1, text)

        for _, value in sorted(items.items(), key=sort_key):
            yield value
        return
    if isinstance(items, (list, tuple)):
        for value in items:
            yield value
        return
    return


def _looks_like_base64(text):
    stripped = text.strip()
    if len(stripped) < 24:
        return False
    return all(ch in B64_CHARS for ch in stripped)


def _is_valid_image_bytes(data):
    try:
        with Image.open(io.BytesIO(data)) as img:
            img.verify()
        return True
    except Exception:
        return False


def _coerce_image_bytes(value):
    if isinstance(value, (bytes, bytearray)):
        data = bytes(value)
    elif isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        if text.startswith("data:image") and "," in text:
            text = text.split(",", 1)[1].strip()
        if not _looks_like_base64(text):
            return None
        try:
            data = base64.b64decode(text.encode("ascii"), validate=False)
        except Exception:
            return None
    else:
        return None

    if not data:
        return None
    if not _is_valid_image_bytes(data):
        return None
    return data


def _normalize_image_list(items):
    normalized = []
    for value in _iter_legacy_container(items):
        data = _coerce_image_bytes(value)
        if data is not None:
            normalized.append(data)
    return normalized


def ensure_entry(order_key):
    store = st.session_state.screenshot_store
    entry = store.get(order_key)
    if not isinstance(entry, dict):
        entry = {}
    entry[CATEGORY_ORDER] = _normalize_image_list(entry.get(CATEGORY_ORDER, []))
    entry[CATEGORY_PAYMENT] = _normalize_image_list(entry.get(CATEGORY_PAYMENT, []))
    store[order_key] = entry
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
    entry[category] = _normalize_image_list(entry.get(category, []))
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
