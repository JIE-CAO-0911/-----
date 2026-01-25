#!/usr/bin/env python3
import io
import zipfile

import streamlit as st
from PIL import ImageGrab


def main():
    st.set_page_config(page_title="Clipboard Paste Test", layout="centered")
    st.title("剪贴板图片采集")
    st.write("先复制图片到剪贴板，再点对应按钮。")

    if "clipboard_images" not in st.session_state:
        st.session_state.clipboard_images = [None] * 10

    def read_clipboard(slot_index):
        try:
            image = ImageGrab.grabclipboard()
            if image is None:
                st.warning("剪贴板里没有图片。")
                return
            st.session_state.clipboard_images[slot_index] = image
        except Exception as exc:
            st.error(f"读取失败：{exc}")

    export_col1, export_col2 = st.columns([1, 2])
    with export_col1:
        export_name = st.text_input("导出文件名", value="clipboard_images.zip")
    with export_col2:
        if st.button("生成导出文件"):
            buffer = io.BytesIO()
            with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for idx, image in enumerate(st.session_state.clipboard_images, 1):
                    if image is None:
                        continue
                    img_buf = io.BytesIO()
                    image.save(img_buf, format="PNG")
                    zf.writestr(f"clipboard_{idx:02d}.png", img_buf.getvalue())
            st.session_state.export_zip = buffer.getvalue()

    if st.session_state.get("export_zip"):
        st.download_button(
            label="下载导出文件",
            data=st.session_state.export_zip,
            file_name=export_name or "clipboard_images.zip",
            mime="application/zip",
        )

    cols = st.columns(2)
    for idx in range(10):
        with cols[idx % 2]:
            st.subheader(f"图片 {idx + 1}")
            if st.button(f"读取到图片 {idx + 1}", key=f"read_{idx}"):
                read_clipboard(idx)
            if st.button(f"清空图片 {idx + 1}", key=f"clear_{idx}"):
                st.session_state.clipboard_images[idx] = None
            image = st.session_state.clipboard_images[idx]
            if image is not None:
                st.image(image, use_container_width=True)

    st.subheader("Fallback: File Uploader")
    uploaded = st.file_uploader("Upload image", type=["png", "jpg", "jpeg"])
    if uploaded:
        st.image(uploaded, caption="Uploaded image", use_container_width=True)

    st.subheader("Camera Input (Alternative)")
    camera = st.camera_input("Take a photo")
    if camera:
        st.image(camera, caption="Camera image", use_container_width=True)


if __name__ == "__main__":
    main()
