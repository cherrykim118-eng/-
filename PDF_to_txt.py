import os
import tempfile

import streamlit as st
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4


def docx_to_pdf_text_only(docx_path: str, pdf_path: str):
    # DOCX 파일 읽기
    document = Document(docx_path)

    # PDF 캔버스 생성
    c = canvas.Canvas(pdf_path, pagesize=A4)
    width, height = A4

    x = 50
    y = height - 50
    line_spacing = 14

    # DOCX 문단을 한 줄씩 PDF에 그리기
    for para in document.paragraphs:
        text = para.text

        # 줄바꿈 처리
        for line in text.split("\n"):
            if y < 50:  # 페이지 끝나면 새 페이지
                c.showPage()
                y = height - 50

            c.drawString(x, y, line)
            y -= line_spacing

    c.save()


def main():
    st.set_page_config(page_title="DOCX → PDF (텍스트만)", page_icon="📝")
    st.title("📝 DOCX → PDF 변환기 (텍스트만)")

    uploaded_file = st.file_uploader("DOCX 파일을 업로드하세요", type=["docx"])

    if st.button("변환 시작"):
        if uploaded_file is None:
            st.warning("먼저 DOCX 파일을 업로드해주세요.")
            return

        with st.spinner("DOCX 처리 중..."):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                tmp_docx.write(uploaded_file.read())
                docx_path = tmp_docx.name

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                pdf_path = tmp_pdf.name

            base_name = os.path.splitext(uploaded_file.name)[0]

            try:
                docx_to_pdf_text_only(docx_path, pdf_path)

                with open(pdf_path, "rb") as f:
                    pdf_data = f.read()

                st.success("변환 완료!")
                st.download_button(
                    label="PDF 다운로드",
                    data=pdf_data,
                    file_name=f"{base_name}.pdf",
                    mime="application/pdf",
                )

            except Exception as e:
                st.error(f"오류 발생: {e}")

            finally:
                try:
                    os.remove(docx_path)
                except:
                    pass
                try:
                    os.remove(pdf_path)
                except:
                    pass


if __name__ == "__main__":
    main()
