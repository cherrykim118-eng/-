import os
import tempfile

import streamlit as st
from docx import Document
from docx2pdf import convert


def docx_to_pdf_simple(docx_path: str, pdf_path: str):
    # docx → pdf 변환 (Windows/Mac에서만 정상 동작)
    convert(docx_path, pdf_path)


def main():
    st.set_page_config(page_title="DOCX → PDF 변환기", page_icon="📝")
    st.title("📝 DOCX를 PDF로 변환하기")
    st.write("워드 파일(DOCX)을 PDF 파일로 변환합니다.")

    uploaded_file = st.file_uploader("DOCX 파일을 업로드하세요", type=["docx"])

    if st.button("변환 시작"):
        if uploaded_file is None:
            st.warning("먼저 DOCX 파일을 업로드해주세요.")
            return

        with st.spinner("DOCX를 처리하는 중입니다..."):
            # 업로드된 DOCX → 임시 파일 저장
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                tmp_docx.write(uploaded_file.read())
                docx_path = tmp_docx.name

            base_name = os.path.splitext(os.path.basename(uploaded_file.name))[0]

            # 변환 후 결과 PDF 파일 저장 경로
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                pdf_path = tmp_pdf.name

            try:
                # DOCX → PDF 변환 실행
                docx_to_pdf_simple(docx_path, pdf_path)

                # 변환된 PDF 읽어오기
                with open(pdf_path, "rb") as f:
                    pdf_data = f.read()

                st.success("변환이 완료되었습니다!")
                st.download_button(
                    label="PDF 파일 다운로드",
                    data=pdf_data,
                    file_name=f"{base_name}.pdf",
                    mime="application/pdf",
                )

            except Exception as e:
                st.error(f"변환 중 오류가 발생했습니다: {e}")

            finally:
                # 임시 파일 삭제
                try:
                    os.remove(docx_path)
                except Exception:
                    pass
                try:
                    os.remove(pdf_path)
                except Exception:
                    pass


if __name__ == "__main__":
    main()
