import io
import os
import subprocess
import sys
import tempfile
import zipfile

from docx import Document
import pandas as pd
import streamlit as st

# ================= 1. НАЛАШТУВАННЯ СТОРІНКИ =================
st.set_page_config(
    page_title="Docx to PDF Generator",
    page_icon="📄",
    layout="centered"
)

st.title("📄 Генератор документів у PDF з архівацією в ZIP")
st.caption("Завантажте Word-шаблон і Excel-таблицю, щоб отримати готовий архів із PDF-файлами.")


# ================= 2. ФУНКЦІЇ ОБРОБКИ ТЕКСТУ В DOCX =================
def clean_and_replace_run_text(paragraphs, company, materials):
    """Склеює розірвані плейсхолдери та підставляє значення зі збереженням форматування."""
    for p in paragraphs:
        # 1. Склеюємо розірвані Вордом шматки плейсхолдерів
        i = 0
        while i < len(p.runs) - 1:
            if any(part in p.runs[i].text for part in ["Company", "Material", "_name"]):
                p.runs[i].text += p.runs[i + 1].text
                p._p.remove(p.runs[i + 1]._r)
                continue
            i += 1

        # 2. Заміна тексту
        for run in p.runs:
            if "Company_name" in run.text:
                run.text = run.text.replace("Company_name", company)
            if "Material_name" in run.text:
                run.text = run.text.replace("Material_name", materials)


def replace_placeholders(doc_path, company, materials, output_docx):
    """Замінює плейсхолдери в основному тексті та в усіх таблицях документа."""
    doc = Document(doc_path)

    # Заміна в основному тексті
    clean_and_replace_run_text(doc.paragraphs, company, materials)

    # Заміна в таблицях
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                clean_and_replace_run_text(cell.paragraphs, company, materials)

    doc.save(output_docx)


def convert_docx_to_pdf(docx_path, output_pdf_path):
    """Універсальна конвертація: через docx2pdf (на Windows) або LibreOffice (на Linux / Streamlit Cloud)."""
    if sys.platform == "win32":
        from docx2pdf import convert
        convert(docx_path, output_pdf_path)
    else:
        output_dir = os.path.dirname(output_pdf_path) or "."
        subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                docx_path,
                "--outdir",
                output_dir,
            ],
            check=True,
        )
        generated_pdf = os.path.splitext(docx_path)[0] + ".pdf"
        if generated_pdf != output_pdf_path and os.path.exists(generated_pdf):
            if os.path.exists(output_pdf_path):
                os.remove(output_pdf_path)
            os.rename(generated_pdf, output_pdf_path)


# ================= 3. ЗАВАНТАЖЕННЯ ФАЙЛІВ =================
st.subheader("1. Завантаження шаблону та таблиці")
col1, col2 = st.columns(2)

with col1:
    uploaded_template = st.file_uploader(
        "Шаблон Word (.docx)", type=["docx"], key="docx_uploader"
    )

with col2:
    uploaded_excel = st.file_uploader(
        "Таблиця даних (.xlsx)", type=["xlsx", "xls", "csv"], key="excel_uploader"
    )

# ================= 4. ОБРОБКА ДАНИХ ТА ГЕНЕРАЦІЯ =================
if uploaded_template and uploaded_excel:
    try:
        if uploaded_excel.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_excel)
        else:
            df = pd.read_excel(uploaded_excel)

        df.columns = df.columns.astype(str).str.strip()

        # Автопошук колонок
        company_col = next(
            (c for c in df.columns if any(k in c.lower() for k in ["company", "назва", "організац", "компан"])),
            None,
        )
        material_col = next(
            (c for c in df.columns if any(k in c.lower() for k in ["material", "матеріал", "продукц"])),
            None,
        )

        if not company_col:
            st.error(f"❌ У таблиці не знайдено колонки з компанією. Знайдені колонки: {list(df.columns)}")
        else:
            df_clean = df.dropna(subset=[company_col]).copy()
            st.write(f"📊 Знайдено рядків для формування документів: **{len(df_clean)}**")
            st.dataframe(df_clean.head(5), use_container_width=True)

            # ================= 5. ЗАПУСК ГЕНЕРАЦІЇ =================
            st.markdown("---")
            if st.button("🚀 Згенерувати PDF та створити ZIP-архів", type="primary"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                log_box = st.container()

                with tempfile.TemporaryDirectory() as temp_dir:
                    template_path = os.path.join(temp_dir, "Template.docx")
                    with open(template_path, "wb") as f:
                        f.write(uploaded_template.getbuffer())

                    generated_pdf_paths = []
                    total = len(df_clean)

                    for index, (_, row) in enumerate(df_clean.iterrows()):
                        company = str(row[company_col]).strip()
                        materials = (
                            str(row[material_col]).strip()
                            if material_col and pd.notna(row[material_col])
                            else ""
                        )

                        safe_company = "".join(c for c in company if c.isalnum() or c in (" ", "_", "-")).rstrip()
                        temp_docx = os.path.join(temp_dir, f"temp_{safe_company}.docx")
                        output_pdf = os.path.join(temp_dir, f"Пропозиція партнерства {safe_company}.pdf")

                        status_text.text(f"Обробка [{index + 1}/{total}]: {company}...")

                        # 1. Заміна плейсхолдерів у DOCX
                        replace_placeholders(template_path, company, materials, temp_docx)

                        # 2. Конвертація у PDF
                        try:
                            convert_docx_to_pdf(temp_docx, output_pdf)
                            if os.path.exists(output_pdf):
                                generated_pdf_paths.append((output_pdf, os.path.basename(output_pdf)))
                                log_box.success(f"✅ Згенеровано: {os.path.basename(output_pdf)}")
                        except Exception as e:
                            log_box.error(f"❌ Помилка для {company}: {e}")
                        finally:
                            if os.path.exists(temp_docx):
                                os.remove(temp_docx)

                        progress_bar.progress((index + 1) / total)

                    # ================= 6. ПАКУВАННЯ В ZIP-АРХІВ =================
                    if generated_pdf_paths:
                        status_text.text("Пакування файлів у ZIP-архів...")
                        zip_buffer = io.BytesIO()

                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for pdf_file_path, arcname in generated_pdf_paths:
                                zip_file.write(pdf_file_path, arcname=arcname)

                        zip_buffer.seek(0)
                        status_text.text("Готово!")
                        st.balloons()
                        st.success(f"🎉 Успішно створено {len(generated_pdf_paths)} PDF-файлів!")

                        st.download_button(
                            label="📥 Завантажити всі PDF в одному ZIP-архіві",
                            data=zip_buffer,
                            file_name="Пропозиції_партнерства_PDF.zip",
                            mime="application/zip",
                            type="primary",
                        )
                    else:
                        st.warning("⚠️ Жодного PDF-файлу не вдалося згенерувати.")

    except Exception as e:
        st.error(f"Помилка зчитування файлів: {e}")
else:
    st.info("👆 Будь ласка, завантажте Word-шаблон та таблицю для початку.")