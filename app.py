import io
import os
import re
from datetime import date
from pathlib import Path

import streamlit as st
from docx import Document
from docx.shared import Cm, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openai import OpenAI
from pypdf import PdfReader
from openpyxl import load_workbook
from PIL import Image

APP_TITLE = "Генератор пояснительных записок"
st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="wide")

st.markdown("""
<style>
.main .block-container {max-width: 1120px; padding-top: 1.8rem; padding-bottom: 3rem;}
.hero {padding: 1.4rem 1.6rem; border: 1px solid #d8dee8; border-radius: 16px;
background: linear-gradient(120deg,#f7f9fc,#edf3fa); margin-bottom: 1.2rem;}
.hero h1 {margin:0 0 .35rem 0; font-size:2rem; color:#172235;}
.hero p {margin:0;color:#526174;}
div.stButton > button[kind="primary"] {width:100%; min-height:3rem; font-weight:650;}
div[data-testid="stFileUploader"] {border:1px dashed #aeb9c8;border-radius:12px;padding:.35rem;}
</style>
""", unsafe_allow_html=True)

st.markdown(f'<div class="hero"><h1>📄 {APP_TITLE}</h1><p>Создание пояснительной записки по BPMN-схеме и проектной документации</p></div>', unsafe_allow_html=True)

def extract_docx(data):
    d = Document(io.BytesIO(data))
    out = [p.text for p in d.paragraphs if p.text.strip()]
    for table in d.tables:
        for row in table.rows:
            out.append(" | ".join(c.text.replace("\n", " / ") for c in row.cells))
    return "\n".join(out)

def extract_pdf(data):
    reader = PdfReader(io.BytesIO(data))
    return "\n".join((p.extract_text() or "") for p in reader.pages)

def extract_xlsx(data):
    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    out = []
    for ws in wb.worksheets:
        out.append(f"Лист: {ws.title}")
        for row in ws.iter_rows(values_only=True):
            vals = [str(v) for v in row if v is not None]
            if vals:
                out.append(" | ".join(vals))
    return "\n".join(out)

def extract_text_file(upload):
    name = upload.name.lower()
    data = upload.getvalue()
    if name.endswith(".docx"):
        return extract_docx(data)
    if name.endswith(".pdf"):
        return extract_pdf(data)
    if name.endswith(".xlsx"):
        return extract_xlsx(data)
    if name.endswith((".txt", ".md", ".csv")):
        return data.decode("utf-8", errors="replace")
    if name.endswith(".bpmn") or name.endswith(".xml"):
        return data.decode("utf-8", errors="replace")
    return ""

def set_run_font(run, size=13, bold=None):
    run.font.name = "Times New Roman"
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Times New Roman")
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold

def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = paragraph.add_run()
    begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = " PAGE "
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
    r._r.append(begin); r._r.append(instr); r._r.append(end)

def add_title_page(doc, title, system, topic, executor, coexecutor, doc_date):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.first_line_indent = Cm(0)
    set_run_font(p.add_run("Приложение № …"), 13)
    for _ in range(4):
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(0)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    set_run_font(p.add_run("ПОЯСНИТЕЛЬНАЯ ЗАПИСКА"), 14, True)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    set_run_font(p.add_run(f"{title}\n{system}"), 13)
    for _ in range(8):
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(0)
    for label, value in [
        ("Шифр темы:", topic or "[указать]"),
        ("Исполнитель:", executor or "[указать]"),
        ("Соисполнители (при наличии):", coexecutor or "[указать]"),
        ("Дата:", doc_date or date.today().strftime("%d.%m.%Y"))
    ]:
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(0)
        set_run_font(p.add_run(f"{label} {value}"), 13)
    doc.add_page_break()

def render_markdown_to_docx(doc, text):
    lines = text.replace("\r\n", "\n").split("\n")
    for raw in lines:
        line = raw.strip()
        if not line:
            continue
        if line.startswith("# "):
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
            set_run_font(p.add_run(line[2:].strip()), 14, True)
        elif line.startswith("## "):
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
            set_run_font(p.add_run(line[3:].strip()), 14, True)
        elif line.startswith("- ") or line.startswith("• "):
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.left_indent = Cm(0.7)
            p.paragraph_format.first_line_indent = Cm(0)
            set_run_font(p.add_run(line[2:].strip()), 13)
        else:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(1.25)
            set_run_font(p.add_run(line), 13)

def make_docx(markdown, fields, diagram_upload):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2); sec.bottom_margin = Cm(2)
    sec.left_margin = Cm(2); sec.right_margin = Cm(1)
    sec.different_first_page_header_footer = True

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    normal.font.size = Pt(13)
    normal.paragraph_format.line_spacing = 1.3
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)

    add_title_page(doc, fields["component"], fields["system"], fields["topic"],
                   fields["executor"], fields["coexecutor"], fields["date"])
    render_markdown_to_docx(doc, markdown)

    if diagram_upload is not None:
        doc.add_page_break()
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        set_run_font(p.add_run("Приложение 1. BPMN-схема"), 14, True)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        data = diagram_upload.getvalue()
        suffix = Path(diagram_upload.name).suffix.lower()
        if suffix in [".png", ".jpg", ".jpeg"]:
            p.add_run().add_picture(io.BytesIO(data), width=Cm(17))
        elif suffix == ".pdf":
            # Diagram PDF is appended as a PDF source attachment in a separate file; DOCX receives a clear placeholder.
            p.add_run("Исходная BPMN-схема приложена отдельным PDF-файлом к результату генерации.")
    for section in doc.sections:
        section.different_first_page_header_footer = True
        add_page_number(section.footer.paragraphs[0])
        section.first_page_footer.paragraphs[0].text = ""
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

def build_prompt(system, component, purpose, reference_text, ftt_text, extras_text):
    return f"""
Ты — системный аналитик. Подготовь пояснительную записку к BPMN-схеме.

СВЕДЕНИЯ:
Информационная система: {system}
Компонент/процесс: {component}
Назначение: {purpose}

ОБЯЗАТЕЛЬНЫЕ ТРЕБОВАНИЯ К СОДЕРЖАНИЮ:
- Раскрой назначение функциональности системы.
- Опиши назначение и функции компонентов, модулей и подсистем.
- Отрази входящие, обрабатываемые, формируемые, сохраняемые и передаваемые данные.
- Опиши взаимодействия компонентов и внешних информационных систем.
- Последовательно опиши функциональность по логике BPMN-схемы.
- Сведения ФТТ включай в описание соответствующих функций, не создавая отдельный большой каталог требований.
- Описывай систему, её функции и информационные потоки, а не инструкции сотрудникам.
- Сохраняй терминологию источников. Не придумывай отсутствующие сведения.
- Если деталей недостаточно, коротко укажи это в соответствующем месте.
- Не добавляй общие рассуждения.

СТРУКТУРА:
Строго следуй структуре, порядку разделов и уровню детализации образца ПЗ. Если образец не предоставлен, используй структуру:
1. Общие положения
2. Цель и назначение схемы
3. Состав и границы охвата схемы
4. Компоненты системы и источники данных
5. Уровни детализации, формат и условные обозначения
6. Описание функциональности (раздели на подразделы по направлениям и этапам схемы)
7. Использование схемы и результаты
8. Заключение

ОБЪЁМ:
Пиши развёрнуто и содержательно. Каждый функциональный этап и каждое направление раскрой отдельным подразделом: назначение, входная информация, выполняемые системой функции, связи с компонентами, выходные данные и результаты. Не ограничивайся пересказом названий блоков. Объём должен соответствовать сложности схемы и быть сопоставимым с образцом, без повторов и искусственного увеличения текста.

ВХОДНЫЕ МАТЕРИАЛЫ:
--- ОБРАЗЕЦ ПЗ ---
{reference_text[:24000]}
--- ФТТ И ДОПОЛНИТЕЛЬНЫЕ МАТЕРИАЛЫ ---
{ftt_text[:36000]}
{extras_text[:16000]}

ФОРМАТ ВЫВОДА:
Верни только готовый текст ПЗ в Markdown. Не пиши вступление, комментарии или список выполненных действий.
"""

st.subheader("1. Сведения о документе")
c1, c2 = st.columns(2)
with c1:
    system_name = st.text_input("Информационная система *", placeholder="Например, ГИС ЦАП")
    component = st.text_input("Компонент / процесс *", placeholder="Например, Мониторинг")
    topic = st.text_input("Шифр темы", placeholder="При наличии")
with c2:
    executor = st.text_input("Исполнитель", placeholder="ФИО")
    coexecutor = st.text_input("Соисполнители", placeholder="При наличии")
    doc_date = st.text_input("Дата", value=date.today().strftime("%d.%m.%Y"))
purpose = st.text_area("Назначение системы или компонента *", placeholder="Укажите назначение и основные задачи...", height=90)

st.subheader("2. Загрузка материалов")
c1, c2 = st.columns(2)
with c1:
    diagram = st.file_uploader("BPMN-схема *", type=["png","jpg","jpeg","pdf","bpmn","xml"], help="Изображение схемы будет вставлено в конец DOCX. PDF/BPMN/XML будут учтены как исходный материал.")
    ftt_files = st.file_uploader("ФТТ и дополнительные документы", type=["docx","pdf","xlsx","txt","md","csv","bpmn","xml"], accept_multiple_files=True)
with c2:
    reference = st.file_uploader("Образец пояснительной записки", type=["docx","pdf"])
    title_template = st.file_uploader("Образец титульного листа", type=["docx","pdf"])
    extra_files = st.file_uploader("Дополнительные материалы", type=["docx","pdf","xlsx","txt","md","csv"], accept_multiple_files=True)

st.subheader("3. Настройки")
model = st.selectbox("Модель", ["gpt-4.1", "gpt-4.1-mini"], index=0)
api_key = st.text_input("OpenAI API key", type="password", help="Ключ используется только для текущего запуска и не сохраняется приложением.")
st.caption("Формат DOCX: Times New Roman 13–14 пт, интервал 1,3, выравнивание по ширине, первая строка 1,25 см, поля слева 2 см и справа 1 см, номер страницы снизу по центру без номера на титульном листе.")

def collect_files(files):
    chunks=[]
    for f in files or []:
        try:
            t=extract_text_file(f)
            if t.strip():
                chunks.append(f"\n### Файл: {f.name}\n{t}")
        except Exception as e:
            chunks.append(f"\n### Файл: {f.name}\nНе удалось извлечь текст: {e}")
    return "\n".join(chunks)

def image_payload(upload):
    import mimetypes
    mt = mimetypes.guess_type(upload.name)[0] or "image/png"
    b64 = __import__("base64").b64encode(upload.getvalue()).decode("ascii")
    return f"data:{mt};base64,{b64}"

if st.button("Сформировать пояснительную записку", type="primary"):
    errors=[]
    if not system_name.strip(): errors.append("Укажите информационную систему.")
    if not component.strip(): errors.append("Укажите компонент или процесс.")
    if not purpose.strip(): errors.append("Укажите назначение.")
    if diagram is None: errors.append("Загрузите BPMN-схему.")
    if not api_key.strip(): errors.append("Введите OpenAI API key.")
    if errors:
        for e in errors: st.error(e)
        st.stop()

    with st.status("Подготовка материалов и формирование ПЗ...", expanded=True) as status:
        try:
            st.write("Извлекаю текст из приложенных документов…")
            ftt_text = collect_files(ftt_files)
            extras_text = collect_files(extra_files)
            ref_text = ""
            if reference:
                ref_text = extract_docx(reference.getvalue()) if reference.name.lower().endswith(".docx") else extract_pdf(reference.getvalue())

            diagram_text = ""
            diagram_image = None
            if diagram.name.lower().endswith((".bpmn",".xml")):
                diagram_text = diagram.getvalue().decode("utf-8", errors="replace")
            elif diagram.name.lower().endswith(".pdf"):
                diagram_text = extract_pdf(diagram.getvalue())
            else:
                diagram_image = image_payload(diagram)

            if title_template:
                try:
                    extras_text += "\n\n### Образец титульного листа:\n" + (
                        extract_docx(title_template.getvalue()) if title_template.name.lower().endswith(".docx") else extract_pdf(title_template.getvalue())
                    )
                except Exception:
                    pass

            prompt = build_prompt(system_name, component, purpose, ref_text, ftt_text, extras_text)
            client = OpenAI(api_key=api_key.strip())
            st.write("Анализирую BPMN-схему и готовлю текст документа…")

            content = [{"type":"input_text","text":prompt}]
            if diagram_image:
                content.append({"type":"input_image","image_url":diagram_image})
            if diagram_text:
                content.append({"type":"input_text","text":"ТЕКСТ BPMN/XML СХЕМЫ:\n"+diagram_text[:30000]})
            if diagram.name.lower().endswith(".pdf"):
                content.append({"type":"input_text","text":"ТЕКСТ ИЗ PDF-СХЕМЫ:\n"+diagram_text[:25000]})

            response = client.responses.create(
                model=model,
                input=[{"role":"user","content":content}],
                max_output_tokens=12000
            )
            generated = response.output_text
            st.write("Формирую DOCX с титульным листом и приложением…")
            fields = {
                "system": system_name.strip(),
                "component": component.strip(),
                "topic": topic.strip(),
                "executor": executor.strip(),
                "coexecutor": coexecutor.strip(),
                "date": doc_date.strip()
            }
            docx_bytes = make_docx(generated, fields, diagram if diagram.name.lower().endswith((".png",".jpg",".jpeg")) else None)
            status.update(label="Пояснительная записка сформирована", state="complete", expanded=False)
            st.success("Документ готов.")
            st.download_button(
                "⬇️ Скачать пояснительную записку (DOCX)",
                data=docx_bytes,
                file_name=f"ПЗ_{re.sub(r'[^А-Яа-яA-Za-z0-9_-]+','_',component)}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )
            if diagram.name.lower().endswith(".pdf"):
                st.download_button("⬇️ Скачать исходную BPMN-схему (PDF)", data=diagram.getvalue(),
                                   file_name=diagram.name, mime="application/pdf")
            st.info("Проверьте итоговый документ: корректность распознавания схемы, названия и сведения на титульном листе.")
        except Exception as e:
            status.update(label="Не удалось сформировать документ", state="error")
            st.exception(e)

st.divider()
st.caption("Веб-приложение для внутренней подготовки документов. Загруженные материалы обрабатываются в рамках текущего сеанса; не вводите секретные или ограниченные к распространению данные без согласования с владельцем сервиса.")
