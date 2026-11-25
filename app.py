import io
import os
import streamlit as st
from docx import Document
from pptx import Presentation
from google import genai

st.set_page_config(
    page_title="သင်တန်းဆင်း Report ထုတ်ပေးသည့် စနစ်",
    page_icon="📄",
    layout="wide",
)

st.title("📄 သင်တန်းဆင်း Report ထုတ်ပေးသည့် စနစ် (Streamlit)")
st.caption("Report နမူနာပုံစံနှင့် သင်ခန်းစာ PowerPoint (.pptx) များကို တင်ပြီး မြန်မာလို Report ပြန်ထုတ်ယူပါ။")


def get_api_key() -> str:
    if "GEMINI_API_KEY" in st.secrets:
        return st.secrets["GEMINI_API_KEY"]
    if "GEMINI_API_KEY" in os.environ:
        return os.environ["GEMINI_API_KEY"]
    return st.sidebar.text_input("Gemini API Key ထည့်ပါ", type="password")


def extract_docx_text(uploaded_file) -> str:
    doc = Document(io.BytesIO(uploaded_file.getvalue()))
    parts = []
    for p in doc.paragraphs:
        if p.text.strip():
            parts.append(p.text.strip())
    for idx, table in enumerate(doc.tables, start=1):
        parts.append(f"[Table {idx}]")
        for row in table.rows:
            row_text = " | ".join(cell.text.strip().replace("\n", " ") for cell in row.cells)
            if row_text.strip():
                parts.append(row_text)
    return "\n".join(parts)


def extract_pptx_text(uploaded_file) -> str:
    prs = Presentation(io.BytesIO(uploaded_file.getvalue()))
    slide_chunks = [f"=== PowerPoint ဖိုင်: {uploaded_file.name} ==="]
    for s_idx, slide in enumerate(prs.slides, start=1):
        texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    line = "".join(run.text for run in para.runs).strip() or para.text.strip()
                    if line:
                        texts.append(line)
            if shape.has_table:
                for row in shape.table.rows:
                    row_str = " | ".join(cell.text.strip().replace("\n", " ") for cell in row.cells)
                    if row_str.strip():
                        texts.append(f"[ဇယား] {row_str}")
        if texts:
            slide_chunks.append(f"--- Slide {s_idx} ---\n" + "\n".join(texts))
    return "\n\n".join(slide_chunks)


def build_docx_report(report_markdown: str, title: str = "သင်တန်းဆင်း အစီရင်ခံစာ") -> bytes:
    doc = Document()
    doc.add_heading(title, level=0)

    for raw_line in report_markdown.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=3)
        elif line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=2)
        elif line.startswith("# "):
            doc.add_heading(line[2:].strip(), level=1)
        elif line.startswith(("- ", "* ")):
            doc.add_paragraph(line[2:].strip(), style="List Bullet")
        else:
            doc.add_paragraph(line)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def generate_myanmar_report(api_key: str, template_text: str, ppt_text: str, extra_note: str) -> str:
    client = genai.Client(api_key=api_key)

    prompt = f"""
သင်သည် အစိုးရဌာန/အဖွဲ့အစည်းများအတွက် သင်တန်းဆင်း အစီရင်ခံစာ (Training Completion Report) ကို စနစ်တကျ ရေးသားပေးနိုင်သော အတွေ့အကြုံရှိ စာရေးသူတစ်ဦး ဖြစ်သည်။

အောက်ပါ အချက်အလက်များကို အခြေခံ၍ သင်တန်းဆင်း Report တစ်စောင်ကို **မြန်မာဘာသာဖြင့်** ပြည့်စုံရှင်းလင်းစွာ ပြန်လည်ရေးသားပေးပါ။

### လိုက်နာရမည့် စည်းကမ်းများ
1. **Report နမူနာပုံစံ (Template) ၏ Flow/အစီအစဉ်အတိုင်း** ခေါင်းစဉ်ခွဲများ၊ အပိုင်းများနှင့် တင်ပြပုံအစီအစဉ်ကို အတတ်နိုင်ဆုံး လိုက်နာရေးသားပါ။
2. သင်ခန်းစာ PowerPoint ဖိုင်များထဲမှ အဓိက သဘောတရားများ၊ ခေါင်းစဉ်များ၊ အချက်အလက်များနှင့် လက်တွေ့အသုံးချနိုင်မည့် အချက်များကို သင်တန်းဆင်း Report ထဲသို့ ထည့်သွင်းပေါင်းစပ်ပါ။
3. PowerPoint ထဲတွင် တိုက်ရိုက်မပါဝင်သော်လည်း Template အရ လိုအပ်သော အပိုင်းများ (ဥပမာ- နိဒါန်း၊ သင်တန်း၏ ရည်ရွယ်ချက်၊ သင်ခန်းစာအနှစ်ချုပ်၊ နိဂုံးနှင့် အကြံပြုချက်) ကို သင့်လျော်သော ဌာနဆိုင်ရာ/သင်တန်းဆိုင်ရာ အရေးအသားဖြင့် ဖြည့်စွက်ရေးသားပါ (မသေချာသော နေ့ရက်၊ လူအမည်၊ ကိန်းဂဏန်းများကို မတီထွင်ဘဲ `[...]` ဖြင့် ဖြည့်စွက်ရန် နေရာချန်ထားပါ)။
4. ဘာသာစကားကို **မြန်မာဘာသာ** သီးသန့်သုံးပါ (နည်းပညာဝေါဟာရ လိုအပ်ပါက မြန်မာစာလုံးပေါင်း သို့မဟုတ် ကွင်းစကွင်းပိတ်ဖြင့် ရှင်းလင်းစွာ ဖော်ပြပါ)။

### ၁။ Report နမူနာပုံစံ (Template Flow & Structure)
{template_text}

### ၂။ သင်ခန်းစာ PowerPoint များမှ ထုတ်ယူထားသော အချက်အလက်များ
{ppt_text}

### ၃။ ထပ်မံမှာကြားချက် (ရှိပါက)
{extra_note if extra_note.strip() else "မရှိပါ"}
""".strip()

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
    )
    return response.text or ""


# --- UI ---
api_key = get_api_key()

col1, col2 = st.columns(2)

with col1:
    st.subheader("၁။ Report နမူနာပုံစံ Upload")
    template_file = st.file_uploader(
        "နမူနာ Report ဖိုင် (.docx) တင်ပါ",
        type=["docx"],
        accept_multiple_files=False,
    )

with col2:
    st.subheader("၂။ သင်ခန်းစာ PowerPoint ဖိုင်များ Upload")
    ppt_files = st.file_uploader(
        "PowerPoint ဖိုင်များ (.pptx) တစ်ခု သို့မဟုတ် များစွာ တင်ပါ",
        type=["pptx"],
        accept_multiple_files=True,
    )

extra_note = st.text_area(
    "ထပ်မံထည့်သွင်းလိုသည့် ညွှန်ကြားချက် (Optional)",
    placeholder="ဥပမာ - သင်တန်းခေါင်းစဉ်၊ အဓိက အလေးပေးလိုသော အပိုင်း၊ အရေးအသား ပုံစံ...",
    height=90,
)

if st.button("🚀 သင်တန်းဆင်း Report ထုတ်မည်", type="primary", use_container_width=True):
    if not api_key:
        st.error("Gemini API Key ထည့်သွင်းရန် လိုအပ်ပါသည်။ (Streamlit Secrets သို့မဟုတ် Sidebar တွင် ထည့်ပါ)")
    elif not template_file:
        st.error("Report နမူနာပုံစံ (.docx) ဖိုင်ကို အရင် Upload တင်ပါ။")
    elif not ppt_files:
        st.error("သင်ခန်းစာ PowerPoint (.pptx) ဖိုင် အနည်းဆုံးတစ်ခု Upload တင်ပါ။")
    else:
        with st.spinner("Template နှင့် PowerPoint ဖိုင်များကို ဖတ်ပြီး Report ပြုစုနေပါသည်..."):
            try:
                template_text = extract_docx_text(template_file)
                ppt_texts = [extract_pptx_text(f) for f in ppt_files]
                combined_ppt_text = "\n\n".join(ppt_texts)

                if not template_text.strip():
                    st.warning("နမူနာ Report ဖိုင်ထဲတွင် စာသား ရှာမတွေ့ပါ။")
                elif not combined_ppt_text.strip():
                    st.warning("တင်ထားသော PowerPoint ဖိုင်များထဲတွင် စာသား ရှာမတွေ့ပါ။")
                else:
                    report_md = generate_myanmar_report(
                        api_key=api_key,
                        template_text=template_text,
                        ppt_text=combined_ppt_text,
                        extra_note=extra_note,
                    )
                    st.session_state["generated_report"] = report_md
                    st.success("သင်တန်းဆင်း Report ထုတ်ပြီးပါပြီ။")
            except Exception as e:
                st.error(f"Error ဖြစ်ပွားခဲ့ပါသည်: {e}")

if "generated_report" in st.session_state and st.session_state["generated_report"]:
    st.divider()
    st.subheader("📝 ထွက်လာသော သင်တန်းဆင်း Report (Preview)")
    st.markdown(st.session_state["generated_report"])

    docx_bytes = build_docx_report(st.session_state["generated_report"])

    d_col1, d_col2 = st.columns(2)
    with d_col1:
        st.download_button(
            label="⬇️ Word ဖိုင် (.docx) ဒေါင်းလုဒ်ဆွဲရန်",
            data=docx_bytes,
            file_name="training_completion_report_mm.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with d_col2:
        st.download_button(
            label="⬇️ Markdown ဖိုင် (.md) ဒေါင်းလုဒ်ဆွဲရန်",
            data=st.session_state["generated_report"].encode("utf-8"),
            file_name="training_completion_report_mm.md",
            mime="text/markdown",
            use_container_width=True,
        )
