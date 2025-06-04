import streamlit as st
import pandas as pd
import openpyxl
import fitz  # PyMuPDF
import io
from openai import OpenAI
from tabulate import tabulate

# 初始化 OpenAI
client = OpenAI(
    api_key=st.secrets["OPENAI_API_KEY"]
)

st.set_page_config(page_title="AI-Powered Course Update Assistant", layout="centered")
st.title("📚 AI-Powered Course Update Assistant")

st.markdown("Upload your **Info Gathering Excel file** and one or more **Old Training PDFs**, and let AI suggest update actions for your course content.")

# 上传文件
info_file = st.file_uploader("📄 Upload Info Gathering Excel", type=["xlsx"])
pdf_files = st.file_uploader("📘 Upload Old Training PDFs (multiple allowed)", type=["pdf"], accept_multiple_files=True)

if info_file and pdf_files:
    st.success("✅ Files uploaded successfully.")
    
    # 读取 Excel
    info_df = pd.read_excel(info_file)

    # 只保留 Permission Needed 为 Administrator Level 的行
    filtered_df = info_df[
        info_df["Permissions Needed"].astype(str).str.strip().str.lower() == "administrator level"
    ]

    # 如果没有满足条件的行，给出提示
    if filtered_df.empty:
        st.warning("⚠️ No features found with 'Permissions Needed' = Administrator Level.")
        st.stop()

    # 将选中的列转为 Markdown 表格
    try:
        feature_md_table = filtered_df[[
            "Feature Name",
            "Feature Overview / Description",
            "Scenario/Real Life Use Cases"
        ]].to_markdown(index=False)
    except Exception as e:
        st.error(f"❌ Failed to generate feature markdown table: {e}")
        st.stop()

    # 提取 PDF 文本
    pdf_text = ""
    for pdf in pdf_files:
        doc = fitz.open(stream=pdf.read(), filetype="pdf")
        for page in doc:
            pdf_text += page.get_text()

    # 构造 Prompt
    prompt = f"""
You are an expert Instructional Designer working on internal Admin training in Rise.

You are given:
1. A set of **PDF exports** of Rise Admin Training modules. Each PDF represents **one full Module**.
2. A table of **product feature updates**, including which permissions level each update applies to.

🧠 Important Structure Rules:
- **Each PDF file is one Module**, and the **Module name must be taken directly from the file name** (e.g., `End-to-End DataGuide Document Template Workflow`)
- **Lesson** corresponds to the individual lesson titles (usually top-level section titles within the PDF).
- **Specific Area** refers to the exact area within the lesson: "first paragraph", "demo hotspot", "screenshot block", etc.

Your job:
- Only review the feature updates where the column **Permissions Needed** is **"Administrator Level"**.
- Among these, **only select those that impact how Admins should use, configure, or understand the system**. Skip those that do not affect Admin workflows.
- Compare the relevant features to the training content.
- Propose precise edits using **instructional tone**, not technical writing.

Output a markdown table with these columns:

- **Location**: Format `Module > Lesson > Specific Area`, e.g.:
  - `End-to-End DataGuide Document Template Workflow > Course Introduction > first paragraph`
  - `End-to-End DataGuide Form Workflow > Lesson 2 > Demo: Submit Form`
- **Suggested Changes**:  
  - Be instructional and clear.  
  - If it's a paragraph update, give the **full revised version**.  
  - If it’s a demo/graphic/hotspot, describe what should be changed and why.
- **RISE Update Action**: Select **one** of:
  `New Text`, `New Header Section`, `New Lesson Block`, `New Lesson`, `Update/Modify Text`, `Base Org Update`, `New Demo/Transcript`, `Update Existing Demo/Transcript`, `New Interaction/Graphic`, `Update Existing Interaction/Graphic`
- **Feature Name**: Must match the `Feature Name` field from the Info Gathering Sheet.

Only return rows where `Permissions Needed = Administrator Level` and the feature is relevant to Admin training.

Training Content (excerpt):
{pdf_text[:8000]}

Feature Update Info Table:
{feature_md_table[:3000]}
"""

    with st.spinner("🤖 AI is analyzing content..."):
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}]
        )
    result_md = response.choices[0].message.content

    st.markdown("### 📝 AI Suggested Updates")
    st.markdown(result_md)

    # Markdown 转 DataFrame
    try:
        rows = [r.strip() for r in result_md.strip().split("\n") if r.startswith("|")]
        header = [h.strip() for h in rows[0].split("|")[1:-1]]
        data = [[cell.strip() for cell in row.split("|")[1:-1]] for row in rows[2:]]
        result_df = pd.DataFrame(data, columns=header)
    except Exception as e:
        st.error(f"⚠️ Failed to parse AI response. Please verify table format.\n\n{str(e)}")
        st.stop()

    st.markdown("### 📥 Download Your Feedback")
    output = io.BytesIO()
    result_df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)

    st.download_button(
        label="⬇️ Download Feedback as Excel",
        data=output,
        file_name="AI_Feedback.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

