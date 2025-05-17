import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile
import os
import json
import language_tool_python  # Grammar checking
import pdfplumber  # PDF text extraction
from docx2python import docx2python  # DOCX text extraction
import pandas as pd

# -------------------- Text Processing Tools --------------------
def extract_text_from_file(uploaded_file):
    """Extract text from PDF, DOC, or DOCX files"""
    try:
        if uploaded_file.type == "application/pdf":
            with pdfplumber.open(uploaded_file) as pdf:
                text = "\n".join([page.extract_text() for page in pdf.pages])
            return text
        
        elif uploaded_file.type in ["application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                  "application/msword"]:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                tmp.write(uploaded_file.getvalue())
                docx_content = docx2python(tmp.name)
                text = docx_content.text
                os.unlink(tmp.name)
                return text
                
    except Exception as e:
        st.error(f"Text extraction error: {str(e)}")
        return ""

def check_grammar(text):
    """Check grammar and syntax using LanguageTool"""
    try:
        tool = language_tool_python.LanguageTool('en-US')
        matches = tool.check(text)
        return matches
    except Exception as e:
        st.error(f"Grammar check error: {str(e)}")
        return []

# -------------------- DocuMorph Engine --------------------
class DocuMorphEngine:
    def __init__(self, docx_file=None):
        self.document = Document(docx_file) if docx_file else Document()

    def set_font(self, font_name, font_size):
        for para in self.document.paragraphs:
            for run in para.runs:
                run.font.name = font_name
                run.font.size = Pt(font_size)

    def set_line_spacing(self, spacing):
        for para in self.document.paragraphs:
            para.paragraph_format.line_spacing = spacing

    def set_alignment(self, alignment):
        align_map = {
            "Left": WD_ALIGN_PARAGRAPH.LEFT,
            "Center": WD_ALIGN_PARAGRAPH.CENTER,
            "Right": WD_ALIGN_PARAGRAPH.RIGHT,
            "Justify": WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        for para in self.document.paragraphs:
            para.alignment = align_map.get(alignment, WD_ALIGN_PARAGRAPH.LEFT)

    def set_margins(self, top, bottom, left, right):
        sec = self.document.sections[0]
        sec.top_margin, sec.bottom_margin = Inches(top), Inches(bottom)
        sec.left_margin, sec.right_margin = Inches(left), Inches(right)

    def add_logo(self, image, width, height):
        hdr = self.document.sections[0].header
        if not hdr.paragraphs:
            hdr.add_paragraph()
        run = hdr.paragraphs[0].add_run()
        run.add_picture(image, width=Inches(width), height=Inches(height))

    def set_header_footer(self, h_text, f_text, size, align):
        align_map = {"Left": WD_ALIGN_PARAGRAPH.LEFT,
                     "Center": WD_ALIGN_PARAGRAPH.CENTER,
                     "Right": WD_ALIGN_PARAGRAPH.RIGHT}
        for sec in self.document.sections:
            if not sec.header.paragraphs:
                sec.header.add_paragraph()
            if not sec.footer.paragraphs:
                sec.footer.add_paragraph()

            h_para = sec.header.paragraphs[0]
            f_para = sec.footer.paragraphs[0]
            h_para.text, f_para.text = h_text, f_text
            if h_para.runs:
                h_para.runs[0].font.size = Pt(size)
            if f_para.runs:
                f_para.runs[0].font.size = Pt(size)
            h_para.alignment = f_para.alignment = align_map.get(align, WD_ALIGN_PARAGRAPH.LEFT)

    def add_section_title(self, title):
        self.document.add_heading(title, level=1)

    def add_bullet_list(self, items):
        for item in items:
            self.document.add_paragraph(item, style="List Bullet")

    def add_figure(self, image, w, h, caption="", pos="Below"):
        if pos == "Above" and caption:
            self.document.add_paragraph(caption, style="Caption")
        p = self.document.add_paragraph()
        run = p.add_run()
        run.add_picture(image, width=Inches(w), height=Inches(h))
        if pos == "Below" and caption:
            self.document.add_paragraph(caption, style="Caption")

    def save(self, path):
        self.document.save(path)

# -------------------- Template Manager --------------------
TEMPLATE_DIR = "templates"
os.makedirs(TEMPLATE_DIR, exist_ok=True)

def list_templates():
    return [f[:-5] for f in os.listdir(TEMPLATE_DIR) if f.endswith('.json')]

def load_template(name):
    path = os.path.join(TEMPLATE_DIR, f"{name}.json")
    return json.load(open(path)) if os.path.exists(path) else None

def save_template(name, cfg):
    with open(os.path.join(TEMPLATE_DIR, f"{name}.json"), 'w') as f:
        json.dump(cfg, f)

def delete_template(name):
    os.remove(os.path.join(TEMPLATE_DIR, f"{name}.json"))

# -------------------- Streamlit UI --------------------
st.set_page_config(page_title="DocuMorph AI Pro", layout="wide")

# Custom styles
st.markdown("""
<style>
    .stButton>button {width:100%; padding:0.75em;}
    .big-download .stDownloadButton>button {background-color:#4E79A7; color:white;}
    .grammar-error { color: red; font-weight: bold; }
    .grammar-suggestion { color: green; }
</style>
""", unsafe_allow_html=True)

# Sidebar: Templates
with st.sidebar:
    st.header("💾 Template Manager")
    templates = list_templates()
    sel = st.selectbox("Load template", ["<none>"] + templates)
    if sel != '<none>':
        cfg = load_template(sel)
    else:
        cfg = {}

    new_name = st.text_input("Save current as", key='new_temp')
    if st.button("💾 Save Template"):
        save_template(new_name, st.session_state)
        st.success(f"Template '{new_name}' saved!")

    if sel != '<none>' and st.button("🗑 Delete Template"):
        delete_template(sel)
        st.warning(f"Template '{sel}' deleted!")
        st.experimental_rerun()

# Main App
st.title("📄 DocuMorph AI Pro – Document Intelligence")

tabs = st.tabs(["Styling", "Logo & HF", "Content", "Text Analysis", "Export"])

# Styling Tab (unchanged)
with tabs[0]:
    st.subheader("🎨 Document Styling")
    c1, c2 = st.columns(2)
    with c1:
        font_name = st.selectbox("Font Style", ["Times New Roman", "Arial", "Calibri", "Georgia"], index=0)
        font_size = st.slider("Font Size", 8, 24, 12)
        line_spacing = st.slider("Line Spacing", 1.0, 2.0, 1.15, 0.05)
    with c2:
        alignment = st.radio("Alignment", ["Left", "Center", "Right", "Justify"], horizontal=True)
        margins = [
            st.number_input(label, 0.1, 3.0, 1.0, 0.1, key=label)
            for label in ["Top Margin", "Bottom Margin", "Left Margin", "Right Margin"]
        ]
        st.session_state['margins'] = margins

# Logo & Header/Footer Tab (unchanged)
with tabs[1]:
    st.subheader("🖼 Logo & Header/Footer")
    col1, col2 = st.columns(2)
    with col1:
        logo = st.file_uploader("Upload Logo", type=['png', 'jpg'])
        logo_w = st.slider("Logo Width (in)", 0.5, 4.0, 1.0, 0.1)
        logo_h = st.slider("Logo Height (in)", 0.5, 4.0, 1.0, 0.1)
    with col2:
        header_text = st.text_input("Header Text")
        footer_text = st.text_input("Footer Text")
        hf_size = st.slider("HF Font Size", 8, 20, 10)
        hf_align = st.selectbox("HF Alignment", ["Left", "Center", "Right"], index=1)

# Content Tab (renamed from Figures & Sections)
with tabs[2]:
    st.subheader("📑 Content Management")
    section_title = st.text_input("Section Title")
    bullets_input = st.text_area("Bullet List (one per line)")
    figure = st.file_uploader("Insert Figure", type=['png', 'jpg'], key='fig')
    fig_w = st.slider("Figure Width", 1.0, 6.0, 4.0)
    fig_h = st.slider("Figure Height", 1.0, 6.0, 3.0)
    caption = st.text_input("Caption")
    caption_pos = st.radio("Caption Position", ['Above', 'Below'], horizontal=True)

# New Text Analysis Tab
with tabs[3]:
    st.subheader("🔍 Text Analysis & Grammar Check")
    
    # Document Analysis Section
    analysis_file = st.file_uploader("Upload Document for Analysis", 
                                   type=['pdf', 'docx', 'doc'],
                                   key='analysis_file')
    
    if analysis_file:
        with st.spinner("Extracting text..."):
            extracted_text = extract_text_from_file(analysis_file)
            
            if extracted_text:
                st.subheader("Extracted Text Preview")
                st.text_area("Full Text", extracted_text, height=200, key='extracted_text')
                
                if st.button("Run Grammar Check"):
                    with st.spinner("Analyzing document..."):
                        matches = check_grammar(extracted_text)
                        
                        if matches:
                            st.warning(f"Found {len(matches)} potential issues")
                            
                            # Create a dataframe for better display
                            issues = []
                            for match in matches[:50]:  # Limit to first 50 issues
                                context_start = max(0, match.offset-20)
                                context = extracted_text[context_start:match.offset+20]
                                issues.append({
                                    "Error": match.message,
                                    "Suggested": match.replacements[0] if match.replacements else "",
                                    "Context": f"...{context}..."
                                })
                            
                            df = pd.DataFrame(issues)
                            st.dataframe(df.style
                                .applymap(lambda x: 'color: red' if x == df['Error'][0] else '')
                            
                            # Show detailed examples
                            st.subheader("Top 5 Issues")
                            for i, match in enumerate(matches[:5]):
                                st.markdown(f"""
                                **{i+1}. {match.message}**  
                                <span class="grammar-error">Error:</span> {extracted_text[match.offset:match.offset+match.errorLength]}  
                                <span class="grammar-suggestion">Suggested:</span> {match.replacements[0] if match.replacements else "None"}  
                                """, unsafe_allow_html=True)
                        else:
                            st.success("No grammar issues found!")

# Export Tab (unchanged)
with tabs[4]:
    st.subheader("📤 Generate & Download")
    uploaded_file = st.file_uploader("Upload DOCX File", type=['docx'])
    if st.button("📝 Generate & Download", key='gen'):
        if not uploaded_file:
            st.error("Please upload a DOCX first!")
        else:
            engine = DocuMorphEngine(uploaded_file)
            engine.set_font(font_name, font_size)
            engine.set_line_spacing(line_spacing)
            engine.set_alignment(alignment)
            engine.set_margins(*st.session_state['margins'])
            if logo:
                logo.seek(0)
                engine.add_logo(logo, logo_w, logo_h)
            engine.set_header_footer(header_text, footer_text, hf_size, hf_align)
            if section_title.strip():
                engine.add_section_title(section_title.strip())
            bullets = [b.strip() for b in bullets_input.split("\n") if b.strip()]
            if bullets:
                engine.add_bullet_list(bullets)
            if figure:
                figure.seek(0)
                engine.add_figure(figure, fig_w, fig_h, caption, caption_pos)

            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
            engine.save(tmp.name)
            tmp.close()
            data = open(tmp.name, 'rb').read()
            st.download_button(
                label="⬇ Download Document",
                data=data,
                file_name='formatted.docx',
                mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                use_container_width=True,
                key='download',
                css_class='big-download'
            )
            os.unlink(tmp.name)