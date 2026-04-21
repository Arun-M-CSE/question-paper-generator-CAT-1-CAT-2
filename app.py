import streamlit as st
import pandas as pd
from utils import extract_units, extract_cos, generate_questions, BLOOMS_TAXONOMY, extract_text_from_file, split_syllabus_by_topics
from formatting import create_pdf, create_word_doc
import json
import os

# Page Configuration
st.set_page_config(
    page_title="CAT Architect",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Simple Dark Theme CSS
st.markdown("""
<style>
    /* Dark background */
    .stApp {
        background-color: #0a0a0a;
        color: #ffffff;
    }
    
    /* Simple header */
    h1 {
        color: #ffffff;
        font-weight: 300;
        margin-bottom: 0.5rem;
    }
    
    /* Subtle text */
    p, label {
        color: #999;
    }
    
    /* Clean buttons */
    div.stButton > button {
        background-color: #1a1a1a;
        color: #fff;
        border: 1px solid #333;
        border-radius: 6px;
        padding: 0.5rem 1rem;
        width: 100%;
    }
    
    div.stButton > button:hover {
        background-color: #222;
        border-color: #444;
    }
    
    /* File uploader */
    [data-testid="stFileUploadDropzone"] {
        border: 1px solid #333 !important;
        background: #0f0f0f !important;
        border-radius: 6px !important;
    }
    
    /* Selectbox */
    div[data-baseweb="select"] {
        background-color: #1a1a1a;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 2px;
        background-color: #0a0a0a;
    }
    
    .stTabs [data-baseweb="tab"] {
        background-color: #1a1a1a;
        color: #999;
        border-radius: 4px 4px 0 0;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #222;
        color: #fff;
    }
    
    /* Question display */
    .question {
        background: #0f0f0f;
        border: 1px solid #222;
        border-radius: 6px;
        padding: 1rem;
        margin-bottom: 1rem;
    }
    
    .question-meta {
        color: #666;
        font-size: 0.85rem;
        margin-bottom: 0.5rem;
    }
    
    /* Download buttons */
    .stDownloadButton > button {
        background-color: #1a1a1a;
        color: #fff;
        border: 1px solid #333;
    }
    
    /* Success/Info messages */
    .stSuccess, .stInfo {
        background-color: #1a1a1a;
        border: 1px solid #333;
        color: #fff;
    }

    /* Page-level credit block: follows sidebar open/close */
    .page-credit {
        position: fixed;
        left: 70px;
        bottom: 12px;
        z-index: 9996;
        width: min(480px, calc(100vw - 84px));
        background: rgba(255, 255, 255, 0.12);
        border: 1px solid #7a7a7a;
        border-radius: 8px;
        color: #f4f4f4;
        font-size: 0.95rem;
        font-weight: 600;
        line-height: 1.45;
        padding: 0.6rem 0.75rem;
        box-shadow: 0 4px 14px rgba(0, 0, 0, 0.35);
        backdrop-filter: blur(2px);
        transition: left 0.25s ease, width 0.25s ease;
    }

    [data-testid="stSidebar"][aria-expanded="true"] ~ [data-testid="stAppViewContainer"] .page-credit,
    [data-testid="stSidebar"][aria-expanded="true"] ~ div .page-credit {
        left: calc(21rem + 14px);
        width: min(480px, calc(100vw - 21rem - 28px));
    }

    [data-testid="stSidebar"][aria-expanded="false"] ~ [data-testid="stAppViewContainer"] .page-credit,
    [data-testid="stSidebar"][aria-expanded="false"] ~ div .page-credit {
        left: 70px;
        width: min(480px, calc(100vw - 84px));
    }

    /* Keep sidebar open/close controls above the credit card */
    [data-testid="collapsedControl"],
    [data-testid="stSidebarCollapseButton"] {
        z-index: 10001 !important;
    }

    @media (max-width: 768px) {
        .page-credit {
            left: 10px;
            right: 10px;
            bottom: 8px;
            width: auto;
            font-size: 0.85rem;
        }
    }
</style>
""", unsafe_allow_html=True)

# Main UI
st.title("CAT Architect")
st.caption("AI-powered question paper generator")

with st.sidebar:
    st.header("Groq API Key")
    groq_api_key = st.text_input(
        "Enter your Groq API key",
        type="password",
        key="groq_api_key_input",
        help="Required every time before generating a question paper."
    )
    if not groq_api_key:
        st.warning("API key is required. Enter it here before generating.")
    else:
        st.success("API key added for this session.")

st.markdown(
    """
    <div class="page-credit">
        <div>Implemented by: Vijay Srinivasan R</div>
        <div>Facilitated by: Arun M, ChinchuNair</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("---")

# File Upload
uploaded_file = st.file_uploader("Upload Syllabus", type=['pdf', 'docx', 'txt'])

if uploaded_file:
    # Check if a new file is uploaded to clear previous state
    if "current_file" not in st.session_state or st.session_state["current_file"] != uploaded_file.name:
        st.session_state["analyzed"] = False
        st.session_state["unit1"] = ""
        st.session_state["unit2"] = ""
        st.session_state["unit3"] = ""
        st.session_state["cos"] = []
        st.session_state["current_file"] = uploaded_file.name
        st.session_state["previous_questions"] = []
        if 'paper' in st.session_state: del st.session_state['paper']

    raw_text = extract_text_from_file(uploaded_file)
    st.session_state['raw_syllabus'] = raw_text
    
    if st.button("Analyze Syllabus"):
        with st.spinner("Analyzing..."):
            u1, u2, u3 = extract_units(raw_text)
            cos = extract_cos(raw_text)
            st.session_state['unit1'] = u1
            st.session_state['unit2'] = u2
            st.session_state['unit3'] = u3
            st.session_state['cos'] = cos
            st.session_state['analyzed'] = True
            st.session_state['previous_questions'] = []
            st.success("Analysis Complete")

# Post-Analysis
if st.session_state.get('analyzed'):
    st.markdown("---")
    
    # Show full syllabus content
    st.subheader("Extracted Syllabus Content")
    
    # Unit 1
    with st.expander("Unit 1 Syllabus", expanded=False):
        if st.session_state['unit1']:
            st.text_area("Unit 1 Content", st.session_state['unit1'], height=200, disabled=True, label_visibility="collapsed")
        else:
            st.warning("Unit 1 not found in syllabus")
    
    # Unit 2
    with st.expander("Unit 2 Syllabus", expanded=False):
        if st.session_state['unit2']:
            st.text_area("Unit 2 Content", st.session_state['unit2'], height=150, disabled=True, label_visibility="collapsed")
            
            # Splitting visualization
            st.markdown("---")
            st.markdown("### Coverage Segments")
            parts = split_syllabus_by_topics(st.session_state['unit2'], n=3)
            for p in parts:
                st.info(p)
        else:
            st.warning("Unit 2 not found in syllabus")
    
    # Unit 3
    with st.expander("Unit 3 Syllabus", expanded=False):
        if st.session_state['unit3']:
            st.text_area("Unit 3 Content", st.session_state['unit3'], height=150, disabled=True, label_visibility="collapsed")
            
            # Splitting visualization
            st.markdown("---")
            st.markdown("### Coverage Segments")
            parts = split_syllabus_by_topics(st.session_state['unit3'], n=3)
            for p in parts:
                st.info(p)
        else:
            st.warning("Unit 3 not found in syllabus")
    
    # Course Outcomes
    with st.expander("Course Outcomes", expanded=False):
        if st.session_state['cos']:
            for co in st.session_state['cos']:
                blooms_info = f" [{co['blooms']}]" if co.get('blooms') else ""
                st.markdown(f"**{co['id']}:** {co['description']}{blooms_info}")
        else:
            st.warning("No Course Outcomes found")
    
    st.markdown("---")
    
    # Generation options
    col1, col2 = st.columns(2)
    with col1:
        exam_type = st.selectbox("Exam Type", ["CAT1", "CAT2"], index=1)
    with col2:
        # Dynamic Difficulty Selectors
        difficulty_config = {}
        if exam_type == "CAT1":
            difficulty_config['u1'] = st.selectbox("Unit 1 Difficulty", ["Easy", "Medium", "Hard"], key="d_u1", index=1)
        else:
            sub_c1, sub_c2 = st.columns(2)
            with sub_c1:
                difficulty_config['u2'] = st.selectbox("Unit 2 Difficulty", ["Easy", "Medium", "Hard"], key="d_u2", index=1)
            with sub_c2:
                difficulty_config['u3'] = st.selectbox("Unit 3 Difficulty", ["Easy", "Medium", "Hard"], key="d_u3", index=1)
    
    if st.button("Generate Question Paper"):
        if not groq_api_key or not groq_api_key.strip():
            st.error("Groq API key is required. Enter it in the left sidebar and try again.")
        else:
            with st.spinner("Generating..."):
                try:
                    subject_name = uploaded_file.name.split('.')[0]
                    paper = generate_questions(
                        st.session_state['unit1'],
                        st.session_state['unit2'], 
                        st.session_state['unit3'], 
                        exam_type,
                        st.session_state['cos'], 
                        difficulty_config, 
                        BLOOMS_TAXONOMY,
                        subject_name=subject_name,
                        exclude_questions=st.session_state.get('previous_questions', []),
                        api_key=groq_api_key
                    )
                    
                    # Add newly generated questions to exclusions for next run
                    current_qs = [q.get('question') for q in paper.get('part_a', []) if q.get('question')]
                    current_qs += [q.get('question') for q in paper.get('part_b', []) if q.get('question')]
                    
                    if 'previous_questions' not in st.session_state:
                        st.session_state['previous_questions'] = []
                    st.session_state['previous_questions'].extend(current_qs)
                    
                    if 'error' in paper and not paper['part_a'] and not paper['part_b']:
                        st.error(f"Error: {paper['error']}")
                    else:
                        if 'validation_errors' in paper and paper['validation_errors']:
                            st.warning("⚠️ Potential Issues Detected:")
                            for err in paper['validation_errors']:
                                st.write(f"- {err}")
                                
                        st.session_state['paper'] = paper
                        
                        # Save to history
                        import datetime
                        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                        
                        if not os.path.exists("history"):
                            os.makedirs("history")
                        
                        safe_subject = "".join([c for c in subject_name if c.isalnum() or c in (' ', '_', '-')]).strip()
                        with open(f"history/{safe_subject}_{ts}.json", "w") as f:
                            json.dump(paper, f, indent=2)
                        
                        # Save Word file
                        folder_name = "cat1" if paper['exam_type'] == "CAT1" else "cat2"
                        if not os.path.exists(folder_name):
                            os.makedirs(folder_name)
                            
                        word_filename = os.path.join(folder_name, f"{paper['exam_type']}_{safe_subject}_{ts}.docx")
                        word_path = os.path.abspath(word_filename)
                        
                        from formatting import create_word_doc
                        saved_path = create_word_doc(paper, output_file=word_path)
                        
                        if saved_path:
                            st.session_state['generated_word_file'] = saved_path
                            st.success(f"Paper generated: {word_filename}")
                        
                except Exception as e:
                    st.error(f"Error: {e}")

# Display Paper
if 'paper' in st.session_state:
    paper = st.session_state['paper']
    
    st.markdown("---")
    st.subheader(f"{paper['exam_type']} Analysis & Export")

    # Distributions Visualization
    if 'distributions' in paper:
        dist = paper['distributions']
        col_v1, col_v2 = st.columns(2)
        
        import matplotlib.pyplot as plt
        
        with col_v1:
            st.write("**Course Outcome Vs Mark Distribution**")
            co_data = dist.get('co', {})
            if co_data and any(v > 0 for v in co_data.values()):
                fig, ax = plt.subplots(figsize=(6, 4))
                ax.bar(co_data.keys(), co_data.values(), color='#1f77b4')
                ax.set_ylabel('Marks')
                ax.set_facecolor('#0f0f0f')
                fig.patch.set_facecolor('#0f0f0f')
                ax.tick_params(colors='white')
                ax.yaxis.label.set_color('white')
                for i, v in enumerate(co_data.values()):
                    ax.text(i, v + 0.2, str(v), color='white', ha='center')
                st.pyplot(fig)
            else:
                st.info("No CO distribution data available")

        with col_v2:
            st.write("**Bloom's Level Vs Mark Distribution**")
            bl_data = dist['blooms']
            if bl_data:
                fig, ax = plt.subplots(figsize=(6, 4))
                # Sort levels for better display
                sorted_keys = sorted(bl_data.keys())
                sorted_vals = [bl_data[k] for k in sorted_keys]
                
                # Filter out zero values
                active_labels = [k for k, v in zip(sorted_keys, sorted_vals) if v > 0]
                active_vals = [v for v in sorted_vals if v > 0]
                
                if active_vals:
                    # Use pctdistance and labeldistance to prevent overlapping
                    ax.pie(active_vals, 
                           labels=active_labels, 
                           autopct='%1.1f%%', 
                           textprops={'color':"white"}, 
                           colors=plt.cm.Paired.colors,
                           pctdistance=0.8,
                           labeldistance=1.1,
                           startangle=140)
                    ax.set_facecolor('#0f0f0f')
                    fig.patch.set_facecolor('#0f0f0f')
                    st.pyplot(fig)
                else:
                    st.info("No mark distribution data")
            else:
                st.info("No Bloom's distribution available")

    st.markdown("---")
    st.subheader(f"{paper['exam_type']} Question Paper")
    
    # Download buttons
    col1, col2, col3 = st.columns(3)
    with col1:
        st.download_button("JSON", data=json.dumps(paper, indent=2), file_name=f"{paper['exam_type']}_Paper.json")
    with col2:
        pdf_bytes = create_pdf(paper)
        st.download_button("PDF", data=pdf_bytes, file_name=f"{paper['exam_type']}_Paper.pdf", mime="application/pdf")
    with col3:
        word_file_path = st.session_state.get('generated_word_file')
        if word_file_path and os.path.exists(word_file_path):
            with open(word_file_path, "rb") as f:
                st.download_button("Word", data=f, file_name=os.path.basename(word_file_path), mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    
    st.markdown("---")
    
    # Questions
    tab1, tab2 = st.tabs(["Part A (MCQ)", "Part B (Subjective)"])
    
    with tab1:
        for q in paper['part_a']:
            st.markdown(f"""
            <div class="question">
                <div class="question-meta">Q{q['number']} | Unit {q['unit']} | {q['co']} | {q['blooms']}</div>
                <p><strong>{q['question']}</strong></p>
                <p>A) {q['options']['A']}</p>
                <p>B) {q['options']['B']}</p>
                <p>C) {q['options']['C']}</p>
                <p>D) {q['options']['D']}</p>
            </div>
            """, unsafe_allow_html=True)
    
    with tab2:
        for q in paper['part_b']:
            st.markdown(f"""
            <div class="question">
                <div class="question-meta">Q{q['number']} | Unit {q['unit']} | {q['co']} | {q['blooms']} | 10 Marks</div>
                <p><strong>{q['question']}</strong></p>
            </div>
            """, unsafe_allow_html=True)

if not uploaded_file:
    st.info("Upload a syllabus file to begin")