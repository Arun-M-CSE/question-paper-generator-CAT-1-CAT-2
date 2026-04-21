from fpdf import FPDF
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os
import re

def sanitize_text(text):
    """Clean text for FPDF compatibility (Standard fonts only support Latin-1)."""
    if not text:
        return ""
    # Map common unicode characters to latin-1 or similar
    chars = {
        '\u2013': '-', '\u2014': '-', 
        '\u2018': "'", '\u2019': "'", 
        '\u201c': '"', '\u201d': '"',
        '\u2022': '*', '\u2026': '...'
    }
    for u_char, r_char in chars.items():
        text = text.replace(u_char, r_char)
    return text.encode('latin-1', 'replace').decode('latin-1')

def create_pdf(paper):
    """Creates a PDF file from the generated question paper."""
    pdf = FPDF()
    pdf.add_page()
    
    # Header
    pdf.set_font("Helvetica", 'B', 16)
    pdf.multi_cell(0, 10, sanitize_text("DR.M.G.R. EDUCATIONAL AND RESEARCH INSTITUTE"), align='C')
    pdf.ln(2)
    pdf.set_font("Helvetica", 'B', 14)
    pdf.cell(0, 8, f"{paper.get('exam_type', 'CAT')} Examination", ln=True, align='C')
    pdf.ln(8)
    
    # Part A
    pdf.set_font("Helvetica", 'B', 12)
    pdf.cell(0, 8, "PART A - (10 x 1 = 10 Marks)", ln=True)
    pdf.set_font("Helvetica", '', 10)
    
    for q in paper.get('part_a', []):
        opts = q.get('options', {"A": "N/A", "B": "N/A", "C": "N/A", "D": "N/A"})
        pdf.set_font("Helvetica", 'B', 10)
        pdf.multi_cell(0, 6, sanitize_text(f"Q{q.get('number', '')}. {q.get('question', 'N/A')}"), 0, 'L')
        pdf.ln(1)
        pdf.set_font("Helvetica", '', 9)
        pdf.cell(10, 5, "")
        pdf.cell(0, 5, sanitize_text(f"A) {opts.get('A', 'N/A')}"), 0, 0, 'L')
        pdf.cell(40, 5, "")
        pdf.cell(0, 5, sanitize_text(f"B) {opts.get('B', 'N/A')}"), 0, 1, 'L')
        
        pdf.cell(10, 5, "")
        pdf.cell(0, 5, sanitize_text(f"C) {opts.get('C', 'N/A')}"), 0, 0, 'L')
        pdf.cell(40, 5, "")
        pdf.cell(0, 5, sanitize_text(f"D) {opts.get('D', 'N/A')}"), 0, 1, 'L')
        
        pdf.set_font("Helvetica", 'I', 7)
        pdf.cell(10, 4, sanitize_text(f"(Unit: {q.get('unit', 'N/A')}, CO: {q.get('co', 'N/A')}, Blooms: {q.get('blooms', 'N/A')})"), 0, 1, 'L')
        pdf.ln(3)
        
    pdf.ln(5)
    
    # Part B
    pdf.set_font("Helvetica", 'B', 12)
    if paper.get('exam_type') == 'CAT1':
        pdf.cell(0, 8, "PART B - (2 x 10 = 20 Marks)", ln=True)
    else:
        pdf.cell(0, 8, "PART B - (4 x 10 = 40 Marks)", ln=True)
    pdf.set_font("Helvetica", '', 10)
    
    for q in paper.get('part_b', []):
        pdf.set_font("Helvetica", 'B', 10)
        pdf.multi_cell(0, 6, sanitize_text(f"Q{q.get('number', '')}. {q.get('question', 'N/A')}"), 0, 'L')
        pdf.ln(1)
        pdf.set_font("Helvetica", 'I', 7)
        pdf.cell(0, 4, sanitize_text(f"(Unit: {q.get('unit', 'N/A')}, CO: {q.get('co', 'N/A')}, Blooms: {q.get('blooms', 'N/A')})"), 0, 0, 'L')
        pdf.set_font("Helvetica", 'B', 8)
        pdf.cell(0, 4, "10 MARKS", 0, 1, 'R')
        pdf.ln(8)
        
    return bytes(pdf.output())

def create_word_doc(paper, template_path=None, output_file=None):
    """Creates a Word Document by filling a template dynamically."""
    if not template_path:
        target = "CAT 1 AI.docx" if paper['exam_type'] == "CAT1" else "CAT-II CD QPP.docx"
        template_path = target if os.path.exists(target) else None

    if template_path and os.path.exists(template_path):
        doc = docx.Document(template_path)
    else:
        doc = docx.Document()

    # 1. Fill Question Paper Table (Part A and Part B)
    if len(doc.tables) >= 1:
        question_table = doc.tables[0]
        
        # Map columns dynamically (Check first 15 rows for header)
        col_map = {"no": 0, "q": 1, "marks": 2, "co": 3, "bl": 4}
        header_row_idx = 0
        for r_idx, row in enumerate(question_table.rows[:15]):
            txt = " ".join([c.text.strip().upper() for c in row.cells])
            if "Q.NO" in txt or "QUESTION" in txt:
                header_row_idx = r_idx
                for c_idx, cell in enumerate(row.cells):
                    t = cell.text.strip().upper()
                    if "Q" in t and ("NO" in t or t == "Q.N" or t == "Q.NO"): col_map["no"] = c_idx
                    elif "QUESTION" in t: col_map["q"] = c_idx
                    elif "MARKS" in t: col_map["marks"] = c_idx
                    elif "CO" in t: col_map["co"] = c_idx
                    elif "BL" in t or "BLOOM" in t: col_map["bl"] = c_idx
                break
        
        def fill_questions(qs, start_row):
            for q in qs:
                num = str(q.get('number', ''))
                clean_num = "".join(filter(str.isdigit, num))
                if not clean_num: continue
                
                for r_idx in range(start_row, len(question_table.rows)):
                    row = question_table.rows[r_idx]
                    cell0 = row.cells[col_map["no"]].text.strip()
                    clean_cell0 = "".join(filter(str.isdigit, cell0))
                    
                    if clean_cell0 == clean_num:
                        # Fill Question Text
                        q_text = q.get('question', 'N/A')
                        if 'options' in q and q.get('options'):
                            o = q['options']
                            q_text += f"\nA) {o.get('A','')}\nB) {o.get('B','')}\nC) {o.get('C','')}\nD) {o.get('D','')}"
                        row.cells[col_map["q"]].text = q_text
                        
                        # Fill Metadata
                        if col_map["marks"] < len(row.cells):
                            row.cells[col_map["marks"]].text = str(q.get('marks', '10' if int(clean_num) > 10 else '1'))
                        if col_map["co"] < len(row.cells):
                            row.cells[col_map["co"]].text = q.get('co', 'CO1')
                        if col_map["bl"] < len(row.cells):
                            row.cells[col_map["bl"]].text = q.get('blooms', 'L1')
                        break

        data_start = header_row_idx + 1
        fill_questions(paper.get('part_a', []), data_start)
        fill_questions(paper.get('part_b', []), data_start)

    # 2. Fill CO Definitions Table
    cos_defs = paper.get('cos_definitions', [])
    if cos_defs:
        co_dict = {c['id'].upper().replace(" ", ""): c['description'] for c in cos_defs}
        for table in doc.tables:
            if len(table.rows) > 4:
                first_row_txt = " ".join([c.text.strip().upper() for c in table.rows[0].cells])
                if any(k in first_row_txt for k in ["CO1", "CO", "COURSE OUTCOME"]):
                    for row in table.rows:
                        label = row.cells[0].text.strip().upper().replace(" ", "").replace("-", "")
                        if label in co_dict:
                            row.cells[1].text = co_dict[label]
                        elif "COURSEOUTCOME" in label:
                            digit = "".join(filter(str.isdigit, label))
                            if digit and f"CO{digit}" in co_dict:
                                row.cells[1].text = co_dict[f"CO{digit}"]

    # 3. Fill Marks Distribution Table
    dist = paper.get('distributions', {})
    if dist:
        for table in doc.tables:
            if len(table.rows) >= 2:
                header_row_txt = " ".join([c.text.strip().upper() for c in table.rows[0].cells])
                if ("CO1" in header_row_txt and "CO2" in header_row_txt) or "COURSE OUTCOME" in header_row_txt:
                    headers = [c.text.strip().upper().replace(" ", "") for c in table.rows[0].cells]
                    mark_row = table.rows[1]
                    for co_id, marks in dist.get('co', {}).items():
                        cid = co_id.upper().replace(" ", "")
                        if cid in headers: mark_row.cells[headers.index(cid)].text = str(marks)
                    for bl_id, marks in dist.get('blooms', {}).items():
                        bid = bl_id.upper().replace(" ", "")
                        if bid in headers: mark_row.cells[headers.index(bid)].text = str(marks)

    # 4. Fill Answer Key Table (Now skips Table 0 and uses STRICT detection)
    all_qs = paper.get('part_a', []) + paper.get('part_b', [])
    if all_qs:
        for t_idx, table in enumerate(doc.tables):
            if t_idx == 0: continue # CRITICAL: Skip the main question paper table to prevent overwriting!
            
            ak_header_row = -1
            target_cols = {"no": -1, "ans": -1, "marks": -1}
            
            # Scan top 20 rows for Answer Key headers
            for r_idx in range(min(20, len(table.rows))):
                row_cells = table.rows[r_idx].cells
                row_txt = " ".join([c.text.strip().upper() for c in row_cells])
                
                # Rule out the main question paper table if it somehow got here
                if "QUESTION" in row_txt and "MARKS" in row_txt and "CO" in row_txt and "BL" in row_txt:
                    continue
                
                temp_cols = {"no": -1, "ans": -1, "marks": -1}
                for c_idx, cell in enumerate(row_cells):
                    t = cell.text.strip().upper()
                    if ("Q" in t and "NO" in t) or t == "Q.N" or t == "Q.NO": temp_cols["no"] = c_idx
                    elif "CONTENTS" in t or "ANSWER" in t: temp_cols["ans"] = c_idx
                    elif "MARKS" in t or "ALLOCATION" in t: temp_cols["marks"] = c_idx
                
                # Stricter check: Must have "CONTENTS" or "ALLOCATION" to be considered AK
                has_ak_marker = any(k in row_txt for k in ["CONTENTS", "ALLOCATION"])
                
                if temp_cols["no"] != -1 and temp_cols["ans"] != -1 and has_ak_marker:
                    ak_header_row = r_idx
                    target_cols = temp_cols
                    break
            
            if ak_header_row != -1:
                # We found an Answer Key table!
                for q in all_qs:
                    num_str = str(q.get('number', ''))
                    clean_num = "".join(filter(str.isdigit, num_str))
                    if not clean_num: continue
                    
                    for r_idx in range(ak_header_row + 1, len(table.rows)):
                        row = table.rows[r_idx]
                        if target_cols["no"] < len(row.cells):
                            cell0 = row.cells[target_cols["no"]].text.strip()
                            clean_cell0 = "".join(filter(str.isdigit, cell0))
                            
                            if clean_cell0 == clean_num:
                                ans_text = ""
                                if 'options' in q and q.get('options'):
                                    key = q.get('answer', 'A')
                                    opt_val = q['options'].get(key, '')
                                    ans_text = f"{key}) {opt_val}"
                                else:
                                    ans_text = q.get('answer', 'N/A')
                                
                                if target_cols["ans"] < len(row.cells):
                                    row.cells[target_cols["ans"]].text = ans_text
                                if target_cols["marks"] < len(row.cells):
                                    m_val = str(q.get('marks', '10' if int(clean_num) > 10 else '1'))
                                    row.cells[target_cols["marks"]].text = m_val
                                break

    # Final Save
    if output_file:
        doc.save(output_file)
        return output_file
    else:
        f_out = io.BytesIO()
        doc.save(f_out)
        return f_out.getvalue()
