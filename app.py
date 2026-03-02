import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Consolidated Marksheet Pro", layout="wide")

st.title("🏫 Student Exam Data Consolidator")

def custom_round(x):
    try:
        val = float(x)
        return int(np.floor(val + 0.5))
    except: return 0

def clean_marks(val):
    if isinstance(val, str):
        v = val.strip().upper()
        if v == 'AB' or v == '': return 0.0
    try: return float(val)
    except: return 0.0

uploaded_file = st.file_uploader("Upload Excel Marksheet", type="xlsx")

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        subj_list = [f'Sub{i+1}' for i in range(6)]
        fetch_subj_map = {f'SUB{i+1}': f'Sub{i+1}' for i in range(6)}
        
        exam_configs = [
            {'label': 'FIRST UNIT TEST (25)', 'sheets': ['FIRST UNIT TEST']},
            {'label': 'FIRST TERM EXAM (50)', 'sheets': ['FIRST TERM']},
            {'label': 'SECOND UNIT TEST (25)', 'sheets': ['SECOND UNIT TEST']},
            {'label': 'ANNUAL EXAM (70/80)', 'sheets': ['ANNUAL EXAM']}
        ]

        result_cols = ['Grand Total', '%', 'Result', 'Remark', 'Rank']
        all_students = {}

        for config in exam_configs:
            sheet_name = next((s for s in xl.sheet_names if s.strip().upper() in config['sheets']), None)
            if sheet_name:
                df = xl.parse(sheet_name)
                df.columns = df.columns.astype(str).str.strip().str.upper()
                t_col = next((c for c in df.columns if 'TOTAL' in c), None)
                p_col = next((c for c in df.columns if '%' in c or 'PERCENT' in c), None)
                r_col = next((c for c in df.columns if 'RESULT' in c), None)

                for _, row in df.iterrows():
                    roll = str(row.get('ROLL NO.', '')).strip()
                    if not roll or roll == 'nan': continue
                    if roll not in all_students:
                        all_students[roll] = {'Name': row.get('STUDENT NAME', 'Unknown'), 'Exams': {}}
                    
                    marks = {v: (str(row.get(k, 0)).strip() if str(row.get(k, 0)).strip().upper() == 'AB' else row.get(k, 0)) for k, v in fetch_subj_map.items()}
                    marks['Grand Total'] = str(row.get(t_col, '')) if t_col else ''
                    try:
                        raw_p = row.get(p_col, '')
                        marks['%'] = str(round(float(raw_p), 2)) if raw_p != '' else str(raw_p)
                    except: marks['%'] = str(row.get(p_col, ''))
                    marks['Result'] = str(row.get(r_col, ''))
                    all_students[roll]['Exams'][config['label']] = marks

        categories = ['FIRST UNIT TEST (25)', 'FIRST TERM EXAM (50)', 'SECOND UNIT TEST (25)', 
                      'ANNUAL EXAM (70/80)', 'INT/PRACTICAL (20/30)', 'Total Marks Out of 200', 
                      'Average Marks 200/2=100']

        rows = []
        for roll in sorted(all_students.keys(), key=lambda x: float(x) if x.replace('.','',1).isdigit() else 0):
            s = all_students[roll]
            for cat in categories:
                row_data = {'Roll No.': roll if cat == 'FIRST UNIT TEST (25)' else '',
                            'Column1': s['Name'] if cat == 'FIRST UNIT TEST (25)' else '',
                            'Column2': cat}
                for col in subj_list + result_cols: row_data[col] = ''
                if cat in s['Exams']: row_data.update(s['Exams'][cat])
                elif cat == 'INT/PRACTICAL (20/30)':
                    for sub in subj_list: row_data[sub] = "0"
                rows.append(row_data)

        base_df = pd.DataFrame(rows)
        for col in base_df.columns: base_df[col] = base_df[col].astype(str).replace('nan', '')

        edited_df = st.data_editor(base_df, hide_index=True, use_container_width=True)

        if st.button("Generate Final Report & Rank"):
            processed = []
            pass_list = []
            for i in range(0, len(edited_df), 7):
                block = edited_df.iloc[i:i+7].copy()
                nums = block.iloc[0:5][subj_list].applymap(clean_marks)
                t200 = nums.sum()
                for sub in subj_list: block.iloc[5, block.columns.get_loc(sub)] = str(int(t200[sub]))
                a100 = t200.apply(lambda x: custom_round(x/2))
                for sub in subj_list: block.iloc[6, block.columns.get_loc(sub)] = str(a100[sub])

                gt = a100.sum()
                pc = round((gt / 600) * 100, 2)
                isp = all(m >= 35 for m in a100)
                block.iloc[6, block.columns.get_loc('Grand Total')] = str(gt)
                block.iloc[6, block.columns.get_loc('%')] = str(pc)
                block.iloc[6, block.columns.get_loc('Result')] = "PASS" if isp else "FAIL"
                processed.append(block)
                if isp: pass_list.append({'idx': i+6, 'total': gt})

            final_df = pd.concat(processed).reset_index(drop=True)
            pass_list.sort(key=lambda x: x['total'], reverse=True)
            curr_r, last_t = 0, -1
            for j, e in enumerate(pass_list):
                if e['total'] != last_t: curr_r = j + 1
                last_t = e['total']
                final_df.at[e['idx'], 'Rank'] = str(curr_r)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Consolidated')
            st.success("✅ Ranks Applied!")
            st.download_button("📥 Download Excel", output.getvalue(), "Final_Consolidated.xlsx")
    except Exception as e: st.error(f"Error: {e}")