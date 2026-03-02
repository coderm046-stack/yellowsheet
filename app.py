import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Consolidated Marksheet Pro", layout="wide")
st.title("🏫 Student Exam Data Consolidator")

def custom_round(x):
    try:
        val = float(x)
        return int(np.floor(val + 0.5))
    except:
        return 0

def clean_marks(val):
    if isinstance(val, str):
        v = val.strip().upper()
        if v == 'AB' or v == '':
            return 0.0
    try:
        return float(val)
    except:
        return 0.0

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
                    if not roll or roll == 'nan':
                        continue
                    if roll not in all_students:
                        all_students[roll] = {'Name': row.get('STUDENT NAME', 'Unknown'), 'Exams': {}}

                    marks = {v: (str(row.get(k, 0)).strip() if str(row.get(k, 0)).strip().upper() == 'AB' else row.get(k, 0))
                             for k, v in fetch_subj_map.items()}
                    marks['Grand Total'] = str(row.get(t_col, '')) if t_col else ''
                    try:
                        raw_p = row.get(p_col, '')
                        marks['%'] = str(round(float(raw_p), 2)) if raw_p != '' else str(raw_p)
                    except:
                        marks['%'] = str(row.get(p_col, ''))
                    marks['Result'] = str(row.get(r_col, ''))
                    all_students[roll]['Exams'][config['label']] = marks

        categories = ['FIRST UNIT TEST (25)', 'FIRST TERM EXAM (50)', 'SECOND UNIT TEST (25)',
                      'ANNUAL EXAM (70/80)', 'INT/PRACTICAL (20/30)', 'Total Marks Out of 200',
                      'Average Marks 200/2=100']

        rows = []
        for roll in sorted(all_students.keys(), key=lambda x: float(x) if x.replace('.', '', 1).isdigit() else 0):
            s = all_students[roll]
            for cat in categories:
                row_data = {'Roll No.': roll if cat == 'FIRST UNIT TEST (25)' else '',
                            'Column1': s['Name'] if cat == 'FIRST UNIT TEST (25)' else '',
                            'Column2': cat}
                for col in subj_list + result_cols:
                    row_data[col] = ''
                if cat in s['Exams']:
                    row_data.update(s['Exams'][cat])
                elif cat == 'INT/PRACTICAL (20/30)':
                    for sub in subj_list:
                        row_data[sub] = "0"
                rows.append(row_data)

        base_df = pd.DataFrame(rows)
        for col in base_df.columns:
            base_df[col] = base_df[col].astype(str).replace('nan', '')

        # ─── Internal Marks Input Section ──────────────────────────────────────────
        st.markdown("---")
        st.subheader("📝 Enter Internal / Practical Marks (20/30)")
        st.info("Fill in the internal/practical marks for each student below. These will be used to calculate totals, percentage, result and rank.")

        student_rolls = [r for r in sorted(all_students.keys(),
                          key=lambda x: float(x) if x.replace('.', '', 1).isdigit() else 0)]

        if 'internal_marks' not in st.session_state:
            st.session_state.internal_marks = {
                roll: {sub: "0" for sub in subj_list}
                for roll in student_rolls
            }

        int_cols = st.columns([1, 2] + [1]*6)
        int_cols[0].markdown("**Roll No.**")
        int_cols[1].markdown("**Student Name**")
        for i, sub in enumerate(subj_list):
            int_cols[i+2].markdown(f"**{sub}**")

        for roll in student_rolls:
            name = all_students[roll]['Name']
            cols = st.columns([1, 2] + [1]*6)
            cols[0].write(roll)
            cols[1].write(name)
            for i, sub in enumerate(subj_list):
                val = cols[i+2].text_input(
                    label=f"{roll}-{sub}",
                    value=st.session_state.internal_marks[roll][sub],
                    key=f"int_{roll}_{sub}",
                    label_visibility="collapsed"
                )
                st.session_state.internal_marks[roll][sub] = val

        # Inject internal marks into base_df before showing editor
        for i, roll in enumerate(student_rolls):
            # INT row is index 4 in each 7-row block
            block_start = i * 7
            int_row_idx = block_start + 4
            for sub in subj_list:
                base_df.at[int_row_idx, sub] = st.session_state.internal_marks[roll][sub]

        st.markdown("---")
        st.subheader("📊 Marks Data Preview & Edit")
        edited_df = st.data_editor(base_df, hide_index=True, use_container_width=True)

        # ─── Generate Report ────────────────────────────────────────────────────────
        if st.button("🚀 Generate Final Report & Rank"):
            processed = []
            pass_list = []

            for i in range(0, len(edited_df), 7):
                block = edited_df.iloc[i:i+7].copy()
                nums = block.iloc[0:5][subj_list].applymap(clean_marks)
                t200 = nums.sum()
                for sub in subj_list:
                    block.iloc[5, block.columns.get_loc(sub)] = str(int(t200[sub]))
                a100 = t200.apply(lambda x: custom_round(x / 2))
                for sub in subj_list:
                    block.iloc[6, block.columns.get_loc(sub)] = str(a100[sub])

                gt = a100.sum()
                pc = round((gt / 600) * 100, 2)
                isp = all(m >= 35 for m in a100)
                block.iloc[6, block.columns.get_loc('Grand Total')] = str(gt)
                block.iloc[6, block.columns.get_loc('%')] = str(pc)
                block.iloc[6, block.columns.get_loc('Result')] = "PASS" if isp else "FAIL"
                processed.append(block)
                if isp:
                    pass_list.append({'idx': i + 6, 'total': gt})

            final_df = pd.concat(processed).reset_index(drop=True)
            pass_list.sort(key=lambda x: x['total'], reverse=True)
            curr_r, last_t = 0, -1
            for j, e in enumerate(pass_list):
                if e['total'] != last_t:
                    curr_r = j + 1
                last_t = e['total']
                final_df.at[e['idx'], 'Rank'] = str(curr_r)

            st.success("✅ Report Generated! Ranks Applied.")
            st.dataframe(final_df, use_container_width=True)

            # ─── Build Excel with live formulas ────────────────────────────────────
            wb = Workbook()
            ws = wb.active
            ws.title = "Consolidated"

            # Styles
            header_font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
            header_fill = PatternFill("solid", start_color="1F4E79")
            cat_font_map = {
                'FIRST UNIT TEST (25)':     PatternFill("solid", start_color="DDEBF7"),
                'FIRST TERM EXAM (50)':     PatternFill("solid", start_color="E2EFDA"),
                'SECOND UNIT TEST (25)':    PatternFill("solid", start_color="FFF2CC"),
                'ANNUAL EXAM (70/80)':      PatternFill("solid", start_color="FCE4D6"),
                'INT/PRACTICAL (20/30)':    PatternFill("solid", start_color="EAD1DC"),
                'Total Marks Out of 200':   PatternFill("solid", start_color="D9D9D9"),
                'Average Marks 200/2=100':  PatternFill("solid", start_color="BDD7EE"),
            }
            thin = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
            center = Alignment(horizontal='center', vertical='center')

            all_cols = ['Roll No.', 'Column1', 'Column2'] + subj_list + result_cols
            headers = ['Roll No.', 'Student Name', 'Exam Type'] + subj_list + result_cols
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_idx, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = center
                cell.border = thin

            # Column widths
            col_widths = {'Roll No.': 10, 'Column1': 22, 'Column2': 28}
            for sub in subj_list:
                col_widths[sub] = 10
            for rc in result_cols:
                col_widths[rc] = 13
            for col_idx, col_name in enumerate(all_cols, 1):
                ws.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(col_name, 12)

            # Subject columns (D to I = indices 4 to 9), result cols after
            sub_col_start = 4  # column D
            gt_col = sub_col_start + 6          # Grand Total column
            pct_col = gt_col + 1                # % column
            result_col = pct_col + 1            # Result column
            rank_col = result_col + 3           # Rank column (after Remark)

            # Map sub column letters
            sub_letters = [get_column_letter(sub_col_start + i) for i in range(6)]
            gt_letter = get_column_letter(gt_col)
            pct_letter = get_column_letter(pct_col)
            result_letter = get_column_letter(result_col)
            rank_letter = get_column_letter(rank_col)

            # Write data rows
            num_students = len(student_rolls)
            data_row = 2

            # We'll collect avg row addresses for RANK formula later
            avg_gt_cells = []  # list of grand total cell addresses for avg rows

            for s_idx, roll in enumerate(student_rolls):
                block_start_row = data_row + s_idx * 7
                name = all_students[roll]['Name']

                for cat_idx, cat in enumerate(categories):
                    excel_row = block_start_row + cat_idx
                    row_vals = final_df.iloc[s_idx * 7 + cat_idx]

                    # Roll No and Name only on first row of block
                    ws.cell(row=excel_row, column=1, value=roll if cat_idx == 0 else "")
                    ws.cell(row=excel_row, column=2, value=name if cat_idx == 0 else "")
                    ws.cell(row=excel_row, column=3, value=cat)

                    fill = cat_font_map.get(cat, PatternFill("solid", start_color="FFFFFF"))

                    if cat == 'Total Marks Out of 200':
                        # SUM of rows above (rows 1-5 of this block)
                        for i, sl in enumerate(sub_letters):
                            r1 = block_start_row
                            r5 = block_start_row + 4
                            cell = ws.cell(row=excel_row, column=sub_col_start + i,
                                           value=f"=SUM({sl}{r1}:{sl}{r5})")
                            cell.fill = fill
                            cell.border = thin
                            cell.alignment = center
                            cell.font = Font(name="Arial", bold=True)
                        # No Grand Total / % / Result for this row
                        for rc_i in range(len(result_cols)):
                            c = ws.cell(row=excel_row, column=gt_col + rc_i, value="")
                            c.fill = fill; c.border = thin

                    elif cat == 'Average Marks 200/2=100':
                        tot_row = excel_row - 1  # Total row just above
                        for i, sl in enumerate(sub_letters):
                            cell = ws.cell(row=excel_row, column=sub_col_start + i,
                                           value=f"=ROUND({sl}{tot_row}/2,0)")
                            cell.fill = fill; cell.border = thin; cell.alignment = center
                            cell.font = Font(name="Arial", bold=True)

                        # Grand Total = SUM of averaged subjects
                        gt_cell = ws.cell(row=excel_row, column=gt_col,
                                          value=f"=SUM({sub_letters[0]}{excel_row}:{sub_letters[-1]}{excel_row})")
                        gt_cell.fill = fill; gt_cell.border = thin; gt_cell.alignment = center
                        gt_cell.font = Font(name="Arial", bold=True, color="1F4E79")

                        # %
                        pct_cell = ws.cell(row=excel_row, column=pct_col,
                                           value=f"=ROUND({gt_letter}{excel_row}/600*100,2)")
                        pct_cell.fill = fill; pct_cell.border = thin; pct_cell.alignment = center

                        # Result: PASS if all sub averages >= 35
                        pass_checks = ",".join([f'{sl}{excel_row}>=35' for sl in sub_letters])
                        result_cell = ws.cell(row=excel_row, column=result_col,
                                              value=f'=IF(AND({pass_checks}),"PASS","FAIL")')
                        result_cell.fill = fill; result_cell.border = thin; result_cell.alignment = center
                        result_cell.font = Font(name="Arial", bold=True)

                        # Remark blank
                        ws.cell(row=excel_row, column=result_col + 1, value="").border = thin
                        ws.cell(row=excel_row, column=result_col + 2, value="").border = thin

                        # Rank placeholder — will fill after all rows written
                        avg_gt_cells.append((excel_row, f"{gt_letter}{excel_row}"))

                        # Style rank cell
                        rank_c = ws.cell(row=excel_row, column=rank_col, value="")
                        rank_c.fill = fill; rank_c.border = thin; rank_c.alignment = center
                        rank_c.font = Font(name="Arial", bold=True, color="C00000")

                    else:
                        # Regular exam/internal row — write values from final_df
                        for i, sub in enumerate(subj_list):
                            v = row_vals.get(sub, "")
                            try:
                                v = float(v)
                            except:
                                pass
                            cell = ws.cell(row=excel_row, column=sub_col_start + i, value=v)
                            cell.fill = fill; cell.border = thin; cell.alignment = center

                        # Write result columns from final_df (non-formula rows)
                        for rc_i, rc in enumerate(result_cols):
                            v = row_vals.get(rc, "")
                            if rc == 'Rank':
                                v = ""  # blank for non-avg rows
                            c = ws.cell(row=excel_row, column=gt_col + rc_i, value=v)
                            c.fill = fill; c.border = thin; c.alignment = center

                    # Style Roll/Name/Category cells
                    for col_i in [1, 2, 3]:
                        c = ws.cell(row=excel_row, column=col_i)
                        c.fill = fill; c.border = thin
                        if col_i == 2:
                            c.font = Font(name="Arial", bold=(cat_idx == 0))
                        else:
                            c.font = Font(name="Arial")

            # ── RANK formulas using RANK.EQ across all avg Grand Total cells ──
            if avg_gt_cells:
                gt_range = ",".join([addr for _, addr in avg_gt_cells])
                # Build a named range string for RANK — use IF+LARGE approach for proper ranking
                # Simpler: use RANK.EQ against the list of gt cells as an array
                # We build a helper column approach using RANK.EQ with an array const
                gt_addresses = [addr for _, addr in avg_gt_cells]
                arr_str = "(" + ",".join(gt_addresses) + ")"

                for (avg_row, gt_addr) in avg_gt_cells:
                    # Only rank if PASS
                    rank_formula = (
                        f'=IF({result_letter}{avg_row}="PASS",'
                        f'RANK.EQ({gt_addr},{{{",".join(gt_addresses)}}},0),"")'
                    )
                    ws.cell(row=avg_row, column=rank_col, value=rank_formula)

            # Freeze header row
            ws.freeze_panes = "A2"

            output = BytesIO()
            wb.save(output)
            output.seek(0)

            st.download_button(
                "📥 Download Excel (with Live Formulas)",
                output.getvalue(),
                "Final_Consolidated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
        import traceback
        st.code(traceback.format_exc())
