import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
import re
from plotly.subplots import make_subplots

# Page config for mobile
st.set_page_config(
    page_title="Student Performance Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Student Performance Dashboard")
st.markdown("---")

# Helper functions
def detect_test_type(sheet_name):
    if sheet_name.upper().startswith("BRTEST"):
        return "BRTEST"
    return "BTEST"

def normalize_name(name):
    if pd.isna(name):
        return None
    name = str(name).upper().strip()
    name = re.sub(r'\s+', ' ', name)
    name = name.replace('KILKARNI', 'KULKARNI')
    if len(name) < 3:
        return None
    return name

def find_data_columns(df_columns):
    col_info = {
        'student_name': None, 'roll_no': None,
        'phy': None, 'phy_rank': None,
        'chem': None, 'chem_rank': None,
        'math': None, 'math_rank': None, 'total': None
    }
    for idx, col in enumerate(df_columns):
        col_str = str(col).upper().strip()
        if col_str.startswith('UNNAMED'):
            continue
        if 'STUDENT NAME' in col_str:
            col_info['student_name'] = idx
        elif 'ROLL' in col_str:
            col_info['roll_no'] = idx
        elif col_str == 'PHY':
            col_info['phy'] = idx
        elif 'PHY RANK' in col_str:
            col_info['phy_rank'] = idx
        elif col_str == 'CHEM':
            col_info['chem'] = idx
        elif 'CHE RANK' in col_str:
            col_info['chem_rank'] = idx
        elif 'MATH' in col_str and 'RANK' not in col_str:
            col_info['math'] = idx
        elif 'MATH RANK' in col_str:
            col_info['math_rank'] = idx
        elif col_str == 'TOTAL':
            col_info['total'] = idx
    return col_info

@st.cache_data
def load_excel_data(uploaded_file):
    all_student_data = {}
    test_metadata = {}
    xl = pd.ExcelFile(uploaded_file)
    test_sheets = [s for s in xl.sheet_names if "Summary" not in s and "Analysis" not in s and "Sheet20" not in s]
    
    for idx, sheet in enumerate(test_sheets, start=1):
        test_type = detect_test_type(sheet)
        if test_type == "BRTEST":
            max_phy, max_chem, max_math = 50, 50, 100
            max_total = 200
        else:
            max_phy, max_chem, max_math = 100, 100, 100
            max_total = 300

        test_metadata[sheet] = {
            "type": test_type, "index": idx,
            "max_phy": max_phy, "max_chem": max_chem,
            "max_math": max_math, "max_total": max_total
        }

        try:
            df_raw = pd.read_excel(xl, sheet_name=sheet, header=None)
            header_row = None
            for i in range(min(30, len(df_raw))):
                row_text = " ".join(df_raw.iloc[i].astype(str).str.upper().tolist())
                if "STUDENT NAME" in row_text and "ROLL" in row_text:
                    header_row = i
                    break
            if header_row is None:
                continue
            
            df = pd.read_excel(xl, sheet_name=sheet, skiprows=header_row)
            df.columns = [str(c).strip() for c in df.columns]
            col_info = find_data_columns(df.columns)
            
            required = ['student_name', 'roll_no', 'total']
            if any(col_info[c] is None for c in required):
                continue
            
            for _, row in df.iterrows():
                row_text = " ".join([str(x).upper() for x in row.values if pd.notna(x)])
                if "NOT IN RANK LIST" in row_text:
                    break
                
                student_name = normalize_name(row.iloc[col_info['student_name']])
                if not student_name:
                    continue
                
                roll_no = pd.to_numeric(row.iloc[col_info['roll_no']], errors='coerce')
                total = pd.to_numeric(row.iloc[col_info['total']], errors='coerce')
                if pd.isna(total) or total < 0:
                    continue
                
                if student_name not in all_student_data:
                    all_student_data[student_name] = {"name": student_name, "roll_numbers": set(), "tests": {}}
                
                if pd.notna(roll_no):
                    all_student_data[student_name]["roll_numbers"].add(int(roll_no))
                
                phy = pd.to_numeric(row.iloc[col_info['phy']], errors='coerce') if col_info['phy'] is not None else 0
                chem = pd.to_numeric(row.iloc[col_info['chem']], errors='coerce') if col_info['chem'] is not None else 0
                math_score = pd.to_numeric(row.iloc[col_info['math']], errors='coerce') if col_info['math'] is not None else 0
                
                all_student_data[student_name]["tests"][sheet] = {
                    "phy": phy, "chem": chem, "math": math_score,
                    "total": total, "type": test_type
                }
        except Exception as e:
            continue
    
    return all_student_data, test_metadata

# File upload
uploaded_file = st.file_uploader("📁 Upload Master Sheet Excel File", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Loading data..."):
        all_student_data, test_metadata = load_excel_data(uploaded_file)
    
    if all_student_data:
        st.success(f"✅ Loaded {len(all_student_data)} students and {len(test_metadata)} tests!")
        
        student_options = sorted(all_student_data.keys())
        selected_student = st.selectbox("🎓 Select Student", student_options)
        
        if selected_student:
            student = all_student_data[selected_student]
            
            btest_results = []
            brtest_results = []
            
            for sheet, meta in test_metadata.items():
                if sheet not in student["tests"]:
                    continue
                marks = student["tests"][sheet]
                pct = round((marks["total"] / meta["max_total"]) * 100, 1)
                
                # Calculate rank
                all_scores = []
                for s_data in all_student_data.values():
                    if sheet in s_data["tests"]:
                        all_scores.append(s_data["tests"][sheet]["total"])
                rank = sum(score > marks["total"] for score in all_scores) + 1
                
                result = {
                    "Test Name": sheet[:30],
                    "Physics": marks["phy"],
                    "Chemistry": marks["chem"],
                    "Maths": marks["math"],
                    "Total": f"{marks['total']:.0f}/{meta['max_total']}",
                    "%": f"{pct}%",
                    "Rank": rank
                }
                
                if meta["type"] == "BTEST":
                    btest_results.append(result)
                else:
                    brtest_results.append(result)
            
            # Metrics
            col1, col2, col3, col4 = st.columns(4)
            total_tests = len(btest_results) + len(brtest_results)
            all_pcts = [float(r["%"].replace("%", "")) for r in btest_results + brtest_results]
            avg_pct = round(np.mean(all_pcts), 1) if all_pcts else 0
            all_ranks = [r["Rank"] for r in btest_results + brtest_results]
            best_rank = min(all_ranks) if all_ranks else 'N/A'
            
            with col1:
                st.metric("📚 Tests", f"{total_tests}/{len(test_metadata)}")
            with col2:
                st.metric("📊 Avg %", f"{avg_pct}%")
            with col3:
                st.metric("🏆 Best Rank", best_rank)
            with col4:
                roll_display = ", ".join([str(r) for r in sorted(student["roll_numbers"])][:2])
                st.metric("🆔 Roll No", roll_display)
            
            st.markdown("---")
            
            if btest_results:
                st.subheader("📘 BTEST/GRAND Tests (300 marks)")
                st.dataframe(pd.DataFrame(btest_results), use_container_width=True)
            
            if brtest_results:
                st.subheader("📘 BRTEST Tests (200 marks)")
                st.dataframe(pd.DataFrame(brtest_results), use_container_width=True)
            
            st.markdown("---")
            
            # Subject graphs
            if btest_results:
                st.subheader("📊 Subject Performance - BTEST")
                btest_names = [t['Test Name'][:20] for t in btest_results]
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=btest_names, y=[t['Physics'] for t in btest_results], mode='lines+markers', name='Physics'))
                fig.add_trace(go.Scatter(x=btest_names, y=[t['Chemistry'] for t in btest_results], mode='lines+markers', name='Chemistry'))
                fig.add_trace(go.Scatter(x=btest_names, y=[t['Maths'] for t in btest_results], mode='lines+markers', name='Maths'))
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)
            
            # Overall trend
            all_names = [t['Test Name'][:20] for t in btest_results + brtest_results]
            fig_trend = go.Figure()
            fig_trend.add_trace(go.Scatter(x=all_names, y=all_pcts, mode='lines+markers'))
            fig_trend.add_hline(y=75, line_dash="dash", line_color="green")
            fig_trend.update_layout(title="Overall Performance Trend", height=400)
            st.plotly_chart(fig_trend, use_container_width=True)
    else:
        st.error("No student data found")
else:
    st.info("👈 Upload Master Sheet Excel file to get started!")
