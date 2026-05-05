import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
import re
import warnings
warnings.filterwarnings('ignore')

# Page config
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
        if 'STUDENT NAME' in col_str:
            col_info['student_name'] = idx
        elif 'ROLL' in col_str and 'NO' in col_str:
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
    
    try:
        xl = pd.ExcelFile(uploaded_file)
        st.info(f"📁 Found {len(xl.sheet_names)} sheets in the file")
        
        # Get test sheets (exclude summary sheets)
        test_sheets = []
        for s in xl.sheet_names:
            s_upper = s.upper()
            if "SUMMARY" not in s_upper and "ANALYSIS" not in s_upper and "SHEET20" not in s_upper:
                test_sheets.append(s)
        
        st.info(f"📊 Found {len(test_sheets)} test sheets")
        
        if len(test_sheets) == 0:
            st.error("No test sheets found! Make sure your sheets don't contain 'Summary' in the name.")
            return None, None
        
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
                # Try reading with skiprows=7 (your original format)
                df = pd.read_excel(xl, sheet_name=sheet, skiprows=7)
                
                # Check if we have the expected columns
                if len(df.columns) >= 10:
                    # Use column indices based on your format
                    name_col = df.columns[1]
                    roll_col = df.columns[2]
                    phy_col = df.columns[3]
                    chem_col = df.columns[5]
                    math_col = df.columns[7]
                    total_col = df.columns[9]
                    
                    students_loaded = 0
                    
                    for _, row in df.iterrows():
                        # Skip empty rows
                        if pd.isna(row[name_col]) and pd.isna(row[roll_col]):
                            continue
                        
                        student_name = normalize_name(row[name_col])
                        if not student_name:
                            continue
                        
                        roll_no = pd.to_numeric(row[roll_col], errors='coerce')
                        total = pd.to_numeric(row[total_col], errors='coerce')
                        
                        if pd.isna(total) or total < 0:
                            continue
                        
                        if student_name not in all_student_data:
                            all_student_data[student_name] = {
                                "name": student_name, "roll_numbers": set(), "tests": {}
                            }
                        
                        if pd.notna(roll_no):
                            all_student_data[student_name]["roll_numbers"].add(int(roll_no))
                        
                        phy = pd.to_numeric(row[phy_col], errors='coerce') if phy_col else 0
                        chem = pd.to_numeric(row[chem_col], errors='coerce') if chem_col else 0
                        math_score = pd.to_numeric(row[math_col], errors='coerce') if math_col else 0
                        
                        all_student_data[student_name]["tests"][sheet] = {
                            "phy": phy if pd.notna(phy) else 0,
                            "chem": chem if pd.notna(chem) else 0,
                            "math": math_score if pd.notna(math_score) else 0,
                            "total": total,
                            "type": test_type
                        }
                        students_loaded += 1
                    
                    if students_loaded > 0:
                        st.success(f"  ✓ {sheet}: Loaded {students_loaded} students")
                    else:
                        st.warning(f"  ⚠️ {sheet}: No students loaded")
                else:
                    st.warning(f"  ⚠️ {sheet}: Not enough columns ({len(df.columns)})")
                    
            except Exception as e:
                st.warning(f"  ⚠️ {sheet}: Error - {str(e)[:50]}")
                continue
        
        if len(all_student_data) == 0:
            st.error("No student data found in any sheet!")
            return None, None
            
        return all_student_data, test_metadata
        
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return None, None

# File upload
uploaded_file = st.file_uploader("📁 Upload Master Sheet Excel File", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("🔄 Loading student data..."):
        all_student_data, test_metadata = load_excel_data(uploaded_file)
    
    if all_student_data and len(all_student_data) > 0:
        st.success(f"✅ Successfully loaded {len(all_student_data)} students and {len(test_metadata)} tests!")
        
        # Show sample of first few students
        with st.expander("📋 View loaded students"):
            st.write(f"Total students: {len(all_student_data)}")
            st.write("Sample students:", list(all_student_data.keys())[:10])
        
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
                rank = sum(score > marks["total"] for score in all_scores) + 1 if all_scores else 1
                
                result = {
                    "Test Name": sheet[:35],
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
            
            if total_tests > 0:
                all_pcts = [float(r["%"].replace("%", "")) for r in btest_results + brtest_results]
                avg_pct = round(np.mean(all_pcts), 1)
                all_ranks = [r["Rank"] for r in btest_results + brtest_results]
                best_rank = min(all_ranks)
                roll_display = ", ".join([str(r) for r in sorted(student["roll_numbers"])][:2])
            else:
                avg_pct = 0
                best_rank = "N/A"
                roll_display = "N/A"
            
            with col1:
                st.metric("📚 Tests Attempted", f"{total_tests}/{len(test_metadata)}")
            with col2:
                st.metric("📊 Average Score", f"{avg_pct}%")
            with col3:
                st.metric("🏆 Best Rank", best_rank)
            with col4:
                st.metric("🆔 Roll Number", roll_display)
            
            st.markdown("---")
            
            if btest_results:
                st.subheader("📘 BTEST/GRAND Tests (JEE Format - 300 marks)")
                st.dataframe(pd.DataFrame(btest_results), use_container_width=True)
            
            if brtest_results:
                st.subheader("📘 BRTEST Tests (CET Format - 200 marks)")
                st.dataframe(pd.DataFrame(brtest_results), use_container_width=True)
            
            if total_tests == 0:
                st.warning("No test data found for this student. They might be absent in all tests.")
            else:
                st.markdown("---")
                st.subheader("📊 Subject Performance Trend")
                
                all_names = [t['Test Name'][:20] for t in btest_results + brtest_results]
                all_physics = [t['Physics'] for t in btest_results + brtest_results]
                all_chemistry = [t['Chemistry'] for t in btest_results + brtest_results]
                all_maths = [t['Maths'] for t in btest_results + brtest_results]
                
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=all_names, y=all_physics, mode='lines+markers', name='Physics'))
                fig.add_trace(go.Scatter(x=all_names, y=all_chemistry, mode='lines+markers', name='Chemistry'))
                fig.add_trace(go.Scatter(x=all_names, y=all_maths, mode='lines+markers', name='Mathematics'))
                fig.update_layout(title="Subject-wise Performance", height=400)
                st.plotly_chart(fig, use_container_width=True)
                
                # Overall percentage trend
                fig_trend = go.Figure()
                fig_trend.add_trace(go.Scatter(x=all_names, y=all_pcts, mode='lines+markers', name='Percentage'))
                fig_trend.add_hline(y=75, line_dash="dash", line_color="green", annotation_text="75% Target")
                fig_trend.update_layout(title="Overall Performance Trend", height=400)
                st.plotly_chart(fig_trend, use_container_width=True)
    else:
        st.error("❌ No student data found in the file. Please check the file format.")
        st.info("""
        **Expected Format:**
        - Each test sheet should have columns: Student Name, Roll No, PHY, CHEM, MATHS, TOTAL
        - Data should start from row 8 (after headers)
        - Sheet names shouldn't contain 'Summary' or 'Analysis'
        """)
else:
    st.info("👈 **Upload the Master Sheet Excel file to get started!**")
    
    with st.expander("📖 Instructions"):
        st.markdown("""
        1. Click 'Browse files' to upload your Master Sheet Excel file
        2. Wait for the data to load
        3. Select a student from the dropdown
        4. View their performance across all tests
        
        **Supported Test Types:**
        - BTEST/GRAND: 300 total marks (100 each subject)
        - BRTEST: 200 total marks (50 PHY, 50 CHEM, 100 MATHS)
        """)
