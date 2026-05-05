import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
import re
from collections import Counter
import warnings
warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(
    page_title="Student Performance Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Student Performance Dashboard")
st.markdown("---")

# ============================================================
# FIXED FILE PATH - Put your Excel file in the same directory
# ============================================================
# Option A: If file is in the same folder as app.py
EXCEL_FILE_PATH = "Master Sheet.xlsx"

# Option B: If using GitHub, you can use raw URL
# EXCEL_FILE_PATH = "https://raw.githubusercontent.com/PRIYANKSAXENA555/student-dashboard/main/Master Sheet.xlsx"

import os
if not os.path.exists(EXCEL_FILE_PATH):
    st.error(f"❌ Excel file not found at: {EXCEL_FILE_PATH}")
    st.info("Please upload the Master Sheet.xlsx file to the same directory as this app.")
    st.stop()

# ============================================================
# HELPER FUNCTIONS
# ============================================================
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

def get_weakest_subject(phy_rank, chem_rank, math_rank):
    if pd.isna(phy_rank) or pd.isna(chem_rank) or pd.isna(math_rank):
        return "Absent"
    p, c, m = float(phy_rank), float(chem_rank), float(math_rank)
    if max(p, c, m) - min(p, c, m) <= 30:
        return "Balanced"
    ranks = {'Physics': p, 'Chemistry': c, 'Maths': m}
    weakest = max(ranks, key=ranks.get)
    other_avg = (sum(ranks.values()) - ranks[weakest]) / 2
    if ranks[weakest] > 1.5 * other_avg:
        return weakest
    return "Balanced"

@st.cache_data
def load_excel_data():
    all_student_data = {}
    test_metadata = {}
    
    try:
        xl = pd.ExcelFile(EXCEL_FILE_PATH)
        
        # Get test sheets
        test_sheets = []
        for s in xl.sheet_names:
            s_upper = s.upper()
            if "SUMMARY" not in s_upper and "ANALYSIS" not in s_upper and "SHEET20" not in s_upper:
                test_sheets.append(s)
        
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
                # Read with skiprows=7
                df = pd.read_excel(xl, sheet_name=sheet, skiprows=7)
                
                if len(df.columns) >= 10:
                    name_col = df.columns[1]
                    roll_col = df.columns[2]
                    phy_col = df.columns[3]
                    chem_col = df.columns[5]
                    math_col = df.columns[7]
                    total_col = df.columns[9]
                    phy_rank_col = df.columns[4]
                    chem_rank_col = df.columns[6]
                    math_rank_col = df.columns[8]
                    
                    for _, row in df.iterrows():
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
                        
                        phy = pd.to_numeric(row[phy_col], errors='coerce')
                        chem = pd.to_numeric(row[chem_col], errors='coerce')
                        math_score = pd.to_numeric(row[math_col], errors='coerce')
                        phy_rank = pd.to_numeric(row[phy_rank_col], errors='coerce')
                        chem_rank = pd.to_numeric(row[chem_rank_col], errors='coerce')
                        math_rank = pd.to_numeric(row[math_rank_col], errors='coerce')
                        
                        all_student_data[student_name]["tests"][sheet] = {
                            "phy": phy if pd.notna(phy) else 0,
                            "chem": chem if pd.notna(chem) else 0,
                            "math": math_score if pd.notna(math_score) else 0,
                            "phy_rank": phy_rank if pd.notna(phy_rank) else None,
                            "chem_rank": chem_rank if pd.notna(chem_rank) else None,
                            "math_rank": math_rank if pd.notna(math_rank) else None,
                            "total": total, "type": test_type
                        }
                        
            except Exception as e:
                continue
        
        return all_student_data, test_metadata
        
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return None, None

# ============================================================
# LOAD DATA
# ============================================================
with st.spinner("🔄 Loading student data..."):
    all_student_data, test_metadata = load_excel_data()

if all_student_data and len(all_student_data) > 0:
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
            
            # Calculate weakest subject
            weakest = get_weakest_subject(marks.get("phy_rank"), marks.get("chem_rank"), marks.get("math_rank"))
            
            # Calculate percentages
            phy_pct = round((marks["phy"] / meta["max_phy"]) * 100, 1) if marks["phy"] > 0 else 0
            chem_pct = round((marks["chem"] / meta["max_chem"]) * 100, 1) if marks["chem"] > 0 else 0
            math_pct = round((marks["math"] / meta["max_math"]) * 100, 1) if marks["math"] > 0 else 0
            
            result = {
                "S.No.": meta["index"],
                "Test Name": sheet,
                "Type": meta["type"],
                "Physics": marks["phy"],
                "Physics %": phy_pct,
                "Chemistry": marks["chem"],
                "Chemistry %": chem_pct,
                "Maths": marks["math"],
                "Maths %": math_pct,
                "Phy Rank": marks.get("phy_rank") if marks.get("phy_rank") is not None else '-',
                "Chem Rank": marks.get("chem_rank") if marks.get("chem_rank") is not None else '-',
                "Math Rank": marks.get("math_rank") if marks.get("math_rank") is not None else '-',
                "Total": f"{marks['total']:.0f}/{meta['max_total']}",
                "%": f"{pct}%",
                "Overall Rank": rank,
                "Weakest Subject": weakest
            }
            
            if meta["type"] == "BTEST":
                btest_results.append(result)
            else:
                brtest_results.append(result)
        
        btest_results.sort(key=lambda x: x["S.No."])
        brtest_results.sort(key=lambda x: x["S.No."])
        
        if not btest_results and not brtest_results:
            st.warning("No tests found for this student")
        else:
            all_tests = btest_results + brtest_results
            avg_pct = round(np.mean([float(t["%"].replace("%", "")) for t in all_tests]), 1)
            ranks = [t["Overall Rank"] for t in all_tests]
            best_rank = min(ranks) if ranks else 'N/A'
            roll_numbers_display = ", ".join([str(r) for r in sorted(student["roll_numbers"])])
            
            # Weak subjects summary
            weak_subjects = [t["Weakest Subject"] for t in all_tests if t["Weakest Subject"] not in ["Balanced", "Absent"]]
            weak_count = Counter(weak_subjects)
            
            # Header with metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("📚 Tests Attempted", f"{len(all_tests)}/{len(test_metadata)}")
            with col2:
                st.metric("📊 Average Score", f"{avg_pct}%")
            with col3:
                st.metric("🏆 Best Rank", best_rank)
            with col4:
                st.metric("🆔 Roll Number", roll_numbers_display)
            
            # Weak subjects summary
            if weak_count:
                st.markdown("---")
                st.subheader("⚠️ Weak Subject Summary")
                weak_cols = st.columns(len(weak_count))
                for i, (subject, count) in enumerate(weak_count.items()):
                    with weak_cols[i]:
                        st.metric(f"⚠️ {subject}", f"Weak in {count} test(s)")
            
            st.markdown("---")
            
            # BTEST Results Table
            if btest_results:
                st.subheader("📘 BTEST/GRAND TESTS (JEE Format - 300 marks)")
                display_cols = ['Test Name', 'Physics', 'Chemistry', 'Maths', 'Phy Rank', 'Chem Rank', 'Math Rank', 
                               'Total', '%', 'Overall Rank', 'Weakest Subject']
                btest_df = pd.DataFrame(btest_results)
                st.dataframe(btest_df[display_cols], use_container_width=True)
            
            # BRTEST Results Table
            if brtest_results:
                st.subheader("📘 BRTEST TESTS (CET Format - 200 marks)")
                display_cols = ['Test Name', 'Physics', 'Chemistry', 'Maths', 'Phy Rank', 'Chem Rank', 'Math Rank',
                               'Total', '%', 'Overall Rank', 'Weakest Subject']
                brtest_df = pd.DataFrame(brtest_results)
                st.dataframe(brtest_df[display_cols], use_container_width=True)
            
            st.markdown("---")
            
            # ====================================================
            # SUBJECT-WISE GRAPHS FOR BTEST
            # ====================================================
            if btest_results:
                st.subheader("📊 Subject-wise Performance - BTEST Tests")
                
                btest_names = [t['Test Name'][:25] for t in btest_results]
                btest_phy = [t['Physics'] for t in btest_results]
                btest_chem = [t['Chemistry'] for t in btest_results]
                btest_math = [t['Maths'] for t in btest_results]
                
                # Combined comparison
                fig_combined = go.Figure()
                fig_combined.add_trace(go.Scatter(x=btest_names, y=btest_phy, mode='lines+markers', 
                                                   name='Physics', line=dict(color='#3498DB', width=2)))
                fig_combined.add_trace(go.Scatter(x=btest_names, y=btest_chem, mode='lines+markers', 
                                                   name='Chemistry', line=dict(color='#9B59B6', width=2)))
                fig_combined.add_trace(go.Scatter(x=btest_names, y=btest_math, mode='lines+markers', 
                                                   name='Mathematics', line=dict(color='#F1C40F', width=2)))
                fig_combined.add_hline(y=75, line_dash="dash", line_color="green", annotation_text="75% Target")
                fig_combined.update_layout(title="Subject Comparison (BTEST)", height=450)
                st.plotly_chart(fig_combined, use_container_width=True)
            
            # ====================================================
            # SUBJECT-WISE GRAPHS FOR BRTEST
            # ====================================================
            if brtest_results:
                st.subheader("📊 Subject-wise Performance - BRTEST Tests (CET Format)")
                
                brtest_names = [t['Test Name'][:25] for t in brtest_results]
                brtest_phy = [t['Physics'] for t in brtest_results]
                brtest_chem = [t['Chemistry'] for t in brtest_results]
                brtest_math = [t['Maths'] for t in brtest_results]
                
                fig_combined = go.Figure()
                fig_combined.add_trace(go.Scatter(x=brtest_names, y=brtest_phy, mode='lines+markers', 
                                                   name='Physics (max 50)', line=dict(color='#3498DB', width=2)))
                fig_combined.add_trace(go.Scatter(x=brtest_names, y=brtest_chem, mode='lines+markers', 
                                                   name='Chemistry (max 50)', line=dict(color='#9B59B6', width=2)))
                fig_combined.add_trace(go.Scatter(x=brtest_names, y=brtest_math, mode='lines+markers', 
                                                   name='Mathematics (max 100)', line=dict(color='#F1C40F', width=2)))
                fig_combined.update_layout(title="Subject Comparison (BRTEST - CET Format)", height=450)
                st.plotly_chart(fig_combined, use_container_width=True)
            
            # ====================================================
            # OVERALL PERFORMANCE TREND
            # ====================================================
            st.subheader("📈 Overall Performance Trend")
            all_names = [t['Test Name'][:25] for t in all_tests]
            all_pcts = [float(t["%"].replace("%", "")) for t in all_tests]
            
            fig_trend = go.Figure()
            
            if btest_results:
                fig_trend.add_trace(go.Scatter(
                    x=[t['Test Name'][:25] for t in btest_results],
                    y=[float(t["%"].replace("%", "")) for t in btest_results],
                    mode='lines+markers', name='BTEST (300 marks)',
                    line=dict(color='#3498DB', width=3), marker=dict(size=10)
                ))
            
            if brtest_results:
                fig_trend.add_trace(go.Scatter(
                    x=[t['Test Name'][:25] for t in brtest_results],
                    y=[float(t["%"].replace("%", "")) for t in brtest_results],
                    mode='lines+markers', name='BRTEST (200 marks)',
                    line=dict(color='#E67E22', width=3), marker=dict(size=10, symbol='diamond')
                ))
            
            fig_trend.add_hline(y=75, line_dash="dash", line_color="green", annotation_text="Target (75%)")
            fig_trend.update_layout(title="Percentage Score Across All Tests", height=450)
            st.plotly_chart(fig_trend, use_container_width=True)
            
            # ====================================================
            # RANK TREND CHART
            # ====================================================
            st.subheader("🏆 Rank Trend (Lower is Better)")
            
            fig_rank = go.Figure()
            
            if btest_results:
                fig_rank.add_trace(go.Scatter(
                    x=[t['Test Name'][:25] for t in btest_results],
                    y=[t['Overall Rank'] for t in btest_results],
                    mode='lines+markers', name='BTEST Rank',
                    line=dict(color='#3498DB', width=3)
                ))
            
            if brtest_results:
                fig_rank.add_trace(go.Scatter(
                    x=[t['Test Name'][:25] for t in brtest_results],
                    y=[t['Overall Rank'] for t in brtest_results],
                    mode='lines+markers', name='BRTEST Rank',
                    line=dict(color='#E67E22', width=3)
                ))
            
            fig_rank.update_layout(title="Rank Performance", yaxis=dict(autorange="reversed"), height=450)
            st.plotly_chart(fig_rank, use_container_width=True)
            
            # ====================================================
            # WEAKEST SUBJECT INSIGHTS
            # ====================================================
            st.subheader("📊 Weakest Subject Analysis")
            
            subject_weakness = {'Physics': 0, 'Chemistry': 0, 'Maths': 0}
            total_weak = len(weak_subjects)
            
            for ws in weak_subjects:
                if ws in subject_weakness:
                    subject_weakness[ws] += 1
            
            if total_weak > 0:
                st.write(f"**Subject-wise weakness breakdown (out of {total_weak} weak instances):**")
                for subject, count in sorted(subject_weakness.items(), key=lambda x: x[1], reverse=True):
                    pct = (count / total_weak) * 100
                    st.progress(int(pct))
                    st.write(f"{subject}: {count} times ({pct:.0f}%)")
            
            weak_tests = [t for t in all_tests if t["Weakest Subject"] not in ["Balanced", "Absent"]]
            if weak_tests:
                st.write(f"**⚠️ Tests where {selected_student} showed weakness:**")
                for t in weak_tests:
                    st.write(f"• {t['Test Name'][:50]} → Weak in: {t['Weakest Subject']}")
            
            balanced_tests = [t for t in all_tests if t["Weakest Subject"] == "Balanced"]
            if balanced_tests:
                st.write(f"**✅ Tests with balanced performance ({len(balanced_tests)} tests):**")
                for t in balanced_tests[:5]:
                    st.write(f"• {t['Test Name'][:50]}")
                if len(balanced_tests) > 5:
                    st.write(f"... and {len(balanced_tests) - 5} more")
            
            # Missing tests
            attempted = set([t['Test Name'] for t in all_tests])
            missing = [t for t in test_metadata.keys() if t not in attempted]
            if missing:
                st.write(f"**⚠️ ABSENT/NO DATA for {len(missing)} tests:**")
                for m in missing[:10]:
                    st.write(f"• {m} ({test_metadata[m]['type']})")
            
            st.markdown("---")
            st.caption("✅ Dashboard Complete | Data Source: Master Sheet Excel")
else:
    st.error("❌ No student data found in the file.")
