import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import re
import io

# Page configuration for mobile
st.set_page_config(
    page_title="Student Performance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for better mobile view
st.markdown("""
<style>
    /* Mobile responsive */
    @media (max-width: 768px) {
        .stPlotlyChart {
            width: 100% !important;
        }
        .stDataFrame {
            font-size: 10px;
        }
        .stButton button {
            width: 100%;
        }
    }
    /* Card style for metrics */
    .metric-card {
        background: linear-gradient(135deg, #1a1a2e, #16213e);
        padding: 15px;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    .success-badge {
        background-color: #28a745;
        padding: 5px 10px;
        border-radius: 20px;
        display: inline-block;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# Helper Functions
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
def load_excel_data(uploaded_file):
    """Load and process the Excel file"""
    all_student_data = {}
    test_metadata = {}
    
    xl = pd.ExcelFile(uploaded_file)
    
    # Get test sheets
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
            
            # Find header row
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
                    all_student_data[student_name] = {
                        "name": student_name, "roll_numbers": set(), "tests": {}
                    }
                
                if pd.notna(roll_no):
                    all_student_data[student_name]["roll_numbers"].add(int(roll_no))
                
                phy = pd.to_numeric(row.iloc[col_info['phy']], errors='coerce') if col_info['phy'] is not None else 0
                chem = pd.to_numeric(row.iloc[col_info['chem']], errors='coerce') if col_info['chem'] is not None else 0
                math_score = pd.to_numeric(row.iloc[col_info['math']], errors='coerce') if col_info['math'] is not None else 0
                phy_rank = pd.to_numeric(row.iloc[col_info['phy_rank']], errors='coerce') if col_info['phy_rank'] is not None else None
                chem_rank = pd.to_numeric(row.iloc[col_info['chem_rank']], errors='coerce') if col_info['chem_rank'] is not None else None
                math_rank = pd.to_numeric(row.iloc[col_info['math_rank']], errors='coerce') if col_info['math_rank'] is not None else None
                
                all_student_data[student_name]["tests"][sheet] = {
                    "phy": phy, "chem": chem, "math": math_score,
                    "phy_rank": phy_rank, "chem_rank": chem_rank, "math_rank": math_rank,
                    "total": total, "type": test_type
                }
                
        except Exception as e:
            continue
    
    return all_student_data, test_metadata

# ============================================================
# Main App
# ============================================================
st.title("📊 Student Performance Dashboard")
st.markdown("---")

# File upload
uploaded_file = st.file_uploader(
    "📁 Upload Master Sheet Excel File",
    type=["xlsx"],
    help="Upload the Excel file containing all test results"
)

if uploaded_file is not None:
    with st.spinner("🔄 Loading student data... This may take a few seconds."):
        all_student_data, test_metadata = load_excel_data(uploaded_file)
    
    if all_student_data:
        st.success(f"✅ Successfully loaded {len(all_student_data)} students and {len(test_metadata)} tests!")
        
        # Student selection with search
        student_options = sorted(all_student_data.keys())
        selected_student = st.selectbox(
            "🎓 Select Student",
            student_options,
            help="Type or select a student name"
        )
        
        if selected_student:
            student = all_student_data[selected_student]
            
            # Organize results by test type
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
                
                # Calculate percentages for each subject
                phy_pct = round((marks["phy"] / meta["max_phy"]) * 100, 1) if marks["phy"] > 0 else 0
                chem_pct = round((marks["chem"] / meta["max_chem"]) * 100, 1) if marks["chem"] > 0 else 0
                math_pct = round((marks["math"] / meta["max_math"]) * 100, 1) if marks["math"] > 0 else 0
                
                # Get weakest subject
                weakest = get_weakest_subject(marks.get("phy_rank"), marks.get("chem_rank"), marks.get("math_rank"))
                
                result = {
                    "Test Name": sheet[:30],
                    "Physics": f"{marks['phy']:.0f}/{meta['max_phy']} ({phy_pct:.0f}%)",
                    "Chemistry": f"{marks['chem']:.0f}/{meta['max_chem']} ({chem_pct:.0f}%)",
                    "Maths": f"{marks['math']:.0f}/{meta['max_math']} ({math_pct:.0f}%)",
                    "Total": f"{marks['total']:.0f}/{meta['max_total']} ({pct:.0f}%)",
                    "Rank": rank,
                    "Weakest": weakest
                }
                
                if meta["type"] == "BTEST":
                    btest_results.append(result)
                else:
                    brtest_results.append(result)
            
            # Display metrics in columns
            col1, col2, col3, col4 = st.columns(4)
            
            total_tests = len(btest_results) + len(brtest_results)
            all_pcts = [float(r["Total"].split("(")[1].replace("%)", "")) for r in btest_results + brtest_results]
            avg_pct = round(np.mean(all_pcts), 1) if all_pcts else 0
            all_ranks = [r["Rank"] for r in btest_results + brtest_results]
            best_rank = min(all_ranks) if all_ranks else 'N/A'
            
            with col1:
                st.metric("📚 Tests Attempted", f"{total_tests}/{len(test_metadata)}")
            with col2:
                st.metric("📊 Average Score", f"{avg_pct}%")
            with col3:
                st.metric("🏆 Best Rank", best_rank)
            with col4:
                roll_display = ", ".join([str(r) for r in sorted(student["roll_numbers"])][:2])
                st.metric("🆔 Roll No", roll_display)
            
            st.markdown("---")
            
            # BTEST Results
            if btest_results:
                st.subheader("📘 BTEST/GRAND Tests (JEE Format - 300 marks)")
                st.dataframe(pd.DataFrame(btest_results), use_container_width=True, height=400)
            
            # BRTEST Results
            if brtest_results:
                st.subheader("📘 BRTEST Tests (CET Format - 200 marks)")
                st.dataframe(pd.DataFrame(brtest_results), use_container_width=True, height=300)
            
            st.markdown("---")
            
            # Subject-wise graphs for BTEST
            if btest_results:
                st.subheader("📊 Subject Performance Trend - BTEST Tests")
                
                btest_names = [t['Test Name'][:20] for t in btest_results]
                btest_phy = [float(t['Physics'].split('/')[0]) for t in btest_results]
                btest_chem = [float(t['Chemistry'].split('/')[0]) for t in btest_results]
                btest_math = [float(t['Maths'].split('/')[0]) for t in btest_results]
                
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=btest_names, y=btest_phy, mode='lines+markers', 
                                        name='Physics', line=dict(color='#3498DB', width=2)))
                fig.add_trace(go.Scatter(x=btest_names, y=btest_chem, mode='lines+markers',
                                        name='Chemistry', line=dict(color='#9B59B6', width=2)))
                fig.add_trace(go.Scatter(x=btest_names, y=btest_math, mode='lines+markers',
                                        name='Mathematics', line=dict(color='#F1C40F', width=2)))
                fig.add_hline(y=75, line_dash="dash", line_color="green", annotation_text="75% Target")
                fig.update_layout(
                    title="Subject-wise Performance (BTEST)",
                    xaxis_title="Tests",
                    yaxis_title="Marks",
                    height=450,
                    hovermode='x unified'
                )
                st.plotly_chart(fig, use_container_width=True)
            
            # Subject-wise graphs for BRTEST
            if brtest_results:
                st.subheader("📊 Subject Performance Trend - BRTEST Tests (CET Format)")
                
                brtest_names = [t['Test Name'][:20] for t in brtest_results]
                brtest_phy = [float(t['Physics'].split('/')[0]) for t in brtest_results]
                brtest_chem = [float(t['Chemistry'].split('/')[0]) for t in brtest_results]
                brtest_math = [float(t['Maths'].split('/')[0]) for t in brtest_results]
                
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=brtest_names, y=brtest_phy, mode='lines+markers',
                                        name='Physics (max 50)', line=dict(color='#3498DB', width=2)))
                fig.add_trace(go.Scatter(x=brtest_names, y=brtest_chem, mode='lines+markers',
                                        name='Chemistry (max 50)', line=dict(color='#9B59B6', width=2)))
                fig.add_trace(go.Scatter(x=brtest_names, y=brtest_math, mode='lines+markers',
                                        name='Mathematics (max 100)', line=dict(color='#F1C40F', width=2)))
                fig.add_hline(y=75, line_dash="dash", line_color="green", annotation_text="75% Target (Maths)")
                fig.update_layout(
                    title="Subject-wise Performance (BRTEST - CET Format)",
                    xaxis_title="Tests",
                    yaxis_title="Marks",
                    height=450,
                    hovermode='x unified'
                )
                st.plotly_chart(fig, use_container_width=True)
            
            # Overall percentage trend
            st.subheader("📈 Overall Performance Trend")
            all_names = [t['Test Name'][:20] for t in btest_results + brtest_results]
            all_pcts_graph = [float(t["Total"].split("(")[1].replace("%)", "")) for t in btest_results + brtest_results]
            
            fig_trend = go.Figure()
            fig_trend.add_trace(go.Scatter(
                x=all_names, y=all_pcts_graph, mode='lines+markers',
                line=dict(color='#E67E22', width=3), marker=dict(size=8)
            ))
            fig_trend.add_hline(y=75, line_dash="dash", line_color="green", annotation_text="75% Target")
            fig_trend.update_layout(
                title="Percentage Score Across All Tests",
                xaxis_title="Tests",
                yaxis_title="Percentage (%)",
                yaxis=dict(range=[0, 100]),
                height=400
            )
            st.plotly_chart(fig_trend, use_container_width=True)
            
            # Rank trend
            st.subheader("🏆 Rank Trend (Lower is Better)")
            all_ranks_graph = [r["Rank"] for r in btest_results + brtest_results]
            
            fig_rank = go.Figure()
            fig_rank.add_trace(go.Scatter(
                x=all_names, y=all_ranks_graph, mode='lines+markers',
                line=dict(color='#FF6B6B', width=3), marker=dict(size=8)
            ))
            fig_rank.update_layout(
                title="Rank Performance",
                xaxis_title="Tests",
                yaxis_title="Rank",
                yaxis=dict(autorange="reversed"),
                height=400
            )
            st.plotly_chart(fig_rank, use_container_width=True)
            
            # Weakest subject analysis
            weak_subjects = [t["Weakest"] for t in btest_results + brtest_results if t["Weakest"] not in ["Balanced", "Absent"]]
            if weak_subjects:
                st.subheader("⚠️ Weak Subject Analysis")
                from collections import Counter
                weak_counts = Counter(weak_subjects)
                
                cols = st.columns(len(weak_counts))
                for i, (subject, count) in enumerate(weak_counts.items()):
                    with cols[i]:
                        st.metric(f"⚠️ {subject}", f"Weak in {count} test(s)")
            
            # Footer
            st.markdown("---")
            st.caption("📊 Dashboard created for student performance tracking | Data source: Master Sheet Excel")
            
    else:
        st.error("❌ No student data found in the file. Please check the file format.")
else:
    st.info("👈 **Please upload the Master Sheet Excel file to get started!**")
    
    # Instructions
    with st.expander("📖 How to use this dashboard"):
        st.markdown("""
        1. **Upload your Excel file** containing all test sheets
        2. **Select a student** from the dropdown
        3. **View performance metrics** across all tests
        4. **Check subject-wise trends** and identify weak areas
        
        **File Requirements:**
        - Excel file with test sheets (BTEST/GRAND = 300 marks, BRTEST = 200 marks)
        - Each sheet should have columns: Student Name, Roll No, PHY, CHEM, MATHS, TOTAL
        - The dashboard automatically detects different test formats
        
        **Features:**
        - 📊 Real-time performance tracking
        - 📈 Subject-wise trend analysis
        - 🏆 Rank tracking across tests
        - ⚠️ Weak subject identification
        """)