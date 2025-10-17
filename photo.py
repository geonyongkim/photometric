import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import re
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import io

try:
    import pdfplumber
except ImportError:
    st.error("PDF 처리 라이브러리 `pdfplumber`가 설치되지 않았습니다. `pip install pdfplumber` 명령어로 설치해주세요.")
    st.stop()

# --- 1. 설정 (Configuration) ---
REGULATIONS = {
    "DOM_HIGH": {
        "H - 5L": {"type": "min", "value": 5100},
        "H - 2.5L": {"type": "min", "value": 20300},
        "H-V": {"type": "min", "value": "0.8*MAX"},
        "H - 2.5R": {"type": "min", "value": 20300},
        "H - 5R": {"type": "min", "value": 5100},
        "MAX": {"type": "minmax", "min_val": 40500, "max_val": 215000},
    },
    "DOM_LOW": {
        "B50L": {"type": "max", "value": 350},
        "BR": {"type": "max", "value": 1750},
        "75R": {"type": "min", "value": 10100},
        "75L": {"type": "max", "value": 10600},
        "50L": {"type": "max", "value": 18500},
        "50R": {"type": "min", "value": 10100},
        "50V": {"type": "min", "value": 5100},
        "25L": {"type": "min", "value": 1700},
        "25R": {"type": "min", "value": 1700},
        "1": {"type": "max", "value": 625}, "2": {"type": "max", "value": 625}, "3": {"type": "max", "value": 625},
        "4": {"type": "max", "value": 625}, "5": {"type": "max", "value": 625}, "6": {"type": "max", "value": 625},
        "7": {"type": "minmax", "min_val": 65, "max_val": 625},
        "8": {"type": "minmax", "min_val": 125, "max_val": 625},
        "ZONE III": {"type": "max", "value": 625},
        "ZONE III_min": {"type": "min", "value": 0},
        "ZONE IV": {"type": "min", "value": 2500},
        "ZONE I": {"type": "max", "value": "2*50R"},
        "1+2+3": {"type": "min", "value": 190},
        "4+5+6": {"type": "min", "value": 375},
    },
    "ECE": {
        "B50R (0.57U-3.43R)": {"type": "max", "value": 350}, "BL (1U-2.5L)": {"type": "max", "value": 1750},
        "Zone III": {"type": "max", "value": 625}, "Zone III_min": {"type": "min", "value": 0},
        "50L (0.86D-1.72L)": {"type": "min", "value": 10100}, "75L (0.57D-1.15L)": {"type": "min", "value": 10100},
        "50V (0.86D-V)": {"type": "min", "value": 5100}, "50R (0.86D-3.43R)": {"type": "max", "value": 18500},
        "75R (0.57D-3.43R)": {"type": "max", "value": 10600}, "25R2 (1.72D-9.00R)": {"type": "min", "value": 1700},
        "25L1 (1.72D-9.00L)": {"type": "min", "value": 1700},
        "Zone IV": {"type": "min", "value": 2500}, "Zone IV_min": {"type": "min", "value": 2500},
        "Zone I (2 x E50L)": {"type": "max", "value": "2x50L"}, "Zone I (2 x E50L)_min": {"type": "min", "value": 0},
        "B1": {"type": "max", "value": 625}, "B2": {"type": "max", "value": 625}, "B3": {"type": "max", "value": 625},
        "B4": {"type": "max", "value": 625}, "B5": {"type": "max", "value": 625}, "B6": {"type": "max", "value": 625},
        "B7": {"type": "max", "value": 625}, "B8": {"type": "max", "value": 625},
        "PT 1+2+3": {"type": "min", "value": 190}, "PT 4+5+6": {"type": "min", "value": 375},
    },
    "NAS_HIGH": {
        "2U - V": {"type": "min", "value": 1800},
        "1U - 3L": {"type": "min", "value": 6000},
        "1U - 3R": {"type": "min", "value": 6000},
        "H - V": {"type": "minmax", "min_val": 4800, "max_val": 60000},
        "H - 3L": {"type": "min", "value": 1800},
        "H - 3R": {"type": "min", "value": 1800},
        "H - 6L": {"type": "min", "value": 6000},
        "H - 6R": {"type": "min", "value": 6000},
        "H - 9L": {"type": "min", "value": 3600},
        "H - 9R": {"type": "min", "value": 3600},
        "H - 12L": {"type": "min", "value": 1800},
        "H - 12R": {"type": "min", "value": 1800},
        "1.5D - V": {"type": "min", "value": 6000},
        "1.5D - 9L": {"type": "min", "value": 2400},
        "1.5D - 9R": {"type": "min", "value": 2400},
        "2.5D - V": {"type": "min", "value": 3000},
        "2.5D - 12L": {"type": "min", "value": 1200},
        "2.5D - 12R": {"type": "min", "value": 1200},
        "4D - V": {"type": "max", "value": 9600},
    },
    "NAS_LOW": {
        "10U 90U": {"type": "minmax", "min_val": 0, "max_val": 100},
        "4U - 8R": {"type": "min", "value": 76.8},
        "4U - 8L": {"type": "min", "value": 76.8},
        "2U-4L": {"type": "min", "value": 162},
        "1.5U - 1R to 3R": {"type": "min", "value": 240},
        "1.5U - 1R to R": {"type": "max", "value": 1120},
        "1U 1.5L to L": {"type": "max", "value": 560},
        "0.5U - 1.5L to L": {"type": "max", "value": 800},
        "0.5U - 1R to 3R": {"type": "minmax", "min_val": 600, "max_val": 2160},
        "H - 4L": {"type": "min", "value": 162},
        "H - 8L": {"type": "min", "value": 76.8},
        "1.5D - 2R": {"type": "min", "value": 1800},
        "2D - 15L": {"type": "min", "value": 1200},
        "2D - 15R": {"type": "min", "value": 1200},
        "4D - 4R": {"type": "max", "value": 10000},
        "0.6D - 1.3R": {"type": "min", "value": 1200},
        "0.86D - V": {"type": "min", "value": 5400},
        "0.86D - 3.5L": {"type": "minmax", "min_val": 2160, "max_val": 9600},
        "2D - 9L": {"type": "min", "value": 1500},
        "2D - 9R": {"type": "min", "value": 1500},
        "4D - 20L": {"type": "min", "value": 360},
        "4D - 20R": {"type": "min", "value": 360},
        "4D - V": {"type": "max", "value": 9600},
    }
}

# --- 2. 핵심 기능 함수 (Core Functions) ---
def to_float(x):
    try:
        s = str(x).strip().replace('\n', ' ')
        if s in ("", "-", "None"): return None
        return float(s)
    except (ValueError, TypeError):
        return None

_PAREN_FIRST = re.compile(r'\(([-+]?\d+(?:\.\d+)?)\)')
_PLAIN_NUM = re.compile(r'(?<!\()([-+]?\d+(?:\.\d+)?)(?!\))')

def parse_num_cell(cell):
    if cell is None: return None, None
    s = str(cell).replace('\n', ' ')
    alt = to_float(_PAREN_FIRST.search(s).group(1)) if _PAREN_FIRST.search(s) else None
    plain_nums = _PLAIN_NUM.findall(s)
    prim = to_float(plain_nums[-1]) if plain_nums else None
    if prim is None and alt is not None:
        all_nums = re.findall(r'[-+]?\d+(?:\.\d+)?', s)
        if len(all_nums) == 1:
            prim = to_float(all_nums[0])
            alt = None
    return prim, alt

def parse_pdf_data(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """PDF에서 모든 테이블을 추출하고 파일 이름을 추가합니다."""
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            all_tables = []
            for p in pdf.pages:
                tables = p.extract_tables()
                if tables:
                    for t in tables:
                        df = pd.DataFrame(t)
                        if not df.empty:
                            all_tables.append(df)

        if not all_tables: 
            st.warning(f"{filename}: PDF에서 테이블을 찾을 수 없습니다.")
            return pd.DataFrame()

        parsed_data = []
        for df in all_tables:
            header_str = " ".join(str(c) for c in df.iloc[0]).replace('\n', ' ')
            if "Function" in header_str:
                df.columns = [str(c).replace('\n', ' ').strip() for c in df.iloc[0]]
                df = df.iloc[1:].reset_index(drop=True)
            else:
                df.columns = ["Function", "I_cd_raw"] + [f"Col{i}" for i in range(2, len(df.columns))]

            # 측정값 컬럼 이름 찾기
            reaim_col = next((c for c in df.columns if "Reaim I" in c), None)
            icd_col = next((c for c in df.columns if "I [cd]" in c), None)

            if not icd_col and not reaim_col:
                continue

            for _, row in df.iterrows():
                point_name = str(row.get("Function", "")).replace('\n', ' ').strip()
                if not point_name: continue
                
                # Reaim 값이 유효하면 우선 사용, 아니면 I[cd] 값 사용
                raw_val = None
                if reaim_col and pd.notna(row.get(reaim_col)) and str(row.get(reaim_col)).strip():
                    raw_val = row.get(reaim_col)
                elif icd_col and pd.notna(row.get(icd_col)):
                    raw_val = row.get(icd_col)
                
                if raw_val is None: continue

                prim, alt = parse_num_cell(raw_val)
                
                measured_min, measured_max = None, None
                if alt is not None:
                    measured_min = alt
                    measured_max = prim
                elif prim is not None:
                    measured_min = prim
                    measured_max = prim

                if measured_min is not None or measured_max is not None:
                    parsed_data.append({
                        "file": filename, 
                        "Point": point_name, 
                        "Measured_min": measured_min,
                        "Measured_max": measured_max
                    })

        return pd.DataFrame(parsed_data).drop_duplicates().reset_index(drop=True)

    except Exception as e:
        st.error(f"{filename} 처리 중 예외가 발생했습니다: {e}")
        return pd.DataFrame()

def calculate_margin(measurement_df: pd.DataFrame, regulation: dict) -> pd.DataFrame:
    if measurement_df.empty: return pd.DataFrame()

    # 1. 법규 딕셔너리를 DataFrame으로 변환
    reg_list = []
    for point, rules in regulation.items():
        if rules.get('type') == 'minmax':
            reg_list.append({'Point': point, 'Type': 'min', 'Standard_cd': rules['min_val']})
            reg_list.append({'Point': point, 'Type': 'max', 'Standard_cd': rules['max_val']})
        else:
            reg_list.append({'Point': point, 'Type': rules.get('type'), 'Standard_cd': rules.get('value')})
    reg_df = pd.DataFrame(reg_list)

    # 2. 동적 기준값 처리
    if "2*50R" in reg_df['Standard_cd'].values:
        val_50R_row = measurement_df[measurement_df['Point'] == "50R"]
        if not val_50R_row.empty:
            # 50R은 단일 값이므로 Measured_max 사용
            val_50R = val_50R_row.iloc[0]['Measured_max']
            reg_df['Standard_cd'] = reg_df['Standard_cd'].replace("2*50R", 2 * val_50R)
        else: 
            reg_df = reg_df[reg_df['Standard_cd'] != "2*50R"]

    if "0.8*MAX" in reg_df['Standard_cd'].values:
        val_max_row = measurement_df[measurement_df['Point'] == "MAX"]
        if not val_max_row.empty:
            # MAX는 단일 값이므로 Measured_max 사용
            val_max = val_max_row.iloc[0]['Measured_max']
            reg_df['Standard_cd'] = reg_df['Standard_cd'].replace("0.8*MAX", 0.8 * val_max)
        else: 
            reg_df = reg_df[reg_df['Standard_cd'] != "0.8*MAX"]
            
    if "2x50L" in reg_df['Standard_cd'].values:
        val_50L_row = measurement_df[measurement_df['Point'].str.contains("50L", na=False)]
        if not val_50L_row.empty:
            val_50L = val_50L_row.iloc[0]['Measured_max']
            reg_df['Standard_cd'] = reg_df['Standard_cd'].replace("2x50L", 2 * val_50L)
        else:
            reg_df = reg_df[reg_df['Standard_cd'] != "2x50L"]

    # 3. 정규화된 키를 사용하여 데이터 병합
    def normalize_key(key):
        s = re.sub(r'\s*\(\d+%\)', '', str(key))
        return re.sub(r'[^a-z0-9]', '', s.lower())

    measurement_df['norm_key'] = measurement_df['Point'].apply(normalize_key)
    reg_df['norm_key'] = reg_df['Point'].apply(normalize_key)
    
    merged_df = pd.merge(measurement_df, reg_df, on="norm_key", how="inner")

    if merged_df.empty:
        st.warning("측정된 포인트와 일치하는 법규 기준을 찾을 수 없습니다.")
        return pd.DataFrame()

    merged_df = merged_df.drop(columns=['Point_y']).rename(columns={'Point_x': 'Point'})
    
    # 4. 여유율 계산
    def calc_unified_margin(row):
        std = row['Standard_cd']
        
        if row['Type'] == 'min':
            meas = row['Measured_min']
        elif row['Type'] == 'max':
            meas = row['Measured_max']
        else:
            return pd.NA

        if pd.isna(std) or pd.isna(meas): return pd.NA
        try: std = float(std)
        except (ValueError, TypeError): return pd.NA

        if row['Type'] == 'min':
            if std == 0: return pd.NA
            return (meas / std) * 100.0
        elif row['Type'] == 'max':
            if meas == 0: return 10000.0
            return (std / meas) * 100.0
        return pd.NA

    merged_df["Margin_%"] = merged_df.apply(calc_unified_margin, axis=1)

    # 5. Min/Max 중복 처리: 여유가 더 나쁜(낮은) 결과만 남김
    merged_df = merged_df.sort_values("Margin_%")
    final_df = merged_df.drop_duplicates(subset=['file', 'norm_key'], keep='first')

    # 6. 최종 표시에 사용할 측정값 선택
    def get_used_cd(row):
        if row['Type'] == 'min':
            return row['Measured_min']
        return row['Measured_max']
    final_df['Used_Measured_cd'] = final_df.apply(get_used_cd, axis=1)

    return final_df.drop(columns=['norm_key', 'Measured_min', 'Measured_max']).round(2)

def judge_level(margin):
    if pd.isna(margin): return ""
    if margin < 100: return "NG"
    if margin < 120: return "Lv.3"
    return "Lv.4"

def style_level_rows(row):
    color = ''
    level = row['결과 판단']
    if level == 'NG': color = 'background-color: #FF4B4B'
    elif level == 'Lv.4': color = 'background-color: #2E8B57; color: white'
    elif level == 'Lv.3': color = 'background-color: #90EE90'
    return [color] * len(row)

def to_excel(results_by_file: dict, summary_df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Summary Sheet First
        if not summary_df.empty:
            summary_df['결과 판단'] = summary_df['Margin_%'].apply(judge_level)
            display_cols = ['file', 'Point', 'Type', 'Standard_cd', 'Used_Measured_cd', 'Margin_%', '결과 판단']
            styled_summary = summary_df[display_cols].style.apply(style_level_rows, axis=1)
            styled_summary.to_excel(writer, index=False, sheet_name='Worst_Points_Summary')

        # Individual File Sheets
        for filename, df in results_by_file.items():
            df['결과 판단'] = df['Margin_%'].apply(judge_level)
            # Sanitize sheet name
            sanitized_name = re.sub(r'[\\/*?_:\[\]]', '', filename)
            if len(sanitized_name) > 31:
                sanitized_name = sanitized_name[:31]
            
            display_cols = ['Point', 'Type', 'Standard_cd', 'Used_Measured_cd', 'Margin_%', '결과 판단']
            styled_df = df[display_cols].style.apply(style_level_rows, axis=1)
            styled_df.to_excel(writer, index=False, sheet_name=sanitized_name)
            
    return output.getvalue()

# --- 3. UI 구성 (Streamlit UI) ---

def detect_regulation_from_filename(filename: str) -> str:
    """파일 이름에서 키워드를 기반으로 법규를 추측합니다."""
    fname_lower = filename.lower()
    
    reg = "DOM" # Default
    if "ece" in fname_lower: reg = "ECE"
    elif "nas" in fname_lower: reg = "NAS"
    
    beam = "LOW" # Default
    if "high" in fname_lower: beam = "HIGH"
    
    # ECE는 LOW/HIGH 구분이 없으므로 ECE로 통일
    if reg == "ECE":
        return "ECE"
        
    return f"{reg}_{beam}"

st.set_page_config(page_title="차량 헤드램프 배광 법규 분석", layout="wide")
st.title("🚘 차량 헤드램프 배광 법규 분석기")
st.write("PDF 형식의 배광 시험 결과 보고서를 업로드하고 법규를 선택하면, 여유율을 자동으로 계산하고 시각화합니다.")

st.sidebar.header("⚙️ 설정")
uploaded_files = st.sidebar.file_uploader("PDF 보고서 업로드 (여러 파일 가능)", type=["pdf"], accept_multiple_files=True)

# 자동 법규 선택 로직
regulation_options = list(REGULATIONS.keys())
regulation_captions = ["국내(상향등)", "국내(하향등)", "유럽", "북미(상향등)", "북미(하향등)"]
default_index = 0
if uploaded_files:
    detected_reg = detect_regulation_from_filename(uploaded_files[0].name)
    if detected_reg in regulation_options:
        default_index = regulation_options.index(detected_reg)

selected_regulation = st.sidebar.radio("법규 선택", options=regulation_options, index=default_index, captions=regulation_captions)
y_axis_max = st.sidebar.number_input("Y축 최대값 설정", min_value=100, value=500, step=50)

if uploaded_files:
    st.sidebar.success(f"{len(uploaded_files)}개 파일 업로드 완료!")
    
    results_by_file = {}
    all_results_list = []

    with st.spinner(f"{len(uploaded_files)}개 PDF 파일을 분석하고 여유율을 계산하는 중입니다..."):
        for uploaded_file in uploaded_files:
            measurement_df_single = parse_pdf_data(uploaded_file.getvalue(), uploaded_file.name)
            if not measurement_df_single.empty:
                # 여러 파일일 경우, 파일명에서 법규를 다시 추측. 단일 파일은 선택된 값 사용.
                reg_key = selected_regulation
                if len(uploaded_files) > 1:
                    reg_key = detect_regulation_from_filename(uploaded_file.name)
                
                if reg_key not in REGULATIONS:
                    st.warning(f"{uploaded_file.name}에 대한 법규({reg_key})를 찾을 수 없습니다. 건너뜁니다.")
                    continue

                result_df_single = calculate_margin(measurement_df_single, REGULATIONS[reg_key])
                if not result_df_single.empty:
                    results_by_file[uploaded_file.name] = result_df_single
                    # 전체 요약용 데이터 추가
                    temp_df = result_df_single.copy()
                    temp_df['file'] = uploaded_file.name
                    all_results_list.append(temp_df)

    if results_by_file:
        # --- 요약 데이터 생성 ---
        summary_df = pd.DataFrame()
        if all_results_list:
            full_summary_df = pd.concat(all_results_list, ignore_index=True)
            worst_points_list = []
            for filename, group in full_summary_df.groupby('file'):
                worst_df = group.dropna(subset=['Margin_%'])
                if not worst_df.empty:
                    worst_point_row = worst_df.loc[worst_df['Margin_%'].idxmin()].copy()
                    worst_points_list.append(worst_point_row)
            if worst_points_list:
                summary_df = pd.DataFrame(worst_points_list).sort_values("Margin_%", ignore_index=True)

        # --- UI 드롭다운 메뉴 생성 ---
        view_options = ["📈 전체 요약"] + list(results_by_file.keys())
        selected_view = st.selectbox("분석 결과 보기", options=view_options)

        # --- 전체 요약 표시 ---
        if selected_view == "📈 전체 요약":
            st.subheader("파일별 Worst Point 요약")
            if not summary_df.empty:
                summary_df['결과 판단'] = summary_df['Margin_%'].apply(judge_level)
                display_cols = ['file', 'Point', 'Type', 'Standard_cd', 'Used_Measured_cd', 'Margin_%', '결과 판단']
                st.dataframe(summary_df[display_cols].style.apply(style_level_rows, axis=1), use_container_width=True, hide_index=True)
            else:
                st.info("요약할 데이터가 없습니다.")

        # --- 파일별 상세 분석 표시 ---
        else:
            file_name = selected_view
            result_df = results_by_file[file_name]
            result_df['결과 판단'] = result_df['Margin_%'].apply(judge_level)

            st.subheader(f"{file_name} 분석 결과")
            worst_df = result_df.dropna(subset=['Margin_%'])
            
            if not worst_df.empty:
                worst_point = worst_df.loc[worst_df['Margin_%'].idxmin()]
                st.info(f"🚨 **Worst Point**: **{worst_point['Point']}** (기준: {worst_point['Type']} {worst_point['Standard_cd']:.2f}cd, 측정: {worst_point['Used_Measured_cd']:.2f}cd, 여유율: {worst_point['Margin_%']:.2f}%)")
                
                display_cols = ['Point', 'Type', 'Standard_cd', 'Used_Measured_cd', 'Margin_%', '결과 판단']
                st.dataframe(result_df[display_cols].style.apply(style_level_rows, axis=1), use_container_width=True, hide_index=True)
                
                st.subheader("여유율 시각화 그래프")
                fig = px.bar(worst_df, x='Point', y='Margin_%', color='Type', title=f'{selected_regulation} 법규 기준 여유율', labels={'Margin_%': '여유율 (%)', 'Point': '측정 포인트'}, text='Margin_%')
                fig.update_layout(yaxis_range=[0, y_axis_max])
                fig.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
                st.plotly_chart(fig, use_container_width=True, key=f"chart1_{file_name}")

                st.subheader("여유율 정렬 그래프 (Worst 순)")
                sorted_worst_df = worst_df.sort_values("Margin_%")
                fig2 = px.bar(sorted_worst_df, x='Point', y='Margin_%', title='여유율 정렬 (Worst 순)', labels={'Margin_%': '여유율 (%)', 'Point': '측정 포인트'}, text='Margin_%')
                fig2.update_layout(yaxis_range=[0, y_axis_max])
                fig2.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
                st.plotly_chart(fig2, use_container_width=True, key=f"chart2_{file_name}")
            else:
                st.info("여유율을 계산할 수 있는 포인트가 없습니다.")
                st.dataframe(result_df, use_container_width=True)

        st.subheader("📥 전체 보고서 다운로드")
        if results_by_file:
            excel_data = to_excel(results_by_file, summary_df)
            st.download_button(label="Excel 보고서 다운로드 (요약 포함)", data=excel_data, file_name=f"photometric_report_all_files.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("업로드된 파일에서 유효한 데이터를 찾을 수 없습니다.")
else:
    st.info("사이드바에서 PDF 파일을 업로드해주세요.")
