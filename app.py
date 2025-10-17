# -*- coding: utf-8 -*-
"""
Headlamp Photometric Analyzer (C MODE PDF 전용)
- 여러 PDF 업로드 → 표 파싱 → 표준화 → 여유율 계산 → 시각화/내보내기
- 여유율(Margin_%) = (I_cd / Min) * 100   (기준치 만족=100%)
"""

import io
import re
import hashlib
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd
import numpy as np
import streamlit as st

# 그래프: Plotly (대화형) + Matplotlib(히스토그램 보조)
import plotly.express as px
import matplotlib.pyplot as plt

# PDF 파서
try:
    import pdfplumber
except ImportError:
    st.error("""
    **[에러] PDF 처리 라이브러리가 설치되지 않았습니다.**

    이 앱은 PDF 파일 분석을 위해 `pdfplumber` 라이브러리가 반드시 필요합니다.
    
    터미널(명령 프롬프트)을 열고 아래 명령어를 실행하여 라이브러리를 설치해주세요.

    ```
    pip install pdfplumber
    ```

    설치 후 앱을 다시 실행해주세요.
    """)
    st.stop()

# ----------------------- 유틸 -----------------------

def _uniq_key(prefix: str, name: str) -> str:
    h = hashlib.md5(name.encode("utf-8")).hexdigest()[:8]
    return f"{prefix}_{h}"

def to_float(x) -> Optional[float]:
    try:
        s = str(x).strip()
        if s in ("", "-", "None", ""):
            return None
        return float(s)
    except Exception:
        return None

# 괄호값(보조값 alt)와 주값(prim) 분리 파서
_NUM = re.compile(r'[-+]?\d+(?:\.\d+)?')
_PAREN_FIRST = re.compile(r'\(([-+]?\d+(?:\.\d+)?)\)')
_PLAIN_NUM = re.compile(r'(?<!\()([-+]?\d+(?:\.\d+)?)(?!\))')

def parse_num_cell(cell) -> Tuple[Optional[float], Optional[float]]:
    if cell is None:
        return None, None
    s = str(cell)
    # alt: 괄호 안 첫 번째 수
    alt = None
    m = _PAREN_FIRST.search(s)
    if m:
        alt = to_float(m.group(1))
    # prim: 괄호 밖 마지막 수 (없으면 마지막 수)
    prim = None
    plain = _PLAIN_NUM.findall(s)
    if plain:
        prim = to_float(plain[-1])
    else:
        nums = _NUM.findall(s)
        if nums:
            prim = to_float(nums[-1])
    return prim, alt

# ----------------------- PDF 파싱 -----------------------

def extract_all_tables(file_bytes: bytes) -> List[pd.DataFrame]:
    if pdfplumber is None:
        return []
    out = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for t in page.extract_tables():
                out.append(pd.DataFrame(t))
    return out

def find_measure_table(dfs: List[pd.DataFrame]) -> pd.DataFrame:
    """헤더에 Function & I [cd] 포함된 메인 측정 테이블 선택."""
    for df in dfs:
        if df.shape[0] >= 2 and df.shape[1] >= 5:
            hdr = df.iloc[0].astype(str).fillna("")
            hdr_line = " ".join(hdr.tolist())
            if "Function" in hdr_line and ("I [cd]" in hdr_line or "I[cd]" in hdr_line):
                return df
    # fallback: 열이 가장 많은 테이블
    return sorted(dfs, key=lambda d: d.shape[1], reverse=True)[0] if dfs else pd.DataFrame()

def clean_measure_table(df: pd.DataFrame, filename: str) -> pd.DataFrame:
    """표준 컬럼으로 정리 + 숫자 파싱."""
    df2 = df.copy()
    df2.columns = [str(c).strip() for c in df2.iloc[0].tolist()]
    df2 = df2.iloc[1:].reset_index(drop=True)

    rename_map = {
        "Function": "Function",
        "Min": "Min",
        "Max": "Max",
        "I [cd]": "I_cd_raw",
        "H [°]": "H_deg_raw",
        "V [°]": "V_deg_raw",
        "N.O.K": "Result",
        "N.O.K.": "Result",
        "N.O.K\n.": "Result",
    }
    for k, v in rename_map.items():
        if k in df2.columns:
            df2 = df2.rename(columns={k: v})

    keep = [c for c in ["Function", "Min", "Max", "I_cd_raw", "H_deg_raw", "V_deg_raw", "Result"] if c in df2.columns]
    df2 = df2[keep].copy()

    out = {"file": [filename] * len(df2)}
    out["Function"] = df2["Function"].astype(str)

    out["Min"] = df2["Min"].apply(to_float) if "Min" in df2.columns else None
    out["Max"] = df2["Max"].apply(to_float) if "Max" in df2.columns else None

    for col, new in [("I_cd_raw", "I_cd"), ("H_deg_raw", "H_deg"), ("V_deg_raw", "V_deg")]:
        prims, alts = [], []
        ser = df2[col] if col in df2.columns else []
        for v in ser:
            p, a = parse_num_cell(v)
            prims.append(p)
            alts.append(a)
        out[new] = prims
        out[new + "_alt"] = alts

    out["Result"] = (
        df2["Result"].astype(str).str.extract(r"(OK|NG)", expand=False).fillna(df2["Result"].astype(str))
        if "Result" in df2.columns else [""] * len(df2)
    )

    cleaned = pd.DataFrame(out)
    return cleaned

def parse_photometric_pdf(file_bytes: bytes, filename: str) -> pd.DataFrame:
    tables = extract_all_tables(file_bytes)
    if not tables:
        return pd.DataFrame()
    mt = find_measure_table(tables)
    if mt.empty:
        return pd.DataFrame()
    df = clean_measure_table(mt, filename)
    return df

# ----------------------- 분석 로직 -----------------------

def add_margin(df: pd.DataFrame) -> pd.DataFrame:
    """여유율(Margin_%) = (I_cd / Min) * 100"""
    df2 = df.copy()
    def calc(row):
        try:
            m = float(row.get("Min"))
            v = float(row.get("I_cd"))
            if m > 0:
                return (v / m) * 100.0
            else:
                # 기준(Min)이 0이거나 0보다 작으면 여유율 계산이 무의미하므로 NaN 처리
                return np.nan
        except (ValueError, TypeError):
            # float 변환 실패 시 (None, 문자열 등) NaN 처리
            return np.nan
    df2["Margin_%"] = df2.apply(calc, axis=1)
    return df2

def build_function_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Function별 평균/최소/최대 여유율, 개수"""
    if "Margin_%" not in df.columns:
        return pd.DataFrame()
    summ = (
        df.groupby("Function", dropna=False)
          .agg(n=("Margin_%", "count"),
               avg_margin=("Margin_%", "mean"),
               min_margin=("Margin_%", "min"),
               max_margin=("Margin_%", "max"))
          .reset_index()
          .sort_values("avg_margin", ascending=True)
    )
    return summ

# ----------------------- Streamlit UI -----------------------

st.set_page_config(page_title="Headlamp Photometric Analyzer", layout="wide")
st.title("📊 Headlamp Photometric Analyzer — C MODE")

with st.sidebar:
    st.markdown("### 1) PDF 업로드")
    files = st.file_uploader(
        "C MODE 형식의 LMT 결과 PDF (여러 개 가능)",
        type=["pdf"],
        accept_multiple_files=True,
        key="uploader_cmode",
        help="헤더: Function / Min / Max / I [cd] / H [°] / V [°] / N.O.K"
    )
    st.markdown("### 2) 옵션")
    show_alt = st.checkbox("괄호값(*_alt) 컬럼도 표시", value=False, key="opt_show_alt")
    topn = st.number_input("Top-N (여유율 낮은 포인트)", min_value=3, max_value=50, value=5, step=1, key="opt_topn")

tabs = st.tabs(["Overview", "Distribution", "All Points", "Function Summary", "Export", "Logs"])
logs: List[str] = []
parsed_list: List[pd.DataFrame] = []

# --------- 파싱 ---------
if files:
    for up in files:
        try:
            fb = up.read()
            df = parse_photometric_pdf(fb, up.name)
            if df.empty:
                logs.append(f"⚠️ {up.name}: 측정 테이블을 찾지 못했습니다.")
            else:
                parsed_list.append(df)
        except Exception as e:
            logs.append(f"❌ {up.name}: 예외 발생 - {e}")

if parsed_list:
    df_all = pd.concat(parsed_list, ignore_index=True)
else:
    df_all = pd.DataFrame(columns=["file","Function","Min","Max","I_cd","I_cd_alt","H_deg","H_deg_alt","V_deg","V_deg_alt","Result"])

# 여유율 계산
if not df_all.empty:
    df_all = add_margin(df_all)

# --------- Overview ---------
with tabs[0]:
    st.subheader("요약")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("파일 수", df_all["file"].nunique() if not df_all.empty else 0)
    c2.metric("포인트 수", len(df_all))
    if not df_all.empty and "Margin_%" in df_all.columns and df_all["Margin_%"].notna().any():
        min_margin = float(df_all["Margin_%"].min())
        avg_margin = float(df_all["Margin_%"].mean())
        worst_row = df_all.loc[df_all["Margin_%"].idxmin()]
        c3.metric("최소 여유율", f"{min_margin:.1f}%")
        c4.metric("평균 여유율", f"{avg_margin:.1f}%")
        st.markdown("#### 여유가 가장 부족한 포인트 (Top-N)")
        top_worst = df_all.nsmallest(int(topn), "Margin_%")[["file","Function","Min","I_cd","Margin_%"]]
        st.dataframe(top_worst, use_container_width=True)
        # 막대 그래프 (Top-N)
        fig = px.bar(top_worst, x="Function", y="Margin_%", text="Margin_%", hover_data=["file","Min","I_cd"])
        fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig.update_layout(yaxis_title="Average Margin [%]", xaxis_title="Function", title="Top-N Lowest Margin")
        st.plotly_chart(fig, use_container_width=True)
    else:
        c3.metric("최소 여유율", "-")
        c4.metric("평균 여유율", "-")
        st.info("여유율을 계산할 수 있는 데이터가 없습니다.")

# --------- Distribution ---------
with tabs[1]:
    st.subheader("전 포인트 여유율 분포")
    if not df_all.empty and df_all["Margin_%"].notna().any():
        fig, ax = plt.subplots(figsize=(8,4), dpi=150)
        ax.hist(df_all["Margin_%"].dropna(), bins=15)
        ax.set_xlabel("Margin [%]")
        ax.set_ylabel("Count")
        ax.set_title("Distribution of Margin (%)")
        ax.grid(axis="y", alpha=0.4)
        st.pyplot(fig, use_container_width=True)
    else:
        st.info("표시할 여유율 데이터가 없습니다.")

# --------- All Points (정렬 표 + 전체 막대) ---------
with tabs[2]:
    st.subheader("전 포인트 정렬표 (여유율 오름차순)")
    if not df_all.empty and df_all["Margin_%"].notna().any():
        show_cols = ["file","Function","Min","I_cd","Margin_%"]
        if show_alt:
            show_cols += ["I_cd_alt","H_deg","H_deg_alt","V_deg","V_deg_alt","Result"]
        sorted_points = df_all.sort_values("Margin_%")[show_cols].reset_index(drop=True)
        st.dataframe(sorted_points, use_container_width=True, hide_index=True)
        st.caption("※ 막대그래프는 Left→Right로 여유율이 증가합니다.")
        fig2 = px.bar(sorted_points, x="Function", y="Margin_%", hover_data=["file","Min","I_cd"])
        fig2.update_layout(xaxis={'categoryorder':'array','categoryarray':sorted_points["Function"].tolist()},
                           yaxis_title="Margin [%]", xaxis_title="Function",
                           title="All Points Sorted by Margin (Ascending)")
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("표시할 데이터가 없습니다.")

# --------- Function Summary ---------
with tabs[3]:
    st.subheader("Function 요약 (평균/최소/최대 여유율)")
    if not df_all.empty and df_all["Margin_%"].notna().any():
        func_sum = build_function_summary(df_all)
        st.dataframe(func_sum, use_container_width=True, hide_index=True)
        fig3 = px.bar(func_sum.sort_values("avg_margin"), x="Function", y="avg_margin",
                      hover_data=["n","min_margin","max_margin"], title="Average Margin by Function")
        fig3.update_layout(yaxis_title="Average Margin [%]")
        st.plotly_chart(fig3, use_container_width=True)
    else:
        st.info("표시할 데이터가 없습니다.")

# --------- Export ---------
with tabs[4]:
    st.subheader("내보내기")
    if not df_all.empty:
        # 정규화 전체 CSV
        csv_bytes = df_all.to_csv(index=False).encode("utf-8-sig")
        st.download_button("전체 정규화 CSV 다운로드", data=csv_bytes,
                           file_name="all_results_normalized.csv",
                           mime="text/csv", key=_uniq_key("dl_csv", "all"))
        # Function 요약 CSV
        func_sum = build_function_summary(df_all)
        if not func_sum.empty:
            csv_bytes2 = func_sum.to_csv(index=False).encode("utf-8-sig")
            st.download_button("Function 요약 CSV 다운로드", data=csv_bytes2,
                               file_name="function_margin_summary.csv",
                               mime="text/csv", key=_uniq_key("dl_csv", "func"))
        # 엑셀(두 시트)
        from io import BytesIO
        xls_buf = BytesIO()
        with pd.ExcelWriter(xls_buf, engine="xlsxwriter") as writer:
            df_all.to_excel(writer, index=False, sheet_name="AllResults")
            if not func_sum.empty:
                func_sum.to_excel(writer, index=False, sheet_name="FunctionSummary")
        st.download_button("Excel(xlsx) 다운로드 (All + Summary)", data=xls_buf.getvalue(),
                           file_name="photometric_analysis.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key=_uniq_key("dl_xlsx", "both"))
    else:
        st.info("내보낼 데이터가 없습니다.")

# --------- Logs ---------
with tabs[5]:
    st.subheader("파싱 로그")
    if logs:
        for m in logs:
            st.write(m)
    else:
        st.caption("문제 없이 처리되었습니다.")
