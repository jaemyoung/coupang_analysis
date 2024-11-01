import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
import copy
import streamlit as st
import pandas as pd
from datetime import datetime

# 전체 집계를 포함한 광고 지표 계산 함수
def calculate_metrics_by_date(df, start_date, end_date):
    df['날짜'] = pd.to_datetime(df['날짜'].astype(str), format='%Y%m%d', errors='coerce')
    df = df.dropna(subset=['날짜'])
    filtered_df = df[(df['날짜'] >= start_date) & (df['날짜'] <= end_date)]
    if filtered_df.empty:
        st.warning("해당 기간에 데이터가 없습니다.")
        return pd.DataFrame()  # 데이터가 없을 경우 빈 DataFrame 반환

    test_df = filtered_df.groupby('광고 노출 지면')[['노출수', '클릭수', '총 주문수(1일)', '총 전환매출액(1일)', '광고비']].sum()
    test_df['클릭률'] = round((test_df['클릭수'] / test_df['노출수']) * 100, 2).astype(str) + '%'
    test_df['전환률'] = round((test_df['총 주문수(1일)'] / test_df['클릭수']) * 100, 2).astype(str) + '%'
    test_df['CPC'] = round(test_df['광고비'] / test_df['클릭수'], 2).astype(str) + '원'
    test_df['ROAS'] = round((test_df['총 전환매출액(1일)'] / test_df['광고비']) * 100, 2).astype(str) + '%'
    test_df['전환당비용'] = round(test_df['광고비'] / test_df['총 주문수(1일)'], 2).astype(str) + '원'
    test_df.rename(columns={'총 주문수(1일)': '주문', '총 전환매출액(1일)': '광고매출'}, inplace=True)

    # 전체 합계 행 추가
    total_row = test_df[['노출수', '클릭수', '주문', '광고매출', '광고비']].sum().astype(int)
    total_row['클릭률'] = str(round((total_row['클릭수'] / total_row['노출수']) * 100, 2)) + '%'
    total_row['전환률'] = str(round((total_row['주문'] / total_row['클릭수']) * 100, 2)) + '%'
    total_row['CPC'] = str(round(total_row['광고비'] / total_row['클릭수'], 2)) + '원'
    total_row['ROAS'] = str(round((total_row['광고매출'] / total_row['광고비']) * 100, 2)) + '%'
    total_row['전환당비용'] = str(round(total_row['광고비'] / total_row['주문'], 2)) + '원'
    total_row.name = '전체'
    test_df = pd.concat([test_df, total_row.to_frame().T])

    return test_df


# 키워드별 광고 지표 계산 함수
def calculate_metrics_by_keyword(df, start_date, end_date):
    df['날짜'] = pd.to_datetime(df['날짜'].astype(str), format='%Y%m%d', errors='coerce')
    df = df.dropna(subset=['날짜', '키워드'])  # 키워드 열의 결측치도 제거
    filtered_df = df[(df['날짜'] >= start_date) & (df['날짜'] <= end_date)]
    
    if filtered_df.empty:
        st.warning("해당 기간에 데이터가 없습니다.")
        return pd.DataFrame()  # 데이터가 없을 경우 빈 DataFrame 반환
    
    # 키워드별 집계
    test_df = filtered_df.groupby('키워드')[['노출수', '클릭수', '총 주문수(1일)', '총 전환매출액(1일)', '광고비']].sum()
    test_df['클릭률'] = round((test_df['클릭수'] / test_df['노출수']) * 100, 2).astype(str) + '%'
    test_df['전환률'] = round((test_df['총 주문수(1일)'] / test_df['클릭수']) * 100, 2).astype(str) + '%'
    test_df['CPC'] = round(test_df['광고비'] / test_df['클릭수'], 2).astype(str) + '원'
    test_df['ROAS'] = round((test_df['총 전환매출액(1일)'] / test_df['광고비']) * 100, 2).astype(str) + '%'
    test_df['전환당비용'] = round(test_df['광고비'] / test_df['총 주문수(1일)'], 2).astype(str) + '원'

    # 노출수를 기준으로 내림차순 정렬
    return test_df.sort_values(by='노출수', ascending=False)


# Streamlit 애플리케이션 인터페이스
st.title("허니인덱스 광고 지표 & 키워드 분석기")
st.write("업로드된 엑셀 파일에서 광고 지표를 분석하고, 기간별로 매출과 광고비, ROAS 통계를 표시합니다.")

# 엑셀 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type="xlsx")

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    date_df =copy.deepcopy(df)
    keyword_df  = copy.deepcopy(df)
    st.write("데이터 미리보기:", df.head())

    # 날짜 선택
    start_date = st.date_input("시작 날짜", datetime(2024, 10, 1))
    end_date = st.date_input("종료 날짜", datetime(2024, 10, 30))

    # 날짜를 datetime 형식으로 변환
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)

    if st.button("통계 자료 표시"):
        # 광고 노출 지면별 통계 자료 계산 및 표시
        metrics_by_date = calculate_metrics_by_date(date_df, start_date, end_date)
        if not metrics_by_date.empty:
            st.write("광고 노출 지면별 통계 자료:", metrics_by_date)

        # 키워드별 통계 자료 계산 및 표시
        metrics_by_keyword = calculate_metrics_by_keyword(keyword_df, start_date, end_date)
        if not metrics_by_keyword.empty:
            st.write("키워드별 통계 자료:", metrics_by_keyword)
