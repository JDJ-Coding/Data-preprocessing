import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import DataBarRule
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import sys
import os

# === 전역 설정 ===
# 분석 설정값
Z_THRESHOLD = 3 

# === 1. 데이터 로드 및 파싱 함수 ===
def load_data_and_map_headers(path):
    print(f"\n[1/3] 데이터 로딩 중: {os.path.basename(path)}")
    
    # 인코딩 자동 감지
    encodings = ['cp949', 'euc-kr', 'utf-8', 'latin1']
    df_raw = None
    
    # 구분자 ; 시도
    for enc in encodings:
        try:
            df_raw = pd.read_csv(path, sep=';', header=None, encoding=enc, low_memory=False)
            break
        except:
            continue
            
    # 실패 시 구분자 , 시도
    if df_raw is None:
        for enc in encodings:
            try:
                df_raw = pd.read_csv(path, sep=',', header=None, encoding=enc, low_memory=False)
                break
            except:
                continue

    if df_raw is None or len(df_raw) < 3:
        raise ValueError("파일을 읽을 수 없거나 데이터 구조가 올바르지 않습니다.")

    # 헤더 파싱
    tags_row = df_raw.iloc[0].astype(str).str.strip()
    names_row = df_raw.iloc[1].astype(str).str.strip()
    
    # 시간 컬럼 찾기
    time_col_idx = 0
    for i, val in enumerate(names_row):
        if 'time' in val.lower() or 'date' in val.lower():
            time_col_idx = i
            break
            
    # 메타데이터 매핑
    sensor_map = {}
    valid_indices = []
    for i in range(len(df_raw.columns)):
        if i == time_col_idx: continue
        sensor_map[i] = {'tag': tags_row.iloc[i], 'name': names_row.iloc[i]}
        valid_indices.append(i)

    # 데이터프레임 변환
    df_data = df_raw.iloc[2:].copy()
    time_col_name = 'Time'
    df_data.rename(columns={time_col_idx: time_col_name}, inplace=True)
    df_data[time_col_name] = pd.to_datetime(df_data[time_col_name], errors='coerce')
    
    for idx in valid_indices:
        df_data[idx] = pd.to_numeric(df_data[idx], errors='coerce')

    return df_data, sensor_map, time_col_name

# === 2. 이상치 탐지 로직 ===
def analyze_data(df_data, sensor_map, time_col_name):
    print("[2/3] 이상치 분석 알고리즘 실행 중...")
    summary_list = []
    all_outliers = []

    for col_idx, meta in sensor_map.items():
        vals = df_data[col_idx].dropna()
        if vals.empty: continue
        
        mean = vals.mean()
        std = vals.std()
        
        if std == 0: continue
        
        # Z-score 계산
        z = (df_data[col_idx] - mean) / std
        outlier_mask = np.abs(z) > Z_THRESHOLD
        count = outlier_mask.sum()
        
        if count > 0:
            summary_list.append({
                'iba Tag': meta['tag'],
                '명칭': meta['name'],
                '이상치 개수': count,
                '비율(%)': round(100 * count / len(df_data), 2),
                '평균값': round(mean, 2),
                '최대 초과값': round(vals[outlier_mask].max(), 2)
            })
            
            # 상세 데이터 수집
            outliers = df_data.loc[outlier_mask, [time_col_name, col_idx]].copy()
            outliers.columns = ['시간', '값']
            outliers['iba Tag'] = meta['tag']
            outliers['명칭'] = meta['name']
            outliers['기준값'] = f'|Z| > {Z_THRESHOLD}'
            outliers['초과 정도'] = (outliers['값'] - mean).abs()
            
            all_outliers.append(outliers[['시간', 'iba Tag', '명칭', '값', '기준값', '초과 정도']])

    # 결과 정리
    summary_df = pd.DataFrame(summary_list)
    if not summary_df.empty:
        summary_df = summary_df.sort_values(by='이상치 개수', ascending=False)
        summary_df.insert(0, '순위', range(1, len(summary_df) + 1))
    
    detail_df = pd.concat(all_outliers, ignore_index=True) if all_outliers else pd.DataFrame()
    return summary_df, detail_df

# === 3. 엑셀 리포트 생성 (디자인 적용) ===
def create_report(summary_df, detail_df, sensor_map, output_path):
    print(f"[3/3] 엑셀 리포트 생성 중...")
    wb = openpyxl.Workbook()
    
    # 스타일 정의
    COLOR_HEADER = "4472C4"
    COLOR_BG_KPI = "E7E6E6"
    font_title = Font(name='맑은 고딕', size=18, bold=True, color="2F5597")
    font_header = Font(name='맑은 고딕', size=11, bold=True, color="FFFFFF")
    border_thin = Border(left=Side(style='thin', color="BFBFBF"), 
                         right=Side(style='thin', color="BFBFBF"), 
                         top=Side(style='thin', color="BFBFBF"), 
                         bottom=Side(style='thin', color="BFBFBF"))
    align_center = Alignment(horizontal='center', vertical='center')

    # --- 대시보드 시트 ---
    ws_dash = wb.active
    ws_dash.title = "분석 대시보드"
    ws_dash.sheet_view.showGridLines = False
    
    ws_dash['B2'] = "소성로 센서 이상치 분석 리포트"
    ws_dash['B2'].font = font_title
    ws_dash['B3'] = f"분석 일시: {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    # KPI
    total = len(sensor_map)
    detected = len(summary_df)
    total_cnt = summary_df['이상치 개수'].sum() if not summary_df.empty else 0
    top_sensor = summary_df.iloc[0]['명칭'] if not summary_df.empty else "없음"

    kpis = [("전체 센서", f"{total} 개", "B5"), ("이상치 센서", f"{detected} 개", "D5"),
            ("총 이상치 건수", f"{total_cnt} 건", "F5"), ("최다 발생", top_sensor, "H5")]
    
    for label, val, cell_addr in kpis:
        c = openpyxl.utils.cell.column_index_from_string(cell_addr[0])
        r = int(cell_addr[1:])
        ws_dash.merge_cells(start_row=r, start_column=c, end_row=r+2, end_column=c+1)
        cell = ws_dash.cell(row=r, column=c)
        cell.value = f"{label}\n{val}"
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.fill = PatternFill(start_color=COLOR_BG_KPI, end_color=COLOR_BG_KPI, fill_type='solid')
        cell.font = Font(name='맑은 고딕', size=11, bold=True)
        for i in range(r, r+3):
            for j in range(c, c+2):
                ws_dash.cell(row=i, column=j).border = border_thin

    # 요약 테이블 및 차트
    if not summary_df.empty:
        # 차트 데이터 (숨김 영역)
        ws_dash['Z1'] = "센서"
        ws_dash['AA1'] = "건수"
        for i, row in summary_df.head(10).iterrows():
            ws_dash[f'Z{i+2}'] = row['명칭']
            ws_dash[f'AA{i+2}'] = row['이상치 개수']
        
        chart = BarChart()
        chart.title = "Top 10 이상치 발생 센서"
        chart.style = 10
        chart.height = 10
        chart.width = 20
        chart.add_data(Reference(ws_dash, min_col=27, min_row=1, max_row=11), titles_from_data=True)
        chart.set_categories(Reference(ws_dash, min_col=26, min_row=2, max_row=11))
        ws_dash.add_chart(chart, "J5")

        # 테이블
        start_row = 10
        for i, col in enumerate(summary_df.columns):
            cell = ws_dash.cell(row=start_row, column=i+2)
            cell.value = col
            cell.fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type='solid')
            cell.font = font_header
            cell.alignment = align_center
            cell.border = border_thin
        
        for r_idx, row in enumerate(dataframe_to_rows(summary_df, index=False, header=False), 1):
            for c_idx, val in enumerate(row):
                cell = ws_dash.cell(row=start_row+r_idx, column=c_idx+2)
                cell.value = val
                cell.border = border_thin
                cell.alignment = align_center
                if c_idx == 2: cell.alignment = Alignment(horizontal='left', vertical='center') # 명칭
        
        ws_dash.column_dimensions['D'].width = 45 # 명칭 열 너비
    else:
        ws_dash['B10'] = "탐지된 이상치가 없습니다."

    # --- 상세 내역 시트 ---
    ws_detail = wb.create_sheet("상세 데이터")
    if not detail_df.empty:
        for i, col in enumerate(detail_df.columns):
            cell = ws_detail.cell(row=1, column=i+1)
            cell.value = col
            cell.fill = PatternFill(start_color="538DD5", end_color="538DD5", fill_type='solid')
            cell.font = font_header
            cell.alignment = align_center
        
        for r in dataframe_to_rows(detail_df, index=False, header=False):
            ws_detail.append(r)
            
        ws_detail.freeze_panes = "A2"
        ws_detail.auto_filter.ref = ws_detail.dimensions
        ws_detail.column_dimensions['A'].width = 20
        ws_detail.column_dimensions['C'].width = 45
        
        # 데이터 막대 (F열)
        rule = DataBarRule(start_type='min', end_type='max', color="638EC6", showValue=True)
        ws_detail.conditional_formatting.add(f"F2:F{ws_detail.max_row}", rule)

    wb.save(output_path)
    return True

# === 메인 프로그램 루프 ===
def main():
    # Tkinter 루트 윈도우 생성 (계속 재사용)
    root = tk.Tk()
    root.withdraw() # 창 숨김
    root.attributes('-topmost', True) # 항상 맨 위에 표시

    print("=== 소성로 센서 분석 도구 시작 ===")

    while True:
        try:
            # 1. 파일 선택
            file_path = filedialog.askopenfilename(
                parent=root,
                title="분석할 CSV 파일을 선택하세요",
                filetypes=[("All Files", "*.*"), ("CSV files", "*.csv"), ("Excel files", "*.xlsx")]
            )

            if not file_path:
                print("파일 선택이 취소되었습니다. 프로그램을 종료합니다.")
                break # 루프 종료

            # 2. 로직 실행
            df, sensor_map, time_col = load_data_and_map_headers(file_path)
            summary_df, detail_df = analyze_data(df, sensor_map, time_col)
            
            output_filename = f"Report_{os.path.basename(file_path).split('.')[0]}_{datetime.now().strftime('%H%M%S')}.xlsx"
            if create_report(summary_df, detail_df, sensor_map, output_filename):
                print(f"완료! 저장됨: {output_filename}")
                
                # 3. 계속할지 묻기
                answer = messagebox.askyesno(
                    "분석 완료", 
                    f"분석이 완료되었습니다.\n저장 파일: {output_filename}\n\n다른 파일을 계속 분석하시겠습니까?",
                    parent=root
                )
                
                if not answer: # '아니오' 선택 시
                    break
            
        except Exception as e:
            print(f"오류 발생: {e}")
            messagebox.showerror("오류", f"처리 중 문제가 발생했습니다.\n{e}", parent=root)
            # 오류가 나도 계속할지 물어볼 수 있음 (선택 사항)
            if not messagebox.askyesno("오류", "오류가 발생했습니다. 다시 시도하시겠습니까?", parent=root):
                break

    print("프로그램을 종료합니다.")
    root.destroy()

if __name__ == "__main__":
    main()
