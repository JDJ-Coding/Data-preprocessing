import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
import requests
import json

# ==============================================================================
# [사용자 설정] 환경변수 확인
# ==============================================================================
POSCO_GPT_KEY = os.environ.get('POSCO_GPT_KEY')

# ==============================================================================

try:
    from sklearn.ensemble import IsolationForest
except ImportError:
    print("[Error] scikit-learn 설치가 필요합니다. (pip install scikit-learn requests)")
    sys.exit()

# ==========================================
# [핵심] 사내 AI API 연동 함수
# ==========================================
def call_corporate_ai_api(prompt_text):
    if not POSCO_GPT_KEY:
        return "[설정 오류] 환경변수 POSCO_GPT_KEY가 없습니다."

    try:
        url = "http://aigpt.posco.net/gpgpta01-gpt/gptApi/personalApi"

        token_value = POSCO_GPT_KEY
        if not token_value.startswith("Bearer "):
            token_value = f"Bearer {token_value}"

        headers = {
            "Authorization": token_value,
            "Content-Type": "application/json"
        }

        payload = {
            "model": "gpt-5.2",
            "messages": [
                {"role": "system", "content": "당신은 포스코 퓨처엠의 설비 예지보전 전문가입니다. 수치를 기반으로 냉철하게 설비 위험도를 판단하십시오."},
                {"role": "user", "content": prompt_text}
            ],
            "frequency_penalty": 0.0,
            "temperature": 0.3
        }

        response = requests.post(url, headers=headers, json=payload, timeout=60)
        response.encoding = 'utf-8'

        if response.status_code == 200:
            try:
                result = response.json()
                if 'choices' in result:
                    return result['choices'][0]['message']['content'].strip()
                elif 'response' in result:
                    return result['response']
                else:
                    return str(result)
            except json.JSONDecodeError:
                return response.text.strip()
        else:
            return f"[API Error] 상태코드: {response.status_code}"

    except Exception as e:
        return f"[시스템 오류] {str(e)}"


# ==========================================
# 1. 데이터 로드 (안정성 강화 버전)
# ==========================================
def load_data(path):
    print(f"[1/5] 데이터 로딩 중: {os.path.basename(path)}")
    encodings = ['cp949', 'euc-kr', 'utf-8']
    df_raw = None
    separators = [',', ';', '\t']

    for sep in separators:
        for enc in encodings:
            try:
                df_raw = pd.read_csv(path, sep=sep, header=None, encoding=enc, low_memory=False)
                if df_raw.shape[1] > 1:
                    break
                else:
                    df_raw = None
            except:
                continue
        if df_raw is not None:
            break

    if df_raw is None or len(df_raw) < 3:
        raise ValueError("파일을 읽을 수 없습니다. (형식이 올바르지 않음)")

    tags = df_raw.iloc[0].astype(str).str.strip()
    names = df_raw.iloc[1].astype(str).str.strip()

    time_col_idx = 0
    for i, val in enumerate(names):
        if 'time' in val.lower() or 'date' in val.lower():
            time_col_idx = i
            break

    sensor_map = {}
    valid_indices = []

    for i in range(len(df_raw.columns)):
        if i == time_col_idx:
            continue
        if not tags.iloc[i] or tags.iloc[i] == 'nan':
            continue
        sensor_map[i] = {'tag': tags.iloc[i], 'name': names.iloc[i]}
        valid_indices.append(i)

    df_data = df_raw.iloc[2:].copy()
    df_data.rename(columns={time_col_idx: 'Time'}, inplace=True)

    for i in valid_indices:
        df_data[i] = pd.to_numeric(df_data[i], errors='coerce')
        df_data[i] = df_data[i].interpolate(method='linear').bfill().fillna(0)

    return df_data, sensor_map


# ==========================================
# 2. 위험 등급 판정 함수
# ==========================================
def classify_risk(score, max_z, max_consecutive, count_ratio, change_ratio):
    """
    3단계 위험 등급 판정:
      위험 : Isolation Forest 심각도 > 0.15 OR Z-score > 5 OR 연속이상 > 20건 OR 이상비율 > 5%
      주의 : Isolation Forest 심각도 > 0.08 OR Z-score > 3 OR 연속이상 > 10건 OR 이상비율 > 2%
      정상 : 위 기준 미달
    """
    if score > 0.15 or max_z > 5 or max_consecutive > 20 or count_ratio > 5:
        return "위험"
    elif score > 0.08 or max_z > 3 or max_consecutive > 10 or count_ratio > 2:
        return "주의"
    else:
        return "정상"


# ==========================================
# 3. 권고 조치 텍스트 생성
# ==========================================
def build_recommendation(sensor_name, risk, direction, count_ratio, max_consecutive, trend, max_z):
    """설비 담당자가 바로 실행할 수 있는 권고 조치 생성"""
    parts = []

    # 위험도별 긴급도
    if risk == "위험":
        parts.append("【즉시 조치】현장 점검 및 상위 관리자 보고 필요.")
    elif risk == "주의":
        parts.append("【조기 대응】점검 주기 단축 및 추이 집중 모니터링 권고.")
    else:
        parts.append("【현 상태 유지】정기 점검 일정 준수.")

    # 온도 센서 특화 판단
    name_upper = sensor_name.upper()
    if "TEMP" in name_upper or "온도" in sensor_name:
        if direction == "상승(▲)":
            parts.append("온도 이상 상승 → 히터 과부하·냉각 불량·단열재 손상 여부 점검.")
        else:
            parts.append("온도 이상 하강 → 히터 단선·전원 공급 불량·열전대 이상 의심.")
    elif "CV" in name_upper or "밸브" in sensor_name or "VALVE" in name_upper:
        if direction == "상승(▲)":
            parts.append("밸브 개도 이상 상승 → 액추에이터 과작동·포지셔너 오류 점검.")
        else:
            parts.append("밸브 개도 이상 하강 → 밸브 고착·공압 공급 불량 점검.")
    elif "PRESS" in name_upper or "압력" in sensor_name:
        if direction == "상승(▲)":
            parts.append("압력 이상 상승 → 배관 막힘·안전밸브 작동 여부 확인.")
        else:
            parts.append("압력 이상 하강 → 누설·펌프 성능 저하 확인.")
    elif "FLOW" in name_upper or "유량" in sensor_name:
        parts.append("유량 이상 → 배관 막힘·유량계 오염·펌프 성능 점검 권고.")

    # 연속 이상 구간
    if max_consecutive >= 30:
        parts.append(f"연속 {max_consecutive}건 지속 이상 → 일시 노이즈 아닌 구조적 결함 강력 의심, 설비 정지 검토.")
    elif max_consecutive >= 10:
        parts.append(f"연속 {max_consecutive}건 지속 이상 → 간헐적 고장이 아닌 지속 결함 가능성, 긴급 점검 권고.")

    # 이상 비율
    if count_ratio > 5:
        parts.append(f"전체 데이터 {count_ratio:.1f}% 이상치 → 기준값 재검토 또는 센서 교체 고려.")

    # Z-score 기반
    if max_z > 5:
        parts.append(f"최대 Z-score {max_z:.1f}σ → 극단적 편차, 센서 오류 또는 설비 고장 즉시 확인.")

    # 추세 분석
    if trend == "상승추세(▲)":
        parts.append("시계열 상승 추세 감지 → 마모·열화 진행 중일 가능성, 수명 주기 검토 필요.")
    elif trend == "하강추세(▼)":
        parts.append("시계열 하강 추세 감지 → 성능 저하 진행 중, 예방 정비 계획 수립 권고.")

    return " ".join(parts)


# ==========================================
# 4. 핵심 분석 엔진 (고도화)
# ==========================================
def compute_enhanced_metrics(df, sensor_map):
    """
    Isolation Forest + Z-score + 이동평균 + 연속이상 + 추세분석을 통합한
    고도화 센서 이상 분석 엔진
    """
    print("[2/5] 고도화 이상 분석 수행 중 (Z-score·이동평균·연속이상·추세 포함)...")

    results = []
    anomaly_details = []

    clf = IsolationForest(contamination=0.01, random_state=42, n_jobs=-1)
    MA_WINDOW = 30  # 이동평균 윈도우 (데이터 포인트 기준)

    total = len(sensor_map)
    for idx_count, (col_idx, meta) in enumerate(sensor_map.items()):
        if (idx_count + 1) % 200 == 0:
            print(f"   ... {idx_count + 1}/{total} 센서 처리 완료")

        series = df[col_idx].reset_index(drop=True).astype(float)

        # 분석 제외 조건: 표준편차 0 또는 0/1 디지털 신호
        if series.std() < 1e-9:
            continue
        unique_vals = series.dropna().unique()
        if len(unique_vals) <= 2 and set(np.round(unique_vals)).issubset({0.0, 1.0}):
            continue
        # 유효 데이터 90% 미만 제외
        if series.isna().mean() > 0.1:
            continue

        # --- 이동평균(MA) 및 이탈량 계산 ---
        ma = series.rolling(window=MA_WINDOW, min_periods=1).mean()
        ma_std = series.rolling(window=MA_WINDOW, min_periods=1).std().fillna(0)
        ma_deviation = series - ma  # 이동평균 이탈량 (양수=상방 이탈, 음수=하방 이탈)

        # --- 전체 Z-score ---
        global_mean = series.mean()
        global_std = series.std() + 1e-9
        z_scores = (series - global_mean) / global_std

        # --- Isolation Forest 이상치 탐지 ---
        X = series.values.reshape(-1, 1)
        clf.fit(X)
        preds = clf.predict(X)
        if_scores = clf.decision_function(X)  # 음수일수록 이상치

        outlier_mask = preds == -1
        outlier_idx = np.where(outlier_mask)[0]
        count = len(outlier_idx)

        if count == 0:
            continue

        # --- 연속 이상 구간 탐지 ---
        max_consecutive = 0
        current_consecutive = 0
        consecutive_runs = []
        run_start = None

        for k in range(len(series)):
            if outlier_mask[k]:
                current_consecutive += 1
                if current_consecutive == 1:
                    run_start = k
                if current_consecutive > max_consecutive:
                    max_consecutive = current_consecutive
            else:
                if current_consecutive > 0:
                    consecutive_runs.append((run_start, k - 1, current_consecutive))
                current_consecutive = 0

        if current_consecutive > 0:
            consecutive_runs.append((run_start, len(series) - 1, current_consecutive))

        # --- 추세 분석 (선형 회귀 기울기) ---
        x_axis = np.arange(len(series))
        slope, intercept = np.polyfit(x_axis, series.values, 1)
        # 기울기를 평균값 대비 상대적 크기로 정규화
        rel_slope = slope / (abs(global_mean) + 1e-9)
        if rel_slope > 0.0001:
            trend = "상승추세(▲)"
        elif rel_slope < -0.0001:
            trend = "하강추세(▼)"
        else:
            trend = "안정(─)"

        # --- 이상치 통계 ---
        anomalies = series.iloc[outlier_idx]
        anomaly_mean = anomalies.mean()
        direction = "상승(▲)" if anomaly_mean > global_mean else "하강(▼)"
        change_ratio = (abs(anomaly_mean - global_mean) / (abs(global_mean) + 1e-9)) * 100

        z_at_anomalies = z_scores.iloc[outlier_idx]
        max_z = z_at_anomalies.abs().max()

        ma_dev_at_anomalies = ma_deviation.iloc[outlier_idx]
        max_ma_dev = ma_dev_at_anomalies.abs().max()
        avg_ma_dev = ma_dev_at_anomalies.abs().mean()

        count_ratio = count / len(series) * 100
        if_score = abs(if_scores[outlier_idx].mean())

        # 전반부/후반부 이상 분포
        n = len(series)
        early_count = int(np.sum(outlier_idx < n // 2))
        late_count = int(np.sum(outlier_idx >= n // 2))

        # --- 위험 등급 ---
        risk = classify_risk(if_score, max_z, max_consecutive, count_ratio, change_ratio)

        # --- 권고 조치 ---
        rec = build_recommendation(
            meta['name'], risk, direction,
            count_ratio, max_consecutive, trend, max_z
        )

        results.append({
            'Col_Idx': col_idx,
            'iba Tag': meta['tag'],
            '명칭': meta['name'],
            '위험등급': risk,
            '이상징후 건수': count,
            '이상 비율(%)': round(count_ratio, 2),
            '패턴 방향': direction,
            '시계열 추세': trend,
            '정상 평균': round(global_mean, 2),
            '이상구간 평균': round(anomaly_mean, 2),
            '정상 대비 변화율(%)': round(change_ratio, 2),
            '최대 Z-score': round(max_z, 2),
            '이동평균 최대이탈': round(max_ma_dev, 2),
            '이동평균 평균이탈': round(avg_ma_dev, 2),
            '연속이상 최대구간(건)': max_consecutive,
            '전반부 이상건수': early_count,
            '후반부 이상건수': late_count,
            '이상치 최대값': round(anomalies.max(), 2),
            '이상치 최소값': round(anomalies.min(), 2),
            'IF 심각도 점수': round(if_score, 4),
            '권고 조치': rec,
            'AI 위험도 판단': '분석 대기'
        })

        # --- 이상치 상세 이력 (타임라인) ---
        times = df['Time'].reset_index(drop=True)
        for oi in outlier_idx:
            anomaly_details.append({
                'iba Tag': meta['tag'],
                '명칭': meta['name'],
                '위험등급': risk,
                '발생 시각': str(times.iloc[oi]) if oi < len(times) else '',
                '측정값': round(float(series.iloc[oi]), 2),
                '정상 평균': round(global_mean, 2),
                '이동평균(MA30)': round(float(ma.iloc[oi]), 2),
                'Z-score': round(float(z_scores.iloc[oi]), 2),
                '이동평균 이탈량': round(float(ma_deviation.iloc[oi]), 2),
                '이상 방향': direction
            })

    print(f"   ... 분석 완료 ({len(results)}개 센서에서 이상 징후 감지)")

    if not results:
        return pd.DataFrame(), pd.DataFrame()

    df_results = pd.DataFrame(results).sort_values(by='IF 심각도 점수', ascending=False).reset_index(drop=True)
    df_details = pd.DataFrame(anomaly_details)

    return df_results, df_details


# ==========================================
# 5. 공통 스타일 정의
# ==========================================
def get_styles():
    styles = {
        'fill_danger':    PatternFill(start_color="FF0000", end_color="FF0000", fill_type='solid'),
        'fill_warning':   PatternFill(start_color="FFC000", end_color="FFC000", fill_type='solid'),
        'fill_ok':        PatternFill(start_color="92D050", end_color="92D050", fill_type='solid'),
        'fill_header':    PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type='solid'),
        'fill_subheader': PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type='solid'),
        'fill_row_alt':   PatternFill(start_color="DEEAF1", end_color="DEEAF1", fill_type='solid'),
        'fill_row_light': PatternFill(start_color="F2F7FB", end_color="F2F7FB", fill_type='solid'),
        'fill_danger_lt': PatternFill(start_color="FFE0E0", end_color="FFE0E0", fill_type='solid'),
        'fill_warning_lt':PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type='solid'),
        'fill_ok_lt':     PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type='solid'),
        'font_header':    Font(color="FFFFFF", bold=True, size=11),
        'font_title':     Font(color="FFFFFF", bold=True, size=13),
        'font_danger':    Font(color="C00000", bold=True),
        'font_warning':   Font(color="7F4F00", bold=True),
        'font_ok':        Font(color="375623", bold=True),
        'font_bold':      Font(bold=True),
        'border_thin': Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'),  bottom=Side(style='thin')
        ),
        'border_medium': Border(
            left=Side(style='medium'), right=Side(style='medium'),
            top=Side(style='medium'),  bottom=Side(style='medium')
        ),
        'align_center': Alignment(horizontal='center', vertical='center', wrap_text=True),
        'align_left':   Alignment(horizontal='left',   vertical='center', wrap_text=True),
        'align_top':    Alignment(horizontal='left',   vertical='top',    wrap_text=True),
    }
    return styles


def apply_risk_style(cell, risk, styles, light=False):
    """위험 등급에 따라 셀 색상·폰트 적용"""
    if risk == "위험":
        cell.fill = styles['fill_danger_lt'] if light else styles['fill_danger']
        cell.font = styles['font_danger'] if light else Font(color="FFFFFF", bold=True)
    elif risk == "주의":
        cell.fill = styles['fill_warning_lt'] if light else styles['fill_warning']
        cell.font = styles['font_warning'] if light else Font(color="7F4F00", bold=True)
    else:
        cell.fill = styles['fill_ok_lt'] if light else styles['fill_ok']
        cell.font = styles['font_ok'] if light else Font(color="375623", bold=True)


# ==========================================
# 6. Sheet 1: 종합 요약 대시보드
# ==========================================
def create_dashboard_sheet(wb, df_summary, styles):
    ws = wb.active
    ws.title = "종합 요약 대시보드"
    ws.sheet_view.showGridLines = False

    danger_count  = int((df_summary['위험등급'] == '위험').sum())
    warning_count = int((df_summary['위험등급'] == '주의').sum())
    ok_count      = int((df_summary['위험등급'] == '정상').sum())
    total_sensors = len(df_summary)

    # ── 타이틀 영역 ──────────────────────────────────────────────
    ws.merge_cells('A1:L1')
    title_cell = ws['A1']
    title_cell.value = "소성로 히터 센서 이상 분석 리포트  |  포스코 퓨처엠 설비 예지보전"
    title_cell.fill = styles['fill_header']
    title_cell.font = Font(color="FFFFFF", bold=True, size=15)
    title_cell.alignment = styles['align_center']
    ws.row_dimensions[1].height = 32

    ws.merge_cells('A2:L2')
    ts_cell = ws['A2']
    ts_cell.value = f"분석 일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}     |     분석 대상: {total_sensors:,}개 센서"
    ts_cell.fill = styles['fill_subheader']
    ts_cell.font = Font(color="FFFFFF", size=10)
    ts_cell.alignment = styles['align_center']
    ws.row_dimensions[2].height = 20

    ws.row_dimensions[3].height = 10

    # ── 위험 등급 카드 영역 ───────────────────────────────────────
    def write_card(start_col, label, count, fill, font_color, desc):
        col_l = openpyxl.utils.get_column_letter
        # 라벨 셀
        ws.merge_cells(f'{col_l(start_col)}4:{col_l(start_col+3)}4')
        c = ws.cell(row=4, column=start_col, value=label)
        c.fill = fill; c.font = Font(color=font_color, bold=True, size=12)
        c.alignment = styles['align_center']
        ws.row_dimensions[4].height = 22

        ws.merge_cells(f'{col_l(start_col)}5:{col_l(start_col+3)}5')
        c = ws.cell(row=5, column=start_col, value=count)
        c.fill = fill; c.font = Font(color=font_color, bold=True, size=28)
        c.alignment = styles['align_center']
        ws.row_dimensions[5].height = 48

        ws.merge_cells(f'{col_l(start_col)}6:{col_l(start_col+3)}6')
        c = ws.cell(row=6, column=start_col, value=desc)
        c.fill = fill; c.font = Font(color=font_color, size=9)
        c.alignment = styles['align_center']
        ws.row_dimensions[6].height = 18

    write_card(1,  "⚠  위험  센서",  f"{danger_count}개",
               styles['fill_danger'],  "FFFFFF",
               "즉시 현장 점검 필요 | 설비 중단 검토")
    write_card(5,  "⚡  주의  센서",  f"{warning_count}개",
               styles['fill_warning'], "7F4F00",
               "점검 주기 단축 필요 | 추이 집중 모니터링")
    write_card(9,  "✓  정상  센서",  f"{ok_count}개",
               styles['fill_ok'],      "375623",
               "현 상태 유지 | 정기 점검 일정 준수")

    ws.row_dimensions[7].height = 12

    # ── 판정 기준 안내 ────────────────────────────────────────────
    ws.merge_cells('A8:L8')
    c = ws.cell(row=8, column=1,
                value="【 판정 기준 】  위험: IF심각도>0.15 OR Z-score>5σ OR 연속이상>20건 OR 이상비율>5%  |  "
                      "주의: IF심각도>0.08 OR Z-score>3σ OR 연속이상>10건 OR 이상비율>2%  |  정상: 기준 미달")
    c.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type='solid')
    c.font = Font(size=9, italic=True, color="595959")
    c.alignment = styles['align_center']
    ws.row_dimensions[8].height = 16

    ws.row_dimensions[9].height = 10

    # ── TOP 20 위험 센서 테이블 ───────────────────────────────────
    headers_top20 = [
        "순위", "위험등급", "iba Tag", "센서 명칭",
        "이상건수", "이상비율(%)", "최대 Z-score",
        "이동평균 최대이탈", "연속이상 최대(건)",
        "패턴 방향", "시계열 추세", "권고 조치"
    ]
    col_widths_top20 = [6, 9, 14, 40, 10, 12, 13, 18, 18, 12, 12, 55]

    ws.merge_cells('A10:L10')
    c = ws.cell(row=10, column=1, value="▶  TOP 20 위험 센서 (즉시 조치 대상)")
    c.fill = styles['fill_subheader']
    c.font = Font(color="FFFFFF", bold=True, size=11)
    c.alignment = styles['align_center']
    ws.row_dimensions[10].height = 20

    for ci, (h, w) in enumerate(zip(headers_top20, col_widths_top20), 1):
        c = ws.cell(row=11, column=ci, value=h)
        c.fill = styles['fill_header']
        c.font = styles['font_header']
        c.alignment = styles['align_center']
        c.border = styles['border_thin']
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w
    ws.row_dimensions[11].height = 20

    top20 = df_summary.head(20)
    col_keys = [
        None, '위험등급', 'iba Tag', '명칭',
        '이상징후 건수', '이상 비율(%)', '최대 Z-score',
        '이동평균 최대이탈', '연속이상 최대구간(건)',
        '패턴 방향', '시계열 추세', '권고 조치'
    ]

    for row_i, (_, row) in enumerate(top20.iterrows(), 12):
        risk_val = row['위험등급']
        row_fill = styles['fill_danger_lt'] if risk_val == '위험' else (
                   styles['fill_warning_lt'] if risk_val == '주의' else styles['fill_ok_lt'])

        for ci, key in enumerate(col_keys, 1):
            if key is None:
                val = row_i - 11  # 순위
            else:
                val = row[key]
            c = ws.cell(row=row_i, column=ci, value=val)
            c.border = styles['border_thin']

            if ci == 2:  # 위험등급 컬럼은 배지 스타일
                apply_risk_style(c, risk_val, styles, light=False)
                c.alignment = styles['align_center']
            elif key == '권고 조치':
                c.alignment = styles['align_top']
                c.fill = row_fill
            elif ci == 1:
                c.alignment = styles['align_center']
                c.font = Font(bold=True)
                c.fill = row_fill
            else:
                c.alignment = styles['align_center']
                c.fill = row_fill

        ws.row_dimensions[row_i].height = 30

    # 컬럼 너비 설정
    for ci, w in enumerate(col_widths_top20, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w


# ==========================================
# 7. Sheet 2: 센서별 상세 분석표
# ==========================================
def create_sensor_detail_sheet(wb, df_summary, styles):
    ws = wb.create_sheet(title="센서별 상세분석")
    ws.sheet_view.showGridLines = False

    # 헤더 정의 (Col_Idx, AI 위험도 판단 제외하고 표시)
    display_cols = [c for c in df_summary.columns if c not in ('Col_Idx',)]

    col_widths = {
        'iba Tag': 14, '명칭': 38, '위험등급': 9,
        '이상징후 건수': 11, '이상 비율(%)': 11,
        '패턴 방향': 11, '시계열 추세': 11,
        '정상 평균': 11, '이상구간 평균': 12,
        '정상 대비 변화율(%)': 14,
        '최대 Z-score': 12, '이동평균 최대이탈': 16, '이동평균 평균이탈': 16,
        '연속이상 최대구간(건)': 16,
        '전반부 이상건수': 13, '후반부 이상건수': 13,
        '이상치 최대값': 12, '이상치 최소값': 12,
        'IF 심각도 점수': 13,
        '권고 조치': 55, 'AI 위험도 판단': 55
    }

    # 타이틀
    ws.merge_cells(f'A1:{openpyxl.utils.get_column_letter(len(display_cols))}1')
    c = ws.cell(row=1, column=1, value="센서별 상세 이상 분석표  ─  모든 이상 징후 센서 전체 목록")
    c.fill = styles['fill_header']
    c.font = styles['font_title']
    c.alignment = styles['align_center']
    ws.row_dimensions[1].height = 28

    # 헤더 행
    for ci, col in enumerate(display_cols, 1):
        c = ws.cell(row=2, column=ci, value=col)
        c.fill = styles['fill_subheader']
        c.font = styles['font_header']
        c.alignment = styles['align_center']
        c.border = styles['border_thin']
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = col_widths.get(col, 13)
    ws.row_dimensions[2].height = 22

    # 데이터 행
    for row_i, (_, row) in enumerate(df_summary.iterrows(), 3):
        risk_val = row['위험등급']
        alt_fill = styles['fill_row_alt'] if row_i % 2 == 0 else styles['fill_row_light']

        for ci, col in enumerate(display_cols, 1):
            val = row[col]
            c = ws.cell(row=row_i, column=ci, value=val)
            c.border = styles['border_thin']

            if col == '위험등급':
                apply_risk_style(c, risk_val, styles, light=False)
                c.alignment = styles['align_center']
            elif col in ('권고 조치', 'AI 위험도 판단', '명칭'):
                c.alignment = styles['align_top']
                c.fill = alt_fill
            else:
                c.alignment = styles['align_center']
                c.fill = alt_fill

            # 숫자 컬럼 조건부 색상 강조
            if col == '최대 Z-score' and isinstance(val, (int, float)):
                if val > 5:
                    c.fill = styles['fill_danger_lt']
                    c.font = styles['font_danger']
                elif val > 3:
                    c.fill = styles['fill_warning_lt']
                    c.font = styles['font_warning']

            if col == '연속이상 최대구간(건)' and isinstance(val, (int, float)):
                if val > 20:
                    c.fill = styles['fill_danger_lt']
                    c.font = styles['font_danger']
                elif val > 10:
                    c.fill = styles['fill_warning_lt']
                    c.font = styles['font_warning']

            if col == '이상 비율(%)' and isinstance(val, (int, float)):
                if val > 5:
                    c.fill = styles['fill_danger_lt']
                    c.font = styles['font_danger']
                elif val > 2:
                    c.fill = styles['fill_warning_lt']
                    c.font = styles['font_warning']

        ws.row_dimensions[row_i].height = 28


# ==========================================
# 8. Sheet 3: 이상치 상세 이력 (타임라인)
# ==========================================
def create_anomaly_timeline_sheet(wb, df_details, styles):
    if df_details.empty:
        return

    ws = wb.create_sheet(title="이상치 상세이력")
    ws.sheet_view.showGridLines = False

    # 위험 등급 순 정렬 (위험 > 주의 > 정상)
    risk_order = {'위험': 0, '주의': 1, '정상': 2}
    df_sorted = df_details.copy()
    df_sorted['_sort'] = df_sorted['위험등급'].map(risk_order)
    df_sorted = df_sorted.sort_values(['_sort', '발생 시각']).drop(columns=['_sort'])

    display_cols = [c for c in df_sorted.columns]
    col_widths_map = {
        'iba Tag': 14, '명칭': 38, '위험등급': 9,
        '발생 시각': 22, '측정값': 12, '정상 평균': 12,
        '이동평균(MA30)': 14, 'Z-score': 11,
        '이동평균 이탈량': 14, '이상 방향': 11
    }

    ws.merge_cells(f'A1:{openpyxl.utils.get_column_letter(len(display_cols))}1')
    c = ws.cell(row=1, column=1,
                value=f"이상치 상세 발생 이력  ─  총 {len(df_sorted):,}건  (위험 등급 → 발생 시각 순 정렬)")
    c.fill = styles['fill_header']
    c.font = styles['font_title']
    c.alignment = styles['align_center']
    ws.row_dimensions[1].height = 28

    # 헤더
    for ci, col in enumerate(display_cols, 1):
        c = ws.cell(row=2, column=ci, value=col)
        c.fill = styles['fill_subheader']
        c.font = styles['font_header']
        c.alignment = styles['align_center']
        c.border = styles['border_thin']
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = col_widths_map.get(col, 13)
    ws.row_dimensions[2].height = 20

    # 데이터 (최대 5000건 제한 - Excel 성능)
    df_limited = df_sorted.head(5000)
    for row_i, (_, row) in enumerate(df_limited.iterrows(), 3):
        risk_val = row['위험등급']
        alt_fill = styles['fill_row_alt'] if row_i % 2 == 0 else styles['fill_row_light']

        for ci, col in enumerate(display_cols, 1):
            val = row[col]
            c = ws.cell(row=row_i, column=ci, value=val)
            c.border = styles['border_thin']

            if col == '위험등급':
                apply_risk_style(c, risk_val, styles, light=False)
                c.alignment = styles['align_center']
            elif col == '명칭':
                c.alignment = styles['align_left']
                c.fill = alt_fill
            else:
                c.alignment = styles['align_center']
                c.fill = alt_fill

            # Z-score 색상 강조
            if col == 'Z-score' and isinstance(val, (int, float)):
                z_abs = abs(val)
                if z_abs > 5:
                    c.fill = styles['fill_danger_lt']
                    c.font = styles['font_danger']
                elif z_abs > 3:
                    c.fill = styles['fill_warning_lt']
                    c.font = styles['font_warning']

        ws.row_dimensions[row_i].height = 16

    if len(df_sorted) > 5000:
        note_row = 2 + len(df_limited) + 1  # 헤더(2행) + 데이터 + 1
        ws.cell(row=note_row, column=1,
                value=f"※ 전체 {len(df_sorted):,}건 중 상위 5,000건만 표시됩니다. 전체 이력은 원본 데이터를 참조하십시오.")


# ==========================================
# 9. Sheet 4: 이동평균 분석 + 시계열 차트 (TOP N 센서)
# ==========================================
def create_moving_average_chart_sheet(wb, df, sensor_map, df_summary, styles, top_n=5):
    """
    위험도 상위 top_n 센서에 대해 원본값 vs 이동평균 시계열을 기록하고
    openpyxl LineChart로 시각화.
    """
    print(f"[4/5] 상위 {top_n}개 센서 시계열 차트 생성 중...")

    ws = wb.create_sheet(title="이동평균 분석(차트)")
    ws.sheet_view.showGridLines = False

    ws.merge_cells('A1:F1')
    c = ws.cell(row=1, column=1,
                value=f"이동평균(MA30) 분석 ─ 위험도 상위 {top_n}개 센서 시계열 차트")
    c.fill = styles['fill_header']
    c.font = styles['font_title']
    c.alignment = styles['align_center']
    ws.row_dimensions[1].height = 28

    MA_WINDOW = 30
    top_sensors = df_summary.head(top_n)

    current_col = 1  # 데이터를 쓸 시작 열

    chart_anchors = []  # (차트 anchor 셀, 차트 객체) 리스트

    times = df['Time'].reset_index(drop=True)

    for sensor_rank, (_, sensor_row) in enumerate(top_sensors.iterrows()):
        col_idx = int(sensor_row['Col_Idx'])
        tag     = sensor_row['iba Tag']
        name    = sensor_row['명칭']
        risk    = sensor_row['위험등급']

        series = df[col_idx].reset_index(drop=True).astype(float)
        ma     = series.rolling(window=MA_WINDOW, min_periods=1).mean()

        # 이 센서의 데이터를 쓸 열 범위: (시각, 원본값, MA)
        col_time = current_col
        col_val  = current_col + 1
        col_ma   = current_col + 2

        # 섹션 제목
        header_row = 3 + sensor_rank * (len(series) + 6)

        ws.merge_cells(
            start_row=header_row, start_column=col_time,
            end_row=header_row,   end_column=col_ma
        )
        c = ws.cell(row=header_row, column=col_time,
                    value=f"[{sensor_rank+1}위] {tag}  |  {name}  |  위험등급: {risk}")
        c.fill = styles['fill_subheader']
        c.font = styles['font_header']
        c.alignment = styles['align_center']
        ws.row_dimensions[header_row].height = 22

        # 컬럼 헤더
        for ci, label in enumerate(['발생 시각', '원본값', f'이동평균(MA{MA_WINDOW})'], col_time):
            c = ws.cell(row=header_row + 1, column=ci, value=label)
            c.fill = styles['fill_header']
            c.font = styles['font_header']
            c.alignment = styles['align_center']
            c.border = styles['border_thin']
            ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = 22
        ws.row_dimensions[header_row + 1].height = 18

        # 데이터 쓰기
        data_start_row = header_row + 2
        for i, (t, v, m) in enumerate(zip(times, series, ma)):
            r = data_start_row + i
            ws.cell(row=r, column=col_time, value=str(t)).alignment = styles['align_center']
            ws.cell(row=r, column=col_val,  value=round(float(v), 2)).alignment = styles['align_center']
            ws.cell(row=r, column=col_ma,   value=round(float(m), 2)).alignment = styles['align_center']
            ws.row_dimensions[r].height = 14

        data_end_row = data_start_row + len(series) - 1

        # LineChart 생성
        chart = LineChart()
        chart.title      = f"{tag} │ {name[:30]}  (위험등급: {risk})"
        chart.style      = 10
        chart.y_axis.title = "센서값"
        chart.x_axis.title = "시간 인덱스"
        chart.width  = 22
        chart.height = 12

        # 원본값 시리즈
        ref_val = Reference(ws, min_col=col_val, min_row=data_start_row,
                            max_row=data_end_row)
        chart.add_data(ref_val, titles_from_data=False)
        chart.series[0].title = SeriesLabel(v="원본값")
        chart.series[0].graphicalProperties.line.solidFill = "FF0000" if risk == "위험" else (
            "FFC000" if risk == "주의" else "4472C4")
        chart.series[0].graphicalProperties.line.width = 15000  # 1.5pt

        # 이동평균 시리즈
        ref_ma = Reference(ws, min_col=col_ma, min_row=data_start_row,
                           max_row=data_end_row)
        chart.add_data(ref_ma, titles_from_data=False)
        chart.series[1].title = SeriesLabel(v=f"이동평균(MA{MA_WINDOW})")
        chart.series[1].graphicalProperties.line.solidFill = "595959"
        chart.series[1].graphicalProperties.line.width = 20000  # 2pt (굵게)

        # 차트 삽입 위치 (데이터 옆 열)
        anchor_col = col_ma + 2
        anchor_cell = f"{openpyxl.utils.get_column_letter(anchor_col)}{header_row}"
        chart_anchors.append((anchor_cell, chart))

    # 차트 워크시트에 삽입
    for anchor, chart in chart_anchors:
        ws.add_chart(chart, anchor)

    # 열 너비 조정
    for ci in range(1, current_col + 1):
        col_letter = openpyxl.utils.get_column_letter(ci)
        if ws.column_dimensions[col_letter].width < 10:
            ws.column_dimensions[col_letter].width = 14


# ==========================================
# 10. Sheet 5: 이상 발생 패턴 분석 (전반부/후반부 비교)
# ==========================================
def create_pattern_analysis_sheet(wb, df_summary, styles):
    ws = wb.create_sheet(title="이상발생 패턴분석")
    ws.sheet_view.showGridLines = False

    ws.merge_cells('A1:I1')
    c = ws.cell(row=1, column=1,
                value="이상 발생 패턴 분석  ─  측정 전반부(1~720분) vs 후반부(721~1440분) 이상 집중도 비교")
    c.fill = styles['fill_header']
    c.font = styles['font_title']
    c.alignment = styles['align_center']
    ws.row_dimensions[1].height = 28

    ws.merge_cells('A2:I2')
    c = ws.cell(row=2, column=1,
                value="※ 후반부에 이상이 집중될수록 시간 경과에 따른 설비 열화·고장 진행을 의미합니다.")
    c.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type='solid')
    c.font = Font(size=9, italic=True, color="7F4F00")
    c.alignment = styles['align_center']
    ws.row_dimensions[2].height = 16

    headers = [
        "iba Tag", "센서 명칭", "위험등급", "전체 이상건수",
        "전반부 이상", "후반부 이상",
        "후반부 집중도(%)", "시계열 추세", "분석 의견"
    ]
    col_widths_p = [14, 40, 10, 14, 12, 12, 16, 13, 50]

    for ci, (h, w) in enumerate(zip(headers, col_widths_p), 1):
        c = ws.cell(row=3, column=ci, value=h)
        c.fill = styles['fill_subheader']
        c.font = styles['font_header']
        c.alignment = styles['align_center']
        c.border = styles['border_thin']
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w
    ws.row_dimensions[3].height = 20

    top50 = df_summary.head(50)  # 상위 50개 센서만 표시
    for row_i, (_, row) in enumerate(top50.iterrows(), 4):
        risk_val    = row['위험등급']
        total_anom  = row['이상징후 건수']
        early       = row['전반부 이상건수']
        late        = row['후반부 이상건수']
        late_ratio  = round(late / (total_anom + 1e-9) * 100, 1)
        trend_val   = row['시계열 추세']

        # 분석 의견
        if late_ratio > 70:
            opinion = f"후반부 집중({late_ratio}%) → 측정 기간 후반에 설비 상태 급격 악화. 즉각 점검 권고."
        elif late_ratio > 55:
            opinion = f"후반부 증가({late_ratio}%) → 시간 경과에 따른 성능 저하 진행 중. 추이 모니터링."
        elif late_ratio < 30:
            opinion = f"전반부 집중({100-late_ratio:.0f}%) → 초기 기동 시 이상 발생 패턴. 기동 절차 점검 권고."
        else:
            opinion = f"균등 분포(전{100-late_ratio:.0f}%/후{late_ratio}%) → 특정 시간대 편향 없음. 전반적 모니터링 유지."

        alt_fill = styles['fill_row_alt'] if row_i % 2 == 0 else styles['fill_row_light']
        row_vals = [
            row['iba Tag'], row['명칭'], risk_val, total_anom,
            early, late, late_ratio, trend_val, opinion
        ]
        for ci, val in enumerate(row_vals, 1):
            c = ws.cell(row=row_i, column=ci, value=val)
            c.border = styles['border_thin']
            if ci == 3:
                apply_risk_style(c, risk_val, styles, light=False)
                c.alignment = styles['align_center']
            elif ci == 7:  # 후반부 집중도
                if late_ratio > 70:
                    c.fill = styles['fill_danger_lt']
                    c.font = styles['font_danger']
                elif late_ratio > 55:
                    c.fill = styles['fill_warning_lt']
                    c.font = styles['font_warning']
                else:
                    c.fill = alt_fill
                c.alignment = styles['align_center']
            elif ci == 9:
                c.alignment = styles['align_top']
                c.fill = alt_fill
            else:
                c.alignment = styles['align_center']
                c.fill = alt_fill
        ws.row_dimensions[row_i].height = 28


# ==========================================
# 11. Sheet 6: 메트릭 정의 (고도화)
# ==========================================
def create_metrics_definition_sheet(wb, styles):
    ws = wb.create_sheet(title="지표 정의 및 해석 가이드")
    ws.sheet_view.showGridLines = False

    ws.merge_cells('A1:F1')
    c = ws.cell(row=1, column=1, value="지표 정의 및 해석 가이드  ─  현장 엔지니어 참조용")
    c.fill = styles['fill_header']
    c.font = styles['font_title']
    c.alignment = styles['align_center']
    ws.row_dimensions[1].height = 28

    definition_data = [
        ["지표명", "계산 방법", "정상 범위 기준", "위험 신호 해석", "현장 조치 가이드"],
        [
            "IF 심각도 점수\n(Isolation Forest)",
            "Isolation Forest의 decision_function 절댓값 평균.\n"
            "0에 가까울수록 정상, 값이 클수록 이상.",
            "0 ~ 0.08 : 정상\n0.08 ~ 0.15 : 주의\n0.15 초과 : 위험",
            "값이 높을수록 해당 센서가 다른 센서들과 매우 다른\n"
            "패턴을 보임. 전체 데이터 상위 1%(contamination=0.01)\n"
            "를 이상치로 판단.",
            "IF 심각도 0.15 초과 → 즉시 현장 확인.\n"
            "다른 인근 센서와 값 비교하여 설비 전체 이상 여부 판단."
        ],
        [
            "최대 Z-score\n(표준 편차 기반)",
            "Z = (측정값 - 전체평균) / 전체표준편차\n"
            "이상치 포인트 중 Z 절댓값이 가장 큰 값.",
            "Z < 3σ : 정상 범위\nZ 3~5σ : 주의\nZ > 5σ : 위험",
            "Z-score = 3 → 정상 분포 가정 시 약 0.27% 확률.\n"
            "Z-score = 5 → 약 0.00006% 확률의 극단 이탈.\n"
            "센서 오류 또는 실제 설비 고장일 가능성 매우 높음.",
            "Z > 5σ : 센서 교정 또는 히터 소자 즉시 점검.\n"
            "Z 3~5σ : 해당 시간대 운전 로그와 대조 분석."
        ],
        [
            "이동평균 이탈량\n(MA30 기준)",
            "MA30 = 직전 30개 포인트 이동 평균.\n"
            "이탈량 = 실제값 - MA30.",
            "이탈량이 작을수록 안정적 운전 상태.",
            "이탈량이 클수록 단기간에 급격한 변동 발생.\n"
            "이동평균 대비 지속적 상방 이탈 → 과열 위험.\n"
            "지속적 하방 이탈 → 히터 열화·단선 의심.",
            "이탈량 > 설정온도의 10% : 현장 점검 권고.\n"
            "MA와 실제값 간 지속적 괴리 → 제어 루프(PID) 점검."
        ],
        [
            "연속이상 최대구간\n(건)",
            "Isolation Forest가 이상(-1)로 판정한 데이터 포인트 중\n"
            "연속으로 이어진 최대 길이.",
            "1~3건 : 일시적 노이즈 가능성\n"
            "4~10건 : 주의 관찰\n"
            "10건 초과 : 지속 이상",
            "연속이상이 길수록 일회성 스파이크가 아닌\n"
            "실제 설비 이상 가능성이 높음.\n"
            "20건 초과는 설비 고장 진행 중을 강력히 시사.",
            "연속 20건 초과 → 운전 정지 후 설비 점검 검토.\n"
            "연속 10~20건 → 다음 정기점검 시 최우선 확인 대상."
        ],
        [
            "이상 비율(%)",
            "이상치 건수 / 전체 데이터 건수 × 100.\n"
            "전체 측정 기간 중 이상 상태 점유 비율.",
            "0 ~ 1% : 정상 범위\n1 ~ 2% : 관찰\n2% 초과 : 주의\n5% 초과 : 위험",
            "이상 비율이 높을수록 설비가 지속적으로 비정상\n"
            "상태에서 운전 중임을 의미.\n"
            "10% 초과는 센서 이상 또는 구조적 설비 문제 의심.",
            "이상 비율 5% 초과 → 해당 센서 보정 및 설비 전수 점검.\n"
            "이상 비율 10% 초과 → 센서 교체 및 설비 정밀 진단."
        ],
        [
            "시계열 추세\n(선형 회귀 기울기)",
            "전체 측정 기간의 선형 회귀 기울기(slope)를\n"
            "평균값 대비 정규화하여 방향 판정.\n"
            "rel_slope = slope / |mean|",
            "rel_slope ∈ (-0.0001, 0.0001) : 안정(─)\n"
            "rel_slope > 0.0001 : 상승추세(▲)\n"
            "rel_slope < -0.0001 : 하강추세(▼)",
            "상승추세: 온도·압력 등이 시간에 따라 점진적 증가 →\n"
            "부품 마모·열화 또는 냉각 성능 저하 의심.\n"
            "하강추세: 출력 저하, 히터 성능 감소 의심.",
            "상승추세 + 위험등급 → 즉시 열화 원인 분석.\n"
            "하강추세 → 히터 전력 공급 및 소자 저항값 측정 권고."
        ],
        [
            "전반부/후반부\n이상건수 비교",
            "전반부: 측정 기간 1~720분 내 이상 건수.\n"
            "후반부: 측정 기간 721~1440분 내 이상 건수.",
            "균등 분포(약 50:50)가 이상적.",
            "후반부 이상 집중 → 시간 경과에 따른 설비 상태\n"
            "악화 진행 중. 가장 위험한 패턴.\n"
            "전반부 집중 → 기동 초기 불안정 패턴.",
            "후반부 집중(70% 이상) → 즉각 점검 및\n"
            "다음 PM 시 해당 설비 최우선 점검 대상 지정."
        ],
        [
            "정상 대비 변화율(%)",
            "abs(이상구간 평균 - 전체 평균) / |전체 평균| × 100.\n"
            "분모 0 방지를 위해 1e-9 보정.",
            "10% 미만 : 소폭 이탈\n10~30% : 중간 이탈\n30% 초과 : 대폭 이탈",
            "변화율이 높을수록 정상 운전 범위에서 크게 벗어남.\n"
            "단, 정상 평균이 0에 가까운 센서는 변화율이 과도하게\n"
            "크게 나타날 수 있으므로 Z-score와 병행 해석 필요.",
            "변화율 30% 초과 AND Z-score 3 초과 →\n"
            "단순 변동이 아닌 실제 이상으로 판단, 즉시 점검."
        ],
    ]

    col_widths_def = [22, 38, 28, 45, 48]
    fill_row_colors = ["DEEAF1", "F2F7FB"]

    for ci, w in enumerate(col_widths_def, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w

    for row_idx, row_data in enumerate(definition_data, 2):
        for col_idx, val in enumerate(row_data, 1):
            c = ws.cell(row=row_idx, column=col_idx, value=val)
            c.border = styles['border_thin']
            c.alignment = Alignment(wrap_text=True, vertical='top')

            if row_idx == 2:
                c.fill = styles['fill_subheader']
                c.font = styles['font_header']
                c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            else:
                fill_color = fill_row_colors[(row_idx - 3) % 2]
                c.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
                if col_idx == 1:
                    c.font = Font(bold=True)

    ws.row_dimensions[2].height = 20
    for i in range(3, len(definition_data) + 2):
        ws.row_dimensions[i].height = 68


# ==========================================
# 12. AI 진단 실행 (상위 10개 센서)
# ==========================================
def run_ai_diagnosis(df_summary):
    print("[3/5] AI(GPT) 정밀 진단 수행 중 (상위 10개 센서)...")
    targets = df_summary.head(10)

    for idx in targets.index:
        row = df_summary.loc[idx]
        print(f"   > AI 분석: {row['명칭'][:40]} (위험등급: {row['위험등급']})")

        prompt = f"""
[역할]
당신은 포스코 퓨처엠 소성로 설비 예지보전 전문가입니다.
현장 엔지니어가 엑셀에서 바로 읽고 즉각 조치할 수 있도록 작성하십시오.

[분석 데이터]
- 센서 명칭: {row['명칭']}
- 위험 등급: {row['위험등급']}
- 전체 이상치 건수: {row['이상징후 건수']}건 (전체 대비 {row['이상 비율(%)']}%)
- 정상 평균: {row['정상 평균']}  /  이상구간 평균: {row['이상구간 평균']}
- 정상 대비 변화율: {row['정상 대비 변화율(%)']}%
- 최대 Z-score: {row['최대 Z-score']}σ
- 이동평균 최대이탈: {row['이동평균 최대이탈']}
- 연속 이상 최대: {row['연속이상 최대구간(건)']}건
- 패턴 방향: {row['패턴 방향']}  /  시계열 추세: {row['시계열 추세']}
- 전반부 이상: {row['전반부 이상건수']}건  /  후반부 이상: {row['후반부 이상건수']}건

[출력 규칙 – 반드시 준수]
1. 아래 형식만 사용, 항목당 1~2문장 이내
2. 고장 위험도는 반드시 % 숫자로 명시
3. 서론·결론·불필요한 설명 추가 금지

[출력 형식]
① 이상 분석: [관측된 이상 패턴 핵심 요약]
② 근거: [수치 근거 – 정상 대비 편차, Z-score, 연속이상 등]
③ 추정 원인: [히터 단선·과부하·센서 오류 등 구체적 원인 추정]
④ 고장 위험도: [XX%] – [위험 등급 근거 요약]
⑤ 즉시 조치: [현장 담당자가 바로 실행할 수 있는 구체적 행동]
"""
        ai_result = call_corporate_ai_api(prompt)
        df_summary.loc[idx, 'AI 위험도 판단'] = ai_result

    return df_summary


# ==========================================
# 13. 최종 엑셀 리포트 생성
# ==========================================
def create_smart_report(df, sensor_map, df_summary, df_details, output_path):
    print(f"[5/5] 엑셀 리포트 저장 중: {output_path}")

    # AI 진단 (API key 있을 때만)
    if POSCO_GPT_KEY:
        df_summary = run_ai_diagnosis(df_summary)

    styles = get_styles()
    wb = openpyxl.Workbook()

    # Sheet 1: 종합 요약 대시보드
    create_dashboard_sheet(wb, df_summary, styles)

    # Sheet 2: 센서별 상세 분석표
    create_sensor_detail_sheet(wb, df_summary, styles)

    # Sheet 3: 이상치 상세 이력
    create_anomaly_timeline_sheet(wb, df_details, styles)

    # Sheet 4: 이동평균 분석 + 차트
    create_moving_average_chart_sheet(wb, df, sensor_map, df_summary, styles, top_n=5)

    # Sheet 5: 이상 발생 패턴 분석
    create_pattern_analysis_sheet(wb, df_summary, styles)

    # Sheet 6: 지표 정의 및 해석 가이드
    create_metrics_definition_sheet(wb, styles)

    wb.save(output_path)
    print(f"   >> 저장 완료: {output_path}")


# ==========================================
# 14. 메인 실행 함수 (GUI)
# ==========================================
def main():
    print("=== [POSCO Future M] AI 예지보전 시스템 가동 ===")

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    if not POSCO_GPT_KEY:
        messagebox.showwarning(
            "키 확인",
            "환경변수 POSCO_GPT_KEY가 설정되지 않았습니다.\n"
            "AI 진단 기능은 비활성화되며, 나머지 분석은 정상 수행됩니다."
        )

    while True:
        file_path = filedialog.askopenfilename(
            parent=root,
            title="분석할 센서 데이터 파일(CSV/TXT) 선택",
            filetypes=[("CSV/TXT Files", "*.csv *.txt"), ("All Files", "*.*")]
        )

        if not file_path:
            print(">> 종료합니다.")
            break

        try:
            df, sensor_map = load_data(file_path)

            # 분석 전 유효 센서 사전 필터링 (전체 0 또는 상수 센서 제외 → 성능 최적화)
            print("[전처리] 유효 센서 필터링 중...")
            import numpy as _np
            valid_sensor_map = {}
            for _ci, _meta in sensor_map.items():
                _s = df[_ci].astype(float)
                if _s.std() < 1e-9:
                    continue
                _uv = _s.dropna().unique()
                if len(_uv) <= 2 and set(_np.round(_uv)).issubset({0.0, 1.0}):
                    continue
                valid_sensor_map[_ci] = _meta
            print(f"   >> 전체 {len(sensor_map)}개 센서 중 유효 {len(valid_sensor_map)}개 선별")

            df_result, df_details = compute_enhanced_metrics(df, valid_sensor_map)

            if df_result.empty:
                messagebox.showinfo("결과", "분석 가능한 아날로그 신호에서 특이점이 발견되지 않았습니다.")
            else:
                danger_n  = int((df_result['위험등급'] == '위험').sum())
                warning_n = int((df_result['위험등급'] == '주의').sum())
                ok_n      = int((df_result['위험등급'] == '정상').sum())

                timestamp  = datetime.now().strftime('%Y%m%d_%H%M%S')
                base_name  = os.path.basename(file_path).rsplit('.', 1)[0]
                output_file = f"SmartReport_{base_name}_{timestamp}.xlsx"

                create_smart_report(df, sensor_map, df_result, df_details, output_file)

                msg = (
                    f"분석 완료!\n\n"
                    f"  위험  센서: {danger_n}개\n"
                    f"  주의  센서: {warning_n}개\n"
                    f"  정상  센서: {ok_n}개\n\n"
                    f"저장 위치: {output_file}\n\n"
                    f"계속 분석하시겠습니까?"
                )
                if not messagebox.askyesno("완료", msg, parent=root):
                    break

        except Exception as e:
            import traceback
            err_detail = traceback.format_exc()
            print(f"[Error]\n{err_detail}")
            messagebox.showerror("오류", f"처리 중 오류가 발생했습니다:\n{str(e)}", parent=root)

    root.destroy()


if __name__ == "__main__":
    main()
