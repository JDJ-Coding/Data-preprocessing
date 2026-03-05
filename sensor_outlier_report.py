import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
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
    print(f"[1/4] 데이터 로딩 중: {os.path.basename(path)}")
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
        # ✅ 수정: fillna(method='bfill') → bfill() (pandas 2.x 호환)
        df_data[i] = df_data[i].interpolate(method='linear').bfill().fillna(0)

    return df_data, sensor_map

# ==========================================
# 2. 머신러닝 분석 (넓은 그물망)
# ==========================================
def analyze_with_ml(df, sensor_map):
    print("[2/4] 머신러닝 분석 중 (모든 이상 징후 탐색)...")
    results = []

    clf = IsolationForest(contamination=0.01, random_state=42, n_jobs=-1)

    for col_idx, meta in sensor_map.items():
        series = df[col_idx]

        if series.std() == 0:
            continue

        unique_vals = np.unique(series.dropna())
        if set(unique_vals).issubset({0, 1, 0.0, 1.0}):
            continue

        X = series.values.reshape(-1, 1)
        clf.fit(X)
        preds = clf.predict(X)

        outlier_idx = np.where(preds == -1)[0]
        count = len(outlier_idx)

        if count > 0:
            anomalies = series.iloc[outlier_idx]
            normal_mean = series.mean()
            anomaly_mean = anomalies.mean()

            change_ratio = (abs(anomaly_mean - normal_mean) / (abs(normal_mean) + 1e-9)) * 100
            direction = "상승(▲)" if anomaly_mean > normal_mean else "하강(▼)"
            score = abs(clf.decision_function(X)[outlier_idx].mean())

            results.append({
                'Col_Idx': col_idx,
                'iba Tag': meta['tag'],
                '명칭': meta['name'],
                '이상징후 건수': count,
                '패턴 유형': direction,
                '정상 평균': round(normal_mean, 2),
                '이상 구간 평균': round(anomaly_mean, 2),
                '변화율(%)': round(change_ratio, 2),
                'Max': round(anomalies.max(), 2),
                'Min': round(anomalies.min(), 2),
                '심각도': score
            })

    if not results:
        return pd.DataFrame()

    return pd.DataFrame(results).sort_values(by='심각도', ascending=False)

# ==========================================
# ✅ 추가: 메트릭 정의 시트 생성 함수
# ==========================================
def create_metrics_definition_sheet(wb):
    ws_def = wb.create_sheet(title="메트릭 정의")

    definition_data = [
        ["헤더명", "영문 필드명", "데이터 출처", "계산 공식", "의미 및 해석", "주의사항"],
        [
            "이상징후 건수", "count",
            "Isolation Forest 모델이 이상치(-1)로 판단한 데이터 포인트 개수",
            "len(np.where(preds == -1)[0])",
            "해당 센서에서 감지된 이상 데이터 포인트의 총 개수. 값이 클수록 이상 발생 빈도가 높음",
            "contamination=0.01 → 전체 데이터의 상위 1%를 이상치로 가정"
        ],
        [
            "패턴 유형", "direction",
            "이상 구간 평균과 정상 평균을 비교한 방향성",
            "'상승(▲)' if anomaly_mean > normal_mean else '하강(▼)'",
            "이상치가 정상 대비 높은 값(상승) 또는 낮은 값(하락)인지 방향성 표시",
            "단순 평균 비교이므로, 센서 특성에 따라 의미가 다를 수 있음"
        ],
        [
            "정상 평균", "normal_mean",
            "전체 센서 시계열 데이터의 산술평균",
            "series.mean()",
            "센서의 정상 작동 시 기준값. 전체 데이터의 평균으로 정상 상태를 대표",
            "표준편차가 0이거나 0/1만 있는 디지털 센서는 분석에서 제외됨"
        ],
        [
            "이상 구간 평균", "anomaly_mean",
            "Isolation Forest가 이상치(-1)로 판정한 데이터 포인트만의 평균값",
            "series.iloc[outlier_idx].mean()",
            "이상 발생 시점의 평균값. 정상 평균과 비교하여 이상 정도 파악",
            "정상 평균이 0에 가까우면 변화율이 과도하게 클 수 있음"
        ],
        [
            "변화율(%)", "change_ratio",
            "정상 평균 대비 이상 구간 평균의 차이 비율",
            "abs(anomaly_mean - normal_mean) / (abs(normal_mean) + 1e-9) * 100",
            "정상 대비 이상치가 얼마나 벗어났는지 백분율로 표시. 높을수록 이상 정도가 큼",
            "분모가 0이 되는 것을 막기 위해 1e-9(매우 작은 수)를 더해 보정"
        ],
        [
            "Max", "anomalies.max()",
            "이상치로 판정된 데이터 포인트 중 최댓값",
            "series.iloc[outlier_idx].max()",
            "이상 구간에서 기록된 최고값. 센서 값의 상한 이탈 정도 파악",
            "순간 스파이크(노이즈)도 포함될 수 있음"
        ],
        [
            "Min", "anomalies.min()",
            "이상치로 판정된 데이터 포인트 중 최솟값",
            "series.iloc[outlier_idx].min()",
            "이상 구간에서 기록된 최저값. 센서 값의 하한 이탈 정도 파악",
            "순간 드롭(노이즈)도 포함될 수 있음"
        ],
        [
            "심각도", "score",
            "Isolation Forest의 decision_function 결과 절댓값 평균",
            "abs(clf.decision_function(X)[outlier_idx].mean())",
            "머신러닝 모델이 계산한 이상도 점수. 높을수록 정상 패턴에서 크게 벗어남. 결과 정렬 기준",
            "decision_function은 음수일수록 이상치에 가까움. 절댓값으로 변환하여 사용"
        ],
    ]

    fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type='solid')
    font_header = Font(color="FFFFFF", bold=True)
    border_thin = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    fill_row = PatternFill(start_color="DEEAF1", end_color="DEEAF1", fill_type='solid')

    col_widths = [15, 18, 40, 48, 60, 45]

    for row_idx, row_data in enumerate(definition_data, 1):
        for col_idx, val in enumerate(row_data, 1):
            cell = ws_def.cell(row=row_idx, column=col_idx, value=val)
            cell.border = border_thin
            cell.alignment = Alignment(wrap_text=True, vertical='top')

            if row_idx == 1:
                cell.fill = fill_header
                cell.font = font_header
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            elif row_idx % 2 == 0:
                cell.fill = fill_row

    for i, width in enumerate(col_widths, 1):
        ws_def.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width

    ws_def.row_dimensions[1].height = 20
    for i in range(2, len(definition_data) + 1):
        ws_def.row_dimensions[i].height = 55

# ==========================================
# 3. 엑셀 리포트 생성 & AI 판단
# ==========================================
def create_smart_report(df_summary, output_path):
    print("[3/4] AI(GPT) 정밀 진단 수행 중...")

    df_summary['AI 위험도 판단'] = "분석 대기"

    targets = df_summary.head(10)

    for idx in targets.index:
        row = df_summary.loc[idx]
        print(f" > Analyzing: {row['명칭']} (변화율: {row['변화율(%)']})%)")

        prompt = f"""
[역할]
당신은 포스코 퓨처엠 설비 예지보전 전문가입니다.
현장 엔지니어가 엑셀에서 바로 읽을 수 있도록 매우 간결하게 작성하십시오.

[분석 데이터]
- 전체 데이터 개수: {len(df_summary)}
- 이상치 개수: {row['이상징후 건수']}
- 정상 평균: {row['정상 평균']}
- 이상 구간 평균: {row['이상 구간 평균']}
- 변화율: {row['변화율(%)']}%
- 패턴 유형: {row['패턴 유형']}

[출력 규칙 – 매우 중요]
1. 반드시 아래 형식만 사용하십시오.
2. 각 항목은 한 문장만 작성하십시오.
3. 불필요한 설명, 서론, 결론, 목록 추가 금지.
4. 고장 위험도는 반드시 % 숫자로 명시하십시오.

[출력 형식]
1. 데이터 분석결과 : 전체 XXX개 중 이상치 XX개 검출
2. 분석결과 근거 : 정상 평균 XXX, 이상 평균 XXX로 평균 편차 발생
3. 최종결과 : 데이터 상태가 ○○ 패턴으로 고장 위험도 XX%
"""

        ai_result = call_corporate_ai_api(prompt)
        df_summary.loc[idx, 'AI 위험도 판단'] = ai_result

    print(f"[4/4] 엑셀 파일 저장 중: {output_path}")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "AI Smart Report"

    fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type='solid')
    fill_high = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type='solid')
    font_header = Font(color="FFFFFF", bold=True)
    border_thin = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    cols = [c for c in df_summary.columns if c != 'Col_Idx']
    for i, col in enumerate(cols):
        cell = ws.cell(row=1, column=i+1, value=col)
        cell.fill = fill_header
        cell.font = font_header
        cell.border = border_thin
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for r_idx, row in enumerate(dataframe_to_rows(df_summary[cols], index=False, header=False), 2):
        for c_idx, val in enumerate(row):
            cell = ws.cell(row=r_idx, column=c_idx+1, value=val)
            cell.border = border_thin

            if c_idx == len(cols) - 1:
                cell.alignment = Alignment(wrap_text=True, vertical='top')
                if "High" in str(val):
                    cell.fill = fill_high
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')

    ws.column_dimensions['C'].width = 30
    ws.column_dimensions[openpyxl.utils.get_column_letter(len(cols))].width = 60

    # ✅ 추가: 메트릭 정의 시트 생성
    create_metrics_definition_sheet(wb)

    wb.save(output_path)

# ==========================================
# 4. 메인 실행 함수 (GUI)
# ==========================================
def main():
    print("=== [POSCO Future M] AI 예지보전 시스템 가동 ===")

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    if not POSCO_GPT_KEY:
        messagebox.showwarning("키 확인", "환경변수 POSCO_GPT_KEY가 설정되지 않았습니다.\nAI 기능이 제한될 수 있습니다.")

    while True:
        file_path = filedialog.askopenfilename(
            parent=root,
            title="분석할 센서 데이터 파일(CSV) 선택",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )

        if not file_path:
            print(">> 종료합니다.")
            break

        try:
            df, sensor_map = load_data(file_path)
            df_result = analyze_with_ml(df, sensor_map)

            if df_result.empty:
                messagebox.showinfo("결과", "분석 가능한 아날로그 신호에서 특이점이 발견되지 않았습니다.")
            else:
                timestamp = datetime.now().strftime('%H%M%S')
                base_name = os.path.basename(file_path).rsplit('.', 1)[0]
                output_file = f"SmartReport_{base_name}_{timestamp}.xlsx"

                create_smart_report(df_result, output_file)

                msg = f"분석 완료!\n\n저장 위치: {output_file}\n\n계속 분석하시겠습니까?"
                if not messagebox.askyesno("완료", msg, parent=root):
                    break

        except Exception as e:
            print(f"[Error] {e}")
            messagebox.showerror("오류", f"처리 중 오류가 발생했습니다:\n{str(e)}", parent=root)

    root.destroy()

if __name__ == "__main__":
    main()