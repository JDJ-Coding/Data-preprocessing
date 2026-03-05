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
        
        # AI가 분석가처럼 행동하도록 온도(Temperature)를 낮춰서 설정
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
    
    # 다양한 인코딩과 구분자로 읽기 시도
    for sep in separators:
        for enc in encodings:
            try:
                df_raw = pd.read_csv(path, sep=sep, header=None, encoding=enc, low_memory=False)
                if df_raw.shape[1] > 1: # 컬럼이 1개 이상이어야 정상
                    break
                else:
                    df_raw = None 
            except: continue
        if df_raw is not None: break

    if df_raw is None or len(df_raw) < 3:
        raise ValueError("파일을 읽을 수 없습니다. (형식이 올바르지 않음)")

    # 1행: Tag, 2행: Name
    tags = df_raw.iloc[0].astype(str).str.strip()
    names = df_raw.iloc[1].astype(str).str.strip()
    
    # 시간 컬럼 찾기
    time_col_idx = 0
    for i, val in enumerate(names):
        if 'time' in val.lower() or 'date' in val.lower():
            time_col_idx = i; break
            
    sensor_map = {}
    valid_indices = []
    
    # 센서 매핑 정보 구축
    for i in range(len(df_raw.columns)):
        if i == time_col_idx: continue
        # 태그나 이름이 비어있으면 스킵
        if not tags.iloc[i] or tags.iloc[i] == 'nan': continue
        
        sensor_map[i] = {'tag': tags.iloc[i], 'name': names.iloc[i]}
        valid_indices.append(i)

    # 실제 데이터 부분 (3행부터)
    df_data = df_raw.iloc[2:].copy()
    df_data.rename(columns={time_col_idx: 'Time'}, inplace=True)
    
    # 숫자 변환 및 결측치 보정
    for i in valid_indices:
        df_data[i] = pd.to_numeric(df_data[i], errors='coerce')
        # 선형 보간 후 앞뒤 채우기
        df_data[i] = df_data[i].interpolate(method='linear').fillna(method='bfill').fillna(0)

    return df_data, sensor_map

# ==========================================
# 2. 머신러닝 분석 (넓은 그물망)
# ==========================================
def analyze_with_ml(df, sensor_map):
    print("[2/4] 머신러닝 분석 중 (모든 이상 징후 탐색)...")
    results = []
    
    # 오염도 0.01 = 상위 1%의 특이값은 무조건 잡겠다
    clf = IsolationForest(contamination=0.01, random_state=42, n_jobs=-1)
    
    for col_idx, meta in sensor_map.items():
        series = df[col_idx]
        
        # 값이 완전히 고정된 경우(표준편차 0)는 분석 가치 없음
        if series.std() == 0: continue

        # 디지털 신호(0, 1만 있는 경우) 제외 -> 아날로그만 분석
        unique_vals = np.unique(series.dropna())
        if set(unique_vals).issubset({0, 1, 0.0, 1.0}):
            continue

        # 학습 및 예측
        X = series.values.reshape(-1, 1)
        clf.fit(X)
        preds = clf.predict(X)
        
        # -1이 이상치(Anomaly)
        outlier_idx = np.where(preds == -1)[0]
        count = len(outlier_idx)
        
        if count > 0:
            anomalies = series.iloc[outlier_idx]
            normal_mean = series.mean()
            anomaly_mean = anomalies.mean()
            
            # 변화율 계산 (절대값 기준)
            change_ratio = (abs(anomaly_mean - normal_mean) / (abs(normal_mean) + 1e-9)) * 100
            
            direction = "상승(▲)" if anomaly_mean > normal_mean else "하강(▼)"
            # 심각도 Score: 결정 경계에서 얼마나 멀리 떨어져 있는가
            score = abs(clf.decision_function(X)[outlier_idx].mean())
            
            results.append({
                'Col_Idx': col_idx,
                'iba Tag': meta['tag'],
                '명칭': meta['name'],
                '이상징후 건수': count,
                '패턴 유형': direction,
                '정상 평균': round(normal_mean, 2),
                '이상 구간 평균': round(anomaly_mean, 2),
                '변화율(%)': round(change_ratio, 2), # % 단위
                'Max': round(anomalies.max(), 2),
                'Min': round(anomalies.min(), 2),
                '심각도': score
            })
            
    if not results:
        return pd.DataFrame()
        
    # 심각도 순으로 정렬해서 반환
    return pd.DataFrame(results).sort_values(by='심각도', ascending=False)

# ==========================================
# 3. 엑셀 리포트 생성 & AI 판단 (정밀 거름망)
# ==========================================
def create_smart_report(df_summary, output_path):
    print("[3/4] AI(GPT) 정밀 진단 수행 중...")
    
    df_summary['AI 위험도 판단'] = "분석 대기"
    
    # 상위 10개(또는 전체)에 대해 AI에게 물어봄
    targets = df_summary.head(10) 
    
    for idx in targets.index:
        row = df_summary.loc[idx]
        print(f"   > Analyzing: {row['명칭']} (변화율: {row['변화율(%)']}%)")
        
        # ★ [핵심] AI에게 '판단'을 시키는 프롬프트
        prompt = f"""
        [상황]
        공장 설비 센서 데이터에서 머신러닝 모델이 평소와 다른 패턴을 감지했습니다.
        이것이 단순 노이즈인지, 실제 설비 고장의 전조인지 판단해 주십시오.

        [센서 정보]
        - 센서명: {row['명칭']} (Tag: {row['iba Tag']})
        - 정상 평균: {row['정상 평균']}
        - 현재 이상값 평균: {row['이상 구간 평균']} ({row['변화율(%)']}% {row['패턴 유형']})
        - 이상 발생 빈도: 전체 데이터 중 {row['이상징후 건수']}건 감지됨

        [판단 기준]
        1. 변화율이 1~2% 내외로 미미하다면 'Low' (단, 압력/진동 등 민감 센서는 예외)
        2. 수치가 급격히 튀었다면 'High'
        3. 센서 이름을 보고 설비 특성을 유추하여 위험도를 평가하시오.

        [필수 출력 포맷]
        - 위험도: [Low / Medium / High] 중 택1
        - 진단: (한 줄 요약)
        - 조치: (엔지니어에게 권장하는 행동)
        """
        
        ai_result = call_corporate_ai_api(prompt)
        df_summary.loc[idx, 'AI 위험도 판단'] = ai_result

    print(f"[4/4] 엑셀 파일 저장 중: {output_path}")
    
    # 엑셀 저장
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "AI Smart Report"
    
    # 스타일 정의
    fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type='solid')
    fill_high = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type='solid') # 빨강 (High)
    font_header = Font(color="FFFFFF", bold=True)
    border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # 헤더 쓰기
    cols = [c for c in df_summary.columns if c != 'Col_Idx']
    for i, col in enumerate(cols):
        cell = ws.cell(row=1, column=i+1, value=col)
        cell.fill = fill_header
        cell.font = font_header
        cell.border = border_thin
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # 데이터 쓰기
    for r_idx, row in enumerate(dataframe_to_rows(df_summary[cols], index=False, header=False), 2):
        for c_idx, val in enumerate(row):
            cell = ws.cell(row=r_idx, column=c_idx+1, value=val)
            cell.border = border_thin
            
            # AI 판단 컬럼(마지막)은 줄바꿈 허용 및 왼쪽 정렬
            if c_idx == len(cols) - 1:
                cell.alignment = Alignment(wrap_text=True, vertical='top')
                # 위험도가 High면 빨간색 표시
                if "High" in str(val):
                    cell.fill = fill_high
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')

    # 컬럼 너비 자동 조정 (대략)
    ws.column_dimensions['C'].width = 30 # 명칭
    ws.column_dimensions[openpyxl.utils.get_column_letter(len(cols))].width = 60 # AI 결과

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
            # 1. 데이터 로드
            df, sensor_map = load_data(file_path)
            
            # 2. 머신러닝 분석
            df_result = analyze_with_ml(df, sensor_map)
            
            if df_result.empty:
                messagebox.showinfo("결과", "분석 가능한 아날로그 신호에서 특이점이 발견되지 않았습니다.")
            else:
                # 3. 리포트 생성 (AI 판단 포함)
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
