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
# [핵심] 사내 AI API 연동 함수 (디버깅 강화 & 표준 준수)
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
                {"role": "system", "content": "당신은 포스코 퓨처엠의 설비 예지보전 전문가입니다."},
                {"role": "user", "content": prompt_text}
            ],
            "frequency_penalty": 0.0, 
            "temperature": 0.7 
        }

        # print(f"   >> [API] 서버 요청 중... (Model: gpt-5.2)")
        
        # 요청 전송
        response = requests.post(url, headers=headers, json=payload, timeout=60)
        
        # ★ [핵심 수정 1] 한글 깨짐 방지 (UTF-8 강제 지정)
        response.encoding = 'utf-8'
        
        if response.status_code == 200:
            try:
                # 1. JSON으로 시도
                result = response.json()
                if 'choices' in result:
                    return result['choices'][0]['message']['content'].strip()
                elif 'response' in result:
                    return result['response']
                else:
                    return str(result)
            except json.JSONDecodeError:
                # ★ [핵심 수정 2] JSON이 아니면, 받은 텍스트 자체가 정답입니다.
                # 깨진 문자가 이제 정상적인 한글로 보일 것입니다.
                return response.text.strip()
        else:
            return f"[API Error] 상태코드: {response.status_code}, 메시지: {response.text[:100]}"

    except Exception as e:
        return f"[시스템 오류] {str(e)}"


# ==========================================
# 1. 데이터 로드
# ==========================================
def load_data(path):
    print(f"[1/4] 데이터 로딩: {os.path.basename(path)}")
    encodings = ['cp949', 'euc-kr', 'utf-8']
    df_raw = None
    
    separators = [';', ','] 
    
    for sep in separators:
        for enc in encodings:
            try:
                df_raw = pd.read_csv(path, sep=sep, header=None, encoding=enc, low_memory=False)
                if df_raw.shape[1] > 1:
                    break
                else:
                    df_raw = None 
            except: continue
        if df_raw is not None: break

    if df_raw is None or len(df_raw) < 3:
        raise ValueError("파일을 읽을 수 없습니다. (구분자 또는 인코딩 문제)")

    tags = df_raw.iloc[0].astype(str).str.strip()
    names = df_raw.iloc[1].astype(str).str.strip()
    
    time_col_idx = 0
    for i, val in enumerate(names):
        if 'time' in val.lower() or 'date' in val.lower():
            time_col_idx = i; break
            
    sensor_map = {}
    valid_indices = []
    for i in range(len(df_raw.columns)):
        if i == time_col_idx: continue
        sensor_map[i] = {'tag': tags.iloc[i], 'name': names.iloc[i]}
        valid_indices.append(i)

    df_data = df_raw.iloc[2:].copy()
    time_col = 'Time'
    df_data.rename(columns={time_col_idx: time_col}, inplace=True)
    
    for i in valid_indices:
        df_data[i] = pd.to_numeric(df_data[i], errors='coerce')
        df_data[i] = df_data[i].interpolate(method='linear').fillna(method='bfill').fillna(0)

    return df_data, sensor_map, time_col

# ==========================================
# 2. 머신러닝 분석 (디지털 값 필터링 적용)
# ==========================================
def analyze_with_ml(df, sensor_map):
    print("[2/4] 머신러닝 분석 중 (디지털 신호 제외)...")
    results = []
    clf = IsolationForest(contamination=0.01, random_state=42, n_jobs=-1)
    
    for col_idx, meta in sensor_map.items():
        series = df[col_idx]
        
        # 1. 값이 변화가 없는 경우 제외
        if series.std() == 0: continue

        # ★ [추가된 로직] 디지털 신호(0, 1) 필터링 ★
        # 고유값이 {0, 1}의 부분집합이면 디지털 신호로 간주하고 Skip
        unique_vals = np.unique(series.dropna())
        if set(unique_vals).issubset({0, 1, 0.0, 1.0}):
            # print(f"   -> [Skip] 디지털 신호: {meta['name']}") 
            continue

        # 아날로그 데이터 분석 시작
        X = series.values.reshape(-1, 1)
        clf.fit(X)
        preds = clf.predict(X)
        
        outlier_idx = np.where(preds == -1)[0]
        count = len(outlier_idx)
        
        if count > 0:
            anomalies = series.iloc[outlier_idx]
            normal_mean = series.mean()
            anomaly_mean = anomalies.mean()
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
                'Max': round(anomalies.max(), 2),
                'Min': round(anomalies.min(), 2),
                '심각도': score
            })
            
    return pd.DataFrame(results).sort_values(by='심각도', ascending=False)

# ==========================================
# 3. 엑셀 리포트 생성
# ==========================================
def create_smart_report(df_summary, output_path):
    print("[3/4] AI 전문가(GPT-5.2) 분석 요청 중 (상위 5개 항목)...")
    
    df_summary['AI 전문가 소견'] = ""
    targets = df_summary.head(5)
    
    for idx, row in targets.iterrows():
        print(f"   > '{row['명칭']}' 정밀 분석 중...")
        
        prompt = f"""
        [설비 아날로그 센서 데이터 정밀 분석 요청]
        
        센서 정보:
        - 이름: {row['명칭']} (Tag: {row['iba Tag']})
        - 상태: 평소 평균 {row['정상 평균']} 대비 {row['패턴 유형']}하는 이상 패턴(Anomaly) 감지됨.
        - 이상치 값: 평균 {row['이상 구간 평균']} (최대 {row['Max']}, 최소 {row['Min']})
        
        요청 사항:
        1. 이 아날로그 값의 비정상적 변화가 의미하는 설비의 잠재적 결함(베어링 파손, 윤활 부족, 정렬 불량 등)은 무엇일 가능성이 높은가?
        2. 현장 엔지니어가 즉시 확인해야 할 점검 포인트 2가지.
        """
        
        ai_opinion = call_corporate_ai_api(prompt)
        df_summary.loc[idx, 'AI 전문가 소견'] = ai_opinion

    print(f"[4/4] 엑셀 리포트 저장 중: {output_path}")
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "예지보전 리포트"
    
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type='solid')
    font_header = Font(name='맑은 고딕', bold=True, color="FFFFFF")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    cols = [c for c in df_summary.columns if c != 'Col_Idx']
    
    for i, col in enumerate(cols):
        cell = ws.cell(row=1, column=i+1, value=col)
        cell.fill = header_fill
        cell.font = font_header
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border
    
    for r, row in enumerate(dataframe_to_rows(df_summary[cols], index=False, header=False), 2):
        for c, val in enumerate(row):
            cell = ws.cell(row=r, column=c+1, value=val)
            cell.border = border
            if c == len(cols)-1:
                cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['K'].width = 60
    
    wb.save(output_path)
    return True

# ==========================================
# 4. 메인 실행 함수
# ==========================================
def main():
    print("=== [POSCO Future M] 스마트 예지보전 시스템 (Analog Only) ===")
    
    root = tk.Tk()
    root.withdraw() 
    root.attributes('-topmost', True)

    if not POSCO_GPT_KEY:
        print("[Warning] 환경변수 POSCO_GPT_KEY가 없습니다. AI 기능이 작동하지 않을 수 있습니다.")
        messagebox.showwarning("키 확인", "환경변수 POSCO_GPT_KEY를 찾을 수 없습니다.")

    while True:
        print("\n>> 파일 선택 창이 뜹니다. 파일을 선택해주세요...")
        
        file_path = filedialog.askopenfilename(
            parent=root,
            title="분석할 데이터 파일 선택", 
            filetypes=[("Data Files", "*.csv *.txt"), ("All Files", "*.*")]
        )
        
        if not file_path:
            print(">> 종료합니다.")
            break
        
        try:
            df, sensor_map, time_col = load_data(file_path)
            
            # 분석 수행 (디지털 필터링 포함)
            df_result = analyze_with_ml(df, sensor_map)
            
            if df_result.empty:
                messagebox.showinfo("결과", "분석 대상인 아날로그 신호에서 특이한 이상 징후가 없습니다.\n(디지털 신호는 자동 제외되었습니다.)", parent=root)
            else:
                timestamp = datetime.now().strftime('%H%M%S')
                base_name = os.path.basename(file_path).rsplit('.', 1)[0]
                output_file = f"Report_{base_name}_{timestamp}.xlsx"
                
                create_smart_report(df_result, output_file)
                
                print(f">> 완료! 저장된 파일: {output_file}")
                
                if not messagebox.askyesno("완료", f"분석 완료!\n파일: {output_file}\n\n다른 파일을 계속 분석하시겠습니까?", parent=root):
                    break
                    
        except Exception as e:
            print(f"[Error] {e}")
            messagebox.showerror("오류", f"오류 발생:\n{str(e)}", parent=root)
            break

    root.destroy()

if __name__ == "__main__":
    main()
