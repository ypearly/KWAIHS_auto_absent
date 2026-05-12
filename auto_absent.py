import pandas as pd
import win32com.client as win32
import os
import holidays
from datetime import timedelta, datetime

# ==========================================
# 1. 초기 설정 및 사용자 입력
# ==========================================
print("="*50)
print("광운AI고 출결 신고서 자동 생성 프로그램 (주말/공휴일 건너뛰기 통합 버전)")
teacher_name = input("담임 교사 성함을 입력하세요: ")
extra_input = input("이번 달 재량휴업일(mm-dd, mm-dd 형식 / 없으면 엔터): ")
print("="*50)

# 공휴일 및 재량휴업일 설정
kr_holidays = holidays.KR()
school_holidays = []
if extra_input.strip():
    raw_dates = [d.strip() for d in extra_input.split(',')]
    for rd in raw_dates:
        try:
            full_date = f"2026-{rd}"
            datetime.strptime(full_date, '%Y-%m-%d')
            school_holidays.append(full_date)
        except ValueError:
            print(f"⚠ 날짜 형식이 잘못되었습니다: {rd}")

base_path = os.getcwd() 
excel_file = os.path.join(base_path, '월별 출결 현황.xlsx')
hwp_template = os.path.join(base_path, '2026학년도 결석신고서.hwp')
output_base_dir = os.path.join(base_path, 'output')

# ==========================================
# 2. 데이터 로드 및 전처리 (지능형 그룹화 로직)
# ==========================================
try:
    df = pd.read_excel(excel_file)
except Exception as e:
    print(f"❌ 엑셀 파일을 읽을 수 없습니다: {e}")
    exit()

def clean_date(date_str):
    if pd.isna(date_str): return date_str
    parts = str(date_str).split('.')
    return ".".join(parts[:3])

df['일자_clean'] = df['일자'].apply(clean_date)
df['일자_dt'] = pd.to_datetime(df['일자_clean'], format='%Y.%m.%d')
df = df.sort_values(by=['번호', '일자_dt']).reset_index(drop=True)

def is_business_day(date):
    """주말, 공휴일, 재량휴업일 여부 확인"""
    if date.weekday() >= 5: return False
    if date in kr_holidays: return False
    if date.strftime('%Y-%m-%d') in school_holidays: return False
    return True

# [지능형 그룹화] 주말/공휴일을 건너뛰고 실제 수업일 기준 연속성 판단
df['group'] = 0
if not df.empty:
    group_id = 1
    df.loc[0, 'group'] = group_id
    
    for i in range(1, len(df)):
        prev = df.iloc[i-1]
        curr = df.iloc[i]
        
        # 기본 조건: 번호, 사유, 출결구분이 동일해야 함
        same_info = (curr['번호'] == prev['번호']) and \
                    (curr['사유'] == prev['사유']) and \
                    (curr['출결구분'] == prev['출결구분'])
        
        if same_info:
            # 날짜 차이 계산
            date_diff = (curr['일자_dt'] - prev['일자_dt']).days
            
            # 1일 차이면 당연히 연속
            if date_diff == 1:
                df.loc[i, 'group'] = group_id
            else:
                # 1일 초과 차이일 때, 그 사이 날짜들이 모두 휴일인지 확인
                gap_days = [prev['일자_dt'] + timedelta(days=d) for d in range(1, date_diff)]
                all_holidays = all(not is_business_day(d) for d in gap_days)
                
                if all_holidays: # 사이 날짜가 전부 휴일이면 같은 그룹
                    df.loc[i, 'group'] = group_id
                else: # 수업일이 끼어있으면 새로운 그룹
                    group_id += 1
                    df.loc[i, 'group'] = group_id
        else:
            group_id += 1
            df.loc[i, 'group'] = group_id

# ==========================================
# 3. 메인 실행 루프
# ==========================================
try:
    hwp = win32.dynamic.Dispatch("HWPFrame.HwpObject")
    hwp.XHwpWindows.Item(0).Visible = True 

    def get_next_business_day(current_date):
        next_day = current_date + timedelta(days=1)
        while not is_business_day(next_day):
            next_day += timedelta(days=1)
        return next_day

    for group_id, target_group in df.groupby('group'):
        student = target_group.iloc[0]
        status = str(student['출결구분'])
        if "미인정" in status: continue

        start_date = target_group['일자_dt'].min()
        end_date = target_group['일자_dt'].max()
        total_days = len(target_group)
        submit_date = get_next_business_day(end_date)
        
        month_folder_name = f"{start_date.month:02d}월"
        month_dir = os.path.join(output_base_dir, month_folder_name)
        if not os.path.exists(month_dir): os.makedirs(month_dir)

        hwp.Open(os.path.abspath(hwp_template), "HWP", "")
        
        def fill(field, value):
            if pd.notna(value): hwp.PutFieldText(field, str(value))

        fill("number", student['번호'])
        fill("name", student['성명'])
        fill("st-name", student['성명'])
        fill("reason", student['사유'])
        fill("t-name", teacher_name)
        fill("start-month", start_date.month); fill("start-day", start_date.day)
        fill("end-month", end_date.month); fill("end-day", end_date.day)
        fill("cal-day", total_days)
        fill("submit-month", submit_date.month); fill("submit-day", submit_date.day)

        # 출결 상태 코드 및 마킹
        status_code = "1" if "결석" in status else ("2" if "지각" in status else ("3" if "조퇴" in status else "4"))
        attendance_type = "결석" if "결석" in status else ("지각" if "지각" in status else ("조퇴" if "조퇴" in status else "결과"))
        type_code = "1" if "질병" in status else ("2" if "출석인정" in status else "3")

        fill(f"sort{status_code}-{type_code}", "V")
        fill("sort-select", attendance_type)

        # 상세 사유 및 의견란
        reason_val = str(student['사유']).strip() if pd.notna(student['사유']) else ""
        opinion_text = f"위 학생의 {reason_val} 사유를 확인하였으며 {attendance_type}을/를 지도함."
        if type_code == "1" or (type_code == "2" and any(kw in reason_val for kw in ["훈련", "연습경기", "리그", "전반기리그"])):
            opinion_text = f"위 학생의 {reason_val} 사유를 확인하였으며 학생 및 학부모 상담을 통해 안정과 치료를 목적으로 {attendance_type}을/를 지도함."

        if type_code == "2":
            target_perm = ""
            if any(kw in reason_val for kw in ["특별교육", "연습경기", "훈련", "리그", "전반기리그"]):
                target_perm = "permission7"
                fill("permission-etc", reason_val)
                fill(f"permission-etc7-{status_code}", "V")
            elif any(kw in reason_val for kw in ["경조사", "상", "부친", "모친"]): target_perm = "permission1"
            elif any(kw in reason_val for kw in ["군특성화", "군특성", "행사"]): target_perm = "permission2"
            elif any(kw in reason_val for kw in ["대회", "축구부", "경기"]): target_perm = "permission3"
            elif "체험학습" in reason_val: target_perm = "permission4"
            elif any(kw in reason_val for kw in ["자격증", "시험"]): target_perm = "permission5"
            elif any(kw in reason_val for kw in ["면접", "대학", "회사"]): target_perm = "permission6"

            if target_perm and target_perm != "permission7":
                fill(target_perm, "V")
                fill(f"{target_perm}-{status_code}", "V")

        fill("teacher", opinion_text)

        # 파일명 (기간 표시)
        date_str = start_date.strftime('%m%d') if total_days == 1 else f"{start_date.strftime('%m%d')}-{end_date.strftime('%m%d')}"
        file_base_name = f"{int(student['번호']):02d}번_{student['성명']}_{date_str}_{attendance_type}"
        
        hwp.SaveAs(os.path.abspath(os.path.join(month_dir, f"{file_base_name}.hwp")), "HWP", "")
        hwp.SaveAs(os.path.abspath(os.path.join(month_dir, f"{file_base_name}.pdf")), "PDF", "")
        hwp.Run("FileClose") 
        print(f"✔ {file_base_name} ({total_days}일분) 생성 완료")

    print(f"\n✨ 모든 작업 완료!")

except Exception as e:
    print(f"❌ 오류 발생: {e}")
