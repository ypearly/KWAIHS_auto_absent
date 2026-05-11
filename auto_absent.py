import pandas as pd
import win32com.client as win32
import os
import holidays
from datetime import timedelta, datetime

# ==========================================
# 1. 초기 설정 및 사용자 입력
# ==========================================
print("="*50)
print("광운AI고 출결 신고서 자동 생성 프로그램 (미인정 제외 버전)")
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

# 경로 설정
base_path = os.getcwd() 
excel_file = os.path.join(base_path, '월별 출결 현황.xlsx')
hwp_template = os.path.join(base_path, '2026학년도 결석신고서.hwp')
output_base_dir = os.path.join(base_path, 'output')

# ==========================================
# 2. 데이터 로드 및 전처리
# ==========================================
try:
    df = pd.read_excel(excel_file)
except Exception as e:
    print(f"❌ 엑셀 파일을 읽을 수 없습니다: {e}")
    exit()

# 날짜 정리 및 연속된 출결 그룹화
def clean_date(date_str):
    if pd.isna(date_str): return date_str
    parts = str(date_str).split('.')
    return ".".join(parts[:3])

df['일자_clean'] = df['일자'].apply(clean_date)
df['일자_dt'] = pd.to_datetime(df['일자_clean'], format='%Y.%m.%d')
df = df.sort_values(by=['번호', '일자_dt'])

# 번호, 사유, 출결구분이 같고 날짜가 연속되면 하나의 신고서로 묶음
df['diff'] = df['일자_dt'].diff().dt.days != 1
df['group'] = (df['diff'] | (df['번호'] != df['번호'].shift()) | 
               (df['사유'] != df['사유'].shift()) | 
               (df['출결구분'] != df['출결구분'].shift())).cumsum()

# ==========================================
# 3. 한글 제어 및 메인 루프
# ==========================================
try:
    hwp = win32.dynamic.Dispatch("HWPFrame.HwpObject")
    hwp.XHwpWindows.Item(0).Visible = True 

    def get_next_business_day(current_date):
        next_day = current_date + timedelta(days=1)
        while True:
            if next_day.weekday() >= 5 or next_day in kr_holidays or next_day.strftime('%Y-%m-%d') in school_holidays:
                next_day += timedelta(days=1)
            else: break
        return next_day

    for group_id, target_group in df.groupby('group'):
        student = target_group.iloc[0]
        status = str(student['출결구분'])
        
        # [수정 사항] 미인정 출결 제외 로직
        if "미인정" in status:
            print(f"⏩ [제외] {student['성명']} ({student['일자_clean']}) - {status}은 신고서 대상이 아닙니다.")
            continue

        start_date = target_group['일자_dt'].min()
        end_date = target_group['일자_dt'].max()
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
        fill("start-month", start_date.month)
        fill("start-day", start_date.day)
        fill("end-month", end_date.month)
        fill("end-day", end_date.day)
        fill("cal-day", len(target_group))
        fill("submit-month", submit_date.month)
        fill("submit-day", submit_date.day)

        status_code = "" 
        attendance_type = ""
        
        if "결석" in status: status_code, attendance_type = "1", "결석"
        elif "지각" in status: status_code, attendance_type = "2", "지각"
        elif "조퇴" in status: status_code, attendance_type = "3", "조퇴"
        elif "결과" in status: status_code, attendance_type = "4", "결과"
        
        type_code = "1" if "질병" in status else ("2" if "출석인정" in status else "3")

        fill(f"sort{status_code}-{type_code}", "V")
        fill("sort-select", attendance_type)

        reason_val = str(student['사유']).strip() if pd.notna(student['사유']) else ""
        opinion_text = ""
        base_opinion = f"위 학생의 {reason_val} 사유를 확인하였으며"

        if type_code == "2":
            target_perm = ""
            if any(kw in reason_val for kw in ["특별교육", "연습경기", "훈련", "리그", "전반기리그"]):
                target_perm = "permission7"
                fill("permission-etc", reason_val)
                fill(f"permission-etc7-{status_code}", "V")
                if any(kw in reason_val for kw in ["훈련", "연습경기", "리그", "전반기리그"]):
                    opinion_text = f"{base_opinion} 학생 및 학부모 상담을 통해 안정과 치료를 목적으로 {attendance_type}을/를 지도함."
                else:
                    opinion_text = f"{base_opinion} {attendance_type}을/를 지도함."
            elif any(kw in reason_val for kw in ["경조사", "상", "부친", "모친"]): target_perm = "permission1"
            elif any(kw in reason_val for kw in ["군특성화", "군특성", "행사"]): target_perm = "permission2"
            elif any(kw in reason_val for kw in ["대회", "축구부", "경기"]): target_perm = "permission3"
            elif "체험학습" in reason_val: target_perm = "permission4"
            elif any(kw in reason_val for kw in ["자격증", "시험"]): target_perm = "permission5"
            elif any(kw in reason_val for kw in ["면접", "대학", "회사"]): target_perm = "permission6"

            if target_perm and target_perm != "permission7":
                fill(target_perm, "V")
                fill(f"{target_perm}-{status_code}", "V")
                opinion_text = f"{base_opinion} {attendance_type}을/를 지도함."

        elif type_code == "1":
            opinion_text = f"{base_opinion} 학생 및 학부모 상담을 통해 안정과 치료를 목적으로 {attendance_type}을/를 지도함."
        else:
             opinion_text = f"{base_opinion} {attendance_type}을/를 지도함."

        fill("teacher", opinion_text)

        student_no = f"{int(student['번호']):02d}"
        file_base_name = f"{student_no}번_{student['성명']}_{start_date.strftime('%m%d')}_{attendance_type}"
        
        hwp.SaveAs(os.path.abspath(os.path.join(month_dir, f"{file_base_name}.hwp")), "HWP", "")
        hwp.SaveAs(os.path.abspath(os.path.join(month_dir, f"{file_base_name}.pdf")), "PDF", "")
        
        hwp.Run("FileClose") 
        print(f"✔ [{month_folder_name}] {file_base_name} 생성 완료")

    print(f"\n✨ 모든 작업이 완료되었습니다. 'output' 폴더를 확인하세요.")

except Exception as e:
    print(f"❌ 실행 중 오류 발생: {e}")