import streamlit as st
import pandas as pd
import calendar
import random
import io

st.set_page_config(page_title="매장 스케줄러", layout="wide", initial_sidebar_state="expanded")

def generate_schedule(year, month, employees, off_days_dict, min_workers, close_workers_count):
    num_days = calendar.monthrange(year, month)[1]
    weekdays_kr = ["월", "화", "수", "목", "금", "토", "일"]
    
    schedule_data = []
    shift_counts = {emp: {'A': 0, 'B': 0, 'C': 0, 'OFF': 0} for emp in employees}
    warnings = []

    for day in range(1, num_days + 1):
        wd_idx = calendar.weekday(year, month, day)
        wd_str = weekdays_kr[wd_idx]
        is_weekend = (wd_idx >= 4)
        close_shift_name = 'C' if is_weekend else 'B'
        
        working_emps = []
        off_emps = []
        
        target_date_str = f"{day}일({wd_str})"
        
        for emp in employees:
            if target_date_str in off_days_dict[emp]:
                off_emps.append(emp)
                shift_counts[emp]['OFF'] += 1
            else:
                working_emps.append(emp)
                
        # ⚠️ 설정한 '최소 근무 인원'보다 적을 때 경고
        if len(working_emps) < min_workers:
            warnings.append(f"⚠️ {day}일({wd_str}) 근무 인원이 {len(working_emps)}명입니다. (권장 최소 인원: {min_workers}명 부족!)")

        # 설정한 '마감조 인원' 이상이 출근한 경우 정상 배정
        if len(working_emps) >= close_workers_count:
            # 마감 횟수가 적은 사람을 우선순위로 정렬
            working_emps.sort(key=lambda emp: shift_counts[emp]['B'] + shift_counts[emp]['C'])
            
            # 동점자 랜덤화를 위해 필요 인원보다 조금 더 많은 후보군 선택
            pool_size = min(len(working_emps), close_workers_count + 2)
            top_candidates = working_emps[:pool_size]
            random.shuffle(top_candidates)
            
            # 마감조 배정 (설정한 인원수만큼)
            close_workers = top_candidates[:close_workers_count]
            open_workers = [emp for emp in working_emps if emp not in close_workers]
        else:
            # 출근자가 마감조 인원보다도 적은 비상 상황 시 전원 마감 배정
            close_workers = working_emps
            open_workers = []

        for emp in close_workers:
            shift_counts[emp][close_shift_name] += 1
        for emp in open_workers:
            shift_counts[emp]['A'] += 1

        schedule_data.append({
            "날짜": f"{day}일({wd_str})",
            "A조 (오픈)": ", ".join(open_workers) if open_workers else "-",
            f"{close_shift_name}조 (마감)": ", ".join(close_workers) if close_workers else "-",
            "OFF (휴무)": ", ".join(off_emps) if off_emps else "-"
        })

    return pd.DataFrame(schedule_data), shift_counts, warnings

# 엑셀 변환 함수
def to_excel(df_schedule, df_stats):
    output = io.BytesIO()
    # openpyxl 엔진을 사용하여 메모리 상에서 엑셀 파일 생성
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_schedule.to_excel(writer, index=False, sheet_name='스케줄표')
        df_stats.to_excel(writer, index=False, sheet_name='근무통계')
    processed_data = output.getvalue()
    return processed_data

# ==========================================
# 🎨 웹앱 UI 구성
# ==========================================
st.title("📅 매장 자동 스케줄러")

with st.sidebar:
    st.header("⚙️ 기본 설정")
    year = st.selectbox("연도", [2026, 2027], index=0)
    month = st.selectbox("월", list(range(1, 13)), index=5)
    
    st.markdown("---")
    emp_input = st.text_input("직원 이름 (쉼표로 구분)", "박미진, 김혜영, 황별이, 장재원, 최지호")
    employees = [e.strip() for e in emp_input.split(",") if e.strip()]
    
    st.markdown("---")
    st.header("🔧 근무 조건 설정")
    # 스케줄 로직에 직접 반영되는 변수들
    min_workers = st.number_input("하루 최소 근무 인원", min_value=1, max_value=10, value=3)
    close_workers_count = st.number_input("마감조(B/C) 필요 인원", min_value=1, max_value=10, value=2)

num_days = calendar.monthrange(year, month)[1]
weekdays_kr = ["월", "화", "수", "목", "금", "토", "일"]
dates = [f"{d}일({weekdays_kr[calendar.weekday(year, month, d)]})" for d in range(1, num_days + 1)]

if "df_input" not in st.session_state or len(st.session_state.df_input.columns) != len(employees) or st.session_state.current_month != month:
    st.session_state.df_input = pd.DataFrame(False, index=dates, columns=employees)
    st.session_state.current_month = month

st.header("🏝️ 휴무일 입력 및 스케줄 생성")
st.info("표에서 쉴 날짜를 체크하고 버튼을 누르세요. 직원이름 옆 숫자는 선택한 휴무일 수입니다.")

column_config = {}
for emp in employees:
    checked_count = st.session_state.df_input[emp].sum() 
    column_config[emp] = st.column_config.CheckboxColumn(
        f"👤 {emp} ({checked_count}일)"
    )

generate_clicked = st.button("🚀 스케줄 자동 생성하기", use_container_width=True, type="primary")

edited_df = st.data_editor(
    st.session_state.df_input, 
    use_container_width=True, 
    height=350, 
    column_config=column_config,
    key="data_editor_state"
)

if not edited_df.equals(st.session_state.df_input):
    st.session_state.df_input = edited_df
    st.rerun()

st.markdown("---")

# 스케줄 생성 로직 (결과를 session_state에 저장하여 화면 유지)
if generate_clicked:
    off_days_dict = {emp: [] for emp in employees}
    for date in dates:
        for emp in employees:
            if edited_df.at[date, emp] == True:
                off_days_dict[emp].append(date)

    with st.spinner('스케줄을 짜는 중입니다...'):
        df_schedule, shift_counts, warnings = generate_schedule(
            year, month, employees, off_days_dict, min_workers, close_workers_count
        )
        
        # 개인별 통계 데이터프레임 만들기
        stat_data = []
        for emp, counts in shift_counts.items():
            stat_data.append({
                "이름": emp,
                "A조 (오픈)": counts['A'],
                "B조 (평일마감)": counts['B'],
                "C조 (주말마감)": counts['C'],
                "총 휴무일": counts['OFF']
            })
        df_stats = pd.DataFrame(stat_data)

        # 화면에 결과를 계속 띄워두기 위해 세션에 저장
        st.session_state.schedule_result = df_schedule
        st.session_state.stats_result = df_stats
        st.session_state.warnings = warnings

# 저장된 스케줄 결과가 있다면 화면에 출력
if "schedule_result" in st.session_state:
    st.header("📊 생성된 스케줄 결과")
    
    # 경고 메시지 출력
    if st.session_state.warnings:
        for w in st.session_state.warnings:
            st.error(w)
    else:
        st.success("✅ 조건 충돌 없이 스케줄이 완벽하게 생성되었습니다!")
    
    # 엑셀 다운로드 버튼
    excel_data = to_excel(st.session_state.schedule_result, st.session_state.stats_result)
    st.download_button(
        label="📥 엑셀 파일(.xlsx)로 다운로드",
        data=excel_data,
        file_name=f"{year}년_{month}월_스케줄표.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    # 화면에 표 출력
    tab1, tab2 = st.tabs(["🗓️ 스케줄표 보기", "📊 개인별 근무 통계"])
    with tab1:
        st.dataframe(st.session_state.schedule_result, use_container_width=True, hide_index=True)
    with tab2:
        st.dataframe(st.session_state.stats_result, use_container_width=True, hide_index=True)