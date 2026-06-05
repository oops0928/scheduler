import streamlit as st
import pandas as pd
import calendar
import random
import io

st.set_page_config(page_title="매장 스케줄러", layout="wide", initial_sidebar_state="expanded")

WEEKDAYS_KR = ["월", "화", "수", "목", "금", "토", "일"]
# 셀 선택 옵션: 근무(=일함), 휴무(=기본 휴무일에 포함), 연차(=기본 휴무일에 추가)
OPT_WORK = "근무"
OPT_OFF = "휴무"
OPT_LEAVE = "연차"
CELL_OPTIONS = [OPT_WORK, OPT_OFF, OPT_LEAVE]


def make_dates(year, month):
    num_days = calendar.monthrange(year, month)[1]
    dates = []
    for d in range(1, num_days + 1):
        wd = calendar.weekday(year, month, d)
        dates.append(f"{d}일({WEEKDAYS_KR[wd]})")
    return dates


def generate_schedule(year, month, employees, mandatory_dict, off_target_dict,
                      min_workers, close_workers_count):
    """
    mandatory_dict[emp] = {date_str: "휴무" or "연차"}  (꼭 쉬어야 하는 날)
    off_target_dict[emp] = int  (한 달 기본 휴무 일수)
    """
    dates = make_dates(year, month)

    # 직원별 OFF 날짜 집합과 종류 기록
    off_set = {emp: set() for emp in employees}
    off_type = {emp: {} for emp in employees}  # date -> 휴무 / 연차 / 랜덤

    warnings = []

    # --- 1. 필수 휴무일 먼저 고정 ---
    for emp in employees:
        for d, t in mandatory_dict[emp].items():
            off_set[emp].add(d)
            off_type[emp][d] = t

    # --- 2. 남은 랜덤 휴무 일수 계산 ---
    # '휴무'로 표시한 필수일은 기본 일수에 포함 / '연차'는 추가(기본 일수를 줄이지 않음)
    remaining = {}
    for emp in employees:
        counted = sum(1 for t in mandatory_dict[emp].values() if t == OPT_OFF)
        remaining[emp] = max(0, off_target_dict[emp] - counted)

    # --- 하루 최대 휴무 가능 인원 (최소 근무 인원 보장) ---
    max_off_per_day = max(0, len(employees) - min_workers)
    day_off_count = {d: 0 for d in dates}
    for emp in employees:
        for d in off_set[emp]:
            day_off_count[d] += 1

    # 필수 휴무만으로 이미 최소 근무 인원을 못 맞추는 날 경고
    for d in dates:
        if day_off_count[d] > max_off_per_day:
            warnings.append(
                f"⚠️ {d}: 필수 휴무가 몰려 최소 근무 인원({min_workers}명)을 못 맞춥니다. "
                f"(이 날 쉬는 사람 {day_off_count[d]}명)"
            )

    # --- 3. 남은 휴무를 공평하게 랜덤 배분 ---
    emp_order = employees[:]
    random.shuffle(emp_order)
    for emp in emp_order:
        need = remaining[emp]
        assigned = 0
        while assigned < need:
            # 후보: 이 직원이 아직 안 쉬는 날 + 그날 휴무 정원이 남은 날
            candidates = [
                d for d in dates
                if d not in off_set[emp] and day_off_count[d] < max_off_per_day
            ]
            if not candidates:
                break
            # 그날 쉬는 사람이 가장 적은 날부터 (몰림 방지), 동률이면 랜덤
            min_c = min(day_off_count[d] for d in candidates)
            best = [d for d in candidates if day_off_count[d] == min_c]
            chosen = random.choice(best)
            off_set[emp].add(chosen)
            off_type[emp][chosen] = "랜덤"
            day_off_count[chosen] += 1
            assigned += 1

        if assigned < need:
            warnings.append(
                f"⚠️ {emp}: 목표 휴무 일수를 다 못 채웠습니다. "
                f"(부족 {need - assigned}일 — 인원/조건이 빡빡합니다)"
            )

    # --- 4. 근무자 오픈/마감 배정 + 결과 만들기 ---
    schedule_data = []
    shift_counts = {emp: {'A': 0, 'B': 0, 'C': 0, 'OFF': 0, 'LEAVE': 0} for emp in employees}

    for day in range(1, len(dates) + 1):
        d = dates[day - 1]
        wd_idx = calendar.weekday(year, month, day)
        is_weekend = (wd_idx >= 4)  # 금/토/일 마감조 = C
        close_shift_name = 'C' if is_weekend else 'B'

        working = [emp for emp in employees if d not in off_set[emp]]
        off = [emp for emp in employees if d in off_set[emp]]

        # 휴무 통계
        off_labels = []
        for emp in off:
            if off_type[emp].get(d) == OPT_LEAVE:
                shift_counts[emp]['LEAVE'] += 1
                off_labels.append(f"{emp}(연차)")
            else:
                shift_counts[emp]['OFF'] += 1
                off_labels.append(emp)

        # 근무 인원 부족 경고
        if len(working) < min_workers:
            warnings.append(
                f"⚠️ {d} 근무 인원 {len(working)}명 (권장 최소 {min_workers}명 미달)"
            )

        # 마감조 배정
        if len(working) >= close_workers_count:
            working.sort(key=lambda e: shift_counts[e]['B'] + shift_counts[e]['C'])
            pool_size = min(len(working), close_workers_count + 2)
            top = working[:pool_size]
            random.shuffle(top)
            close_workers = top[:close_workers_count]
            open_workers = [e for e in working if e not in close_workers]
        else:
            close_workers = working
            open_workers = []

        for emp in close_workers:
            shift_counts[emp][close_shift_name] += 1
        for emp in open_workers:
            shift_counts[emp]['A'] += 1

        schedule_data.append({
            "날짜": d,
            "A조 (오픈)": ", ".join(open_workers) if open_workers else "-",
            f"{close_shift_name}조 (마감)": ", ".join(close_workers) if close_workers else "-",
            "OFF (휴무)": ", ".join(off_labels) if off_labels else "-",
        })

    # 중복 경고 제거(순서 유지)
    warnings = list(dict.fromkeys(warnings))
    return pd.DataFrame(schedule_data), shift_counts, warnings


def to_excel(df_schedule, df_stats):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_schedule.to_excel(writer, index=False, sheet_name='스케줄표')
        df_stats.to_excel(writer, index=False, sheet_name='근무통계')
    return output.getvalue()


# ==========================================
# 🎨 웹앱 UI
# ==========================================
st.title("📅 매장 자동 스케줄러")

with st.sidebar:
    st.header("⚙️ 기본 설정")
    year = st.selectbox("연도", [2025, 2026, 2027], index=1)
    month = st.selectbox("월", list(range(1, 13)), index=5)

    st.markdown("---")
    emp_input = st.text_input("직원 이름 (쉼표로 구분)", "박미진, 김혜영, 황별이, 장재원, 최지호")
    employees = [e.strip() for e in emp_input.split(",") if e.strip()]

    st.markdown("---")
    st.header("🔧 근무 조건")
    default_off = st.number_input("월 기본 휴무 일수 (전원 공통)", min_value=0, max_value=31, value=8)
    min_workers = st.number_input("하루 최소 근무 인원", min_value=1, max_value=20, value=3)
    close_workers_count = st.number_input("마감조(B/C) 필요 인원", min_value=1, max_value=20, value=2)

    # 직원별 휴무 일수 다르게(파트타임 등) — 필요할 때만
    with st.expander("👥 직원별 휴무 일수 따로 설정 (선택)"):
        st.caption("비워두면 위의 '월 기본 휴무 일수'를 그대로 사용합니다.")
        off_target_dict = {}
        for emp in employees:
            off_target_dict[emp] = st.number_input(
                f"{emp}", min_value=0, max_value=31, value=int(default_off), key=f"off_{emp}"
            )

dates = make_dates(year, month)

# 세션 상태: 직원/월/연도 바뀌면 입력표 초기화
sig = (tuple(employees), year, month)
if "sig" not in st.session_state or st.session_state.sig != sig:
    st.session_state.df_input = pd.DataFrame(OPT_WORK, index=dates, columns=employees)
    st.session_state.sig = sig

st.header("🏝️ 꼭 쉬어야 하는 날 입력")
st.info(
    "각 직원이 **반드시 쉬어야 하는 날**만 선택하세요. "
    f"`{OPT_OFF}`=기본 휴무 일수에 포함 / `{OPT_LEAVE}`=기본 휴무에 추가(연차). "
    "나머지 휴무는 버튼을 누르면 시스템이 공평하게 자동 배분합니다."
)

# 헤더에 선택한 필수 휴무 개수 표시
column_config = {}
for emp in employees:
    cnt = int((st.session_state.df_input[emp] != OPT_WORK).sum())
    column_config[emp] = st.column_config.SelectboxColumn(
        f"👤 {emp} (필수 {cnt}일)",
        options=CELL_OPTIONS,
        required=True,
        width="small",
    )

generate_clicked = st.button("🚀 스케줄 자동 생성하기", use_container_width=True, type="primary")

edited_df = st.data_editor(
    st.session_state.df_input,
    use_container_width=True,
    height=400,
    column_config=column_config,
    key="data_editor_state",
)

if not edited_df.equals(st.session_state.df_input):
    st.session_state.df_input = edited_df
    st.rerun()

st.markdown("---")

if generate_clicked:
    # 입력표 -> 필수 휴무 dict
    mandatory_dict = {emp: {} for emp in employees}
    for date in dates:
        for emp in employees:
            val = edited_df.at[date, emp]
            if val in (OPT_OFF, OPT_LEAVE):
                mandatory_dict[emp][date] = val

    with st.spinner('스케줄을 짜는 중입니다...'):
        df_schedule, shift_counts, warnings = generate_schedule(
            year, month, employees, mandatory_dict, off_target_dict,
            min_workers, close_workers_count
        )

        stat_data = []
        for emp, c in shift_counts.items():
            total_rest = c['OFF'] + c['LEAVE']
            stat_data.append({
                "이름": emp,
                "목표 휴무": off_target_dict[emp],
                "실제 휴무": c['OFF'],
                "연차": c['LEAVE'],
                "총 쉬는 날": total_rest,
                "A조 (오픈)": c['A'],
                "B조 (평일마감)": c['B'],
                "C조 (주말마감)": c['C'],
            })
        df_stats = pd.DataFrame(stat_data)

        st.session_state.schedule_result = df_schedule
        st.session_state.stats_result = df_stats
        st.session_state.warnings = warnings

if "schedule_result" in st.session_state:
    st.header("📊 생성된 스케줄 결과")

    if st.session_state.warnings:
        for w in st.session_state.warnings:
            st.warning(w)
    else:
        st.success("✅ 조건 충돌 없이 스케줄이 완성되었습니다!")

    excel_data = to_excel(st.session_state.schedule_result, st.session_state.stats_result)
    st.download_button(
        label="📥 엑셀 파일(.xlsx)로 다운로드",
        data=excel_data,
        file_name=f"{year}년_{month}월_스케줄표.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    tab1, tab2 = st.tabs(["🗓️ 스케줄표 보기", "📊 개인별 근무 통계"])
    with tab1:
        st.dataframe(st.session_state.schedule_result, use_container_width=True, hide_index=True)
    with tab2:
        st.dataframe(st.session_state.stats_result, use_container_width=True, hide_index=True)
