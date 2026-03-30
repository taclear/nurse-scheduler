import streamlit as st
import io
import nurse_scheduler_complete as ns 

st.set_page_config(page_title="간호사 스마트 스케줄러", page_icon="👩‍⚕️", layout="centered")

st.title("👩‍⚕️ 병동 간호사 스마트 스케줄러 AI")
st.markdown("수간호사 선생님의 스트레스를 0으로! 엑셀 파일을 올리면 AI가 최적의 근무표를 짜줍니다.")
st.divider()

# 👇 1. 빈 양식 다운로드 기능
st.markdown("### 1️⃣ 먼저 빈 양식을 다운로드하세요 (엑셀이 없는 경우)")
try:
    with open("nurse_scheduler_complete.xlsx", "rb") as template_file:
        st.download_button(
            label="📝 빈 근무표 엑셀 양식 다운로드",
            data=template_file,
            file_name="기본_근무표_양식.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
except FileNotFoundError:
    st.info("💡 앗! 관리자가 서버에 기본 양식을 아직 안 올려뒀네요. 깃허브에 원본 엑셀 파일을 업로드해주세요!")

st.markdown("### 2️⃣ 작성한 엑셀 파일을 업로드하세요")
# 👇 2. 파일 업로드 칸
uploaded_file = st.file_uploader("다운받은 양식을 작성한 뒤 여기에 올려주세요!", type=['xlsx'])

# 👇 3. AI 스케줄러 실행 버튼
if st.button("🚀 AI 근무표 자동 생성 시작", use_container_width=True):
    if uploaded_file is not None:
        with st.spinner("AI가 수백만 개의 경우의 수를 계산 중입니다... ⏳ (최대 5분 소요)"):
            try:
                temp_filename = "temp_input.xlsx"
                with open(temp_filename, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                wb, cfg, nurses, is_senior, is_nk, max_offs, holidays, no_night, off_requests, prefs, prev_shifts, preceptors, is_ghost = ns.load_inputs(temp_filename)
                
                precheck_issues = ns.preliminary_checks(cfg, nurses, holidays, no_night, off_requests)
                if precheck_issues:
                    st.error("🚨 스케줄을 짜기 전 문제점이 발견되었습니다!")
                    for issue in precheck_issues:
                        st.warning(issue)
                else:
                    model, x, pref_miss = ns.build_model(cfg, nurses, is_senior, is_nk, max_offs, holidays, no_night, off_requests, prefs, prev_shifts, preceptors, is_ghost)
                    solver, status = ns.solve_model(model, cfg)

                    if status == ns.cp_model.INFEASIBLE:
                        st.error("🚨 에러: 현재 설정된 규칙들끼리 서로 충돌하여 근무표를 완성할 수 없습니다. (인원수나 오프를 확인해주세요)")
                    elif status == ns.cp_model.UNKNOWN:
                        st.error("🚨 시간 초과: 컴퓨터가 정답을 찾기엔 규칙이 너무 복잡합니다. (Setup 시트의 시간을 늘려주세요)")
                    elif status in (ns.cp_model.OPTIMAL, ns.cp_model.FEASIBLE):
                        schedule = ns.extract_schedule(solver, x, nurses, cfg)
                        wb = ns.write_outputs(wb, cfg, nurses, holidays, no_night, off_requests, prefs, schedule, pref_miss, is_ghost)
                        
                        # 🌟 수정된 기능: 필요한 시트 4개만 남기고 다 지우기
                        sheets_to_keep = ["Schedule", "RequestsSummary", "Validation", "DailyCoverage"]
                        for sheet_name in wb.sheetnames:
                            if sheet_name not in sheets_to_keep:
                                del wb[sheet_name]
                        
                        output_filename = "result_schedule.xlsx"
                        wb.save(output_filename)
                        
                        st.balloons()
                        st.success(f"✅ 최적의 근무표 생성 완료! (AI가 받은 벌점: {solver.ObjectiveValue()}점)")
                        
                        with open(output_filename, "rb") as f:
                            st.download_button(
                                label="📥 완성된 스케줄 엑셀 다운로드 (보고서 포함)",
                                data=f,
                                file_name="완성된_병동근무표.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
            except Exception as e:
                st.error(f"🚨 엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
    else:
        st.error("🚨 먼저 엑셀 파일을 업로드해 주세요!")
