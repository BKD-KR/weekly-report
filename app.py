import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pandas as pd
import json
import ssl

# SSL 인증 우회 (회사 방화벽 대응)
ssl._create_default_https_context = ssl._create_unverified_context

# 페이지 설정
st.set_page_config(
    page_title="주간 업무 보고",
    page_icon="📝",
    layout="wide"
)

# Google Sheets 연결
@st.cache_resource
def connect_to_sheets():
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]
    
    # Streamlit secrets에서 인증 정보 가져오기
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    
    # 스프레드시트 열기
    spreadsheet = client.open_by_key(st.secrets["sheet_key"])
    return spreadsheet

# 보고 기간 불러오기
def load_period_settings(spreadsheet):
    try:
        # "설정" 시트를 이름으로 찾기
        settings_sheet = spreadsheet.worksheet("설정")
        data = settings_sheet.get_all_records()
        settings = {}
        for row in data:
            settings[row['항목']] = row['값']
        return settings.get('실적기간', ''), settings.get('계획기간', ''), settings.get('월간보고', '')
    except Exception as e:
        st.error(f"설정 불러오기 실패: {e}")
        return '', '', ''

# 보고 기간 저장하기
def save_period_settings(spreadsheet, result_period, plan_period, monthly_info):
    try:
        # "설정" 시트를 이름으로 찾기
        settings_sheet = spreadsheet.worksheet("설정")
        
        # 기존 데이터 업데이트
        settings_sheet.update('B2', result_period)  # 실적기간
        settings_sheet.update('B3', plan_period)    # 계획기간
        settings_sheet.update('B4', monthly_info)   # 월간보고
        return True
    except Exception as e:
        st.error(f"저장 실패 상세: {str(e)}")
        return False

# 데이터 로드
def load_data(sheet):
    try:
        data = sheet.get_all_records()
        if not data:
            return pd.DataFrame(columns=['id', '작성시간', '팀', '작성자', '구분', '내용'])
        return pd.DataFrame(data)
    except:
        return pd.DataFrame(columns=['id', '작성시간', '팀', '작성자', '구분', '내용'])

# 데이터 저장
def save_data(sheet, report):
    row = [
        report['id'],
        report['작성시간'],
        report['팀'],
        report['작성자'],
        report['구분'],
        report['내용']
    ]
    sheet.append_row(row)

# 데이터 삭제
def delete_data(sheet, report_id):
    all_rows = sheet.get_all_values()
    for idx, row in enumerate(all_rows):
        if idx == 0:  # 헤더 건너뛰기
            continue
        if str(row[0]) == str(report_id):  # id 컬럼 확인
            sheet.delete_rows(idx + 1)
            return True
    return False

# 메인 앱
def main():
    st.title("📝 주간 업무 보고")
    st.caption("개발자: 반경돈")
    
    # Google Sheets 연결
    try:
        spreadsheet = connect_to_sheets()
        sheet = spreadsheet.worksheet("시트1")  # 보고 데이터는 "시트1"에 저장
    except Exception as e:
        st.error(f"Google Sheets 연결 실패: {e}")
        st.info("관리자에게 문의하세요.")
        return
    
    # 보고 기간 불러오기
    saved_result_period, saved_plan_period, saved_monthly_info = load_period_settings(spreadsheet)
    
    # 보고 기간 설정 (관리자용)
    with st.expander("📅 보고 기간 설정 (관리자용)", expanded=False):
        st.warning("⚠️ 이 설정은 모든 사용자에게 적용됩니다!")
        col1, col2 = st.columns(2)
        with col1:
            result_period_input = st.text_input(
                "실적 기간",
                value=saved_result_period if saved_result_period else "1.19. ~ 2.1.(2주)",
                key="result_period_input"
            )
        with col2:
            plan_period_input = st.text_input(
                "계획 기간",
                value=saved_plan_period if saved_plan_period else "2.2. ~ 2.8.(1주)",
                key="plan_period_input"
            )
        
        monthly_info_input = st.text_input(
            "월간보고 안내",
            value=saved_monthly_info if saved_monthly_info else "1월 업무 보고",
            key="monthly_info_input"
        )
        
        if st.button("💾 보고 기간 저장", type="primary"):
            if save_period_settings(spreadsheet, result_period_input, plan_period_input, monthly_info_input):
                st.success("✅ 보고 기간이 저장되었습니다!")
                st.balloons()
                st.cache_resource.clear()
                st.rerun()
            else:
                st.error("❌ 저장 실패! '설정' 시트가 있는지 확인하세요.")
    
    # 저장된 보고 기간 표시
    display_result = saved_result_period if saved_result_period else "설정되지 않음"
    display_plan = saved_plan_period if saved_plan_period else "설정되지 않음"
    display_monthly = saved_monthly_info if saved_monthly_info else "설정되지 않음"
    
    st.markdown(f"""
**보고기간**  
ㅇ 실        적: {display_result}  
ㅇ 계        획: {display_plan}  
ㅇ 월간보고: {display_monthly}
    """)
    
    st.markdown("---")
    
    # 탭 구성
    tab1, tab2 = st.tabs(["📄 보고서 작성", "📊 전체 보기"])
    
    # 탭 1: 보고서 작성
    with tab1:
        st.subheader("업무 보고 작성")
        
        col1, col2 = st.columns(2)
        
        with col1:
            writer = st.selectbox(
                "작성자 이름",
                ["연효흠", "여준섭", "류용재", "한은정", "박정현", "반경돈", "박희우"]
            )
        
        with col2:
            team = st.selectbox(
                "팀 선택",
                ["다각화사업팀", "개발팀", "기획팀", "디자인팀", "마케팅팀"]
            )
        
        st.markdown("---")
        
        # 사업개발 - 실적
        st.markdown("### 📌 [사업개발 - 실적]")
        st.caption("3자검사를 제외한 실적 내용 작성")
        
        business_result_template = """ㅇ 주제~작성
  - (일시 및 장소) 
  - (참석자) 
  - (내용) 
"""
        
        business_result_content = st.text_area(
            "사업개발 실적 내용 입력",
            value=business_result_template,
            height=200,
            key="business_result_input"
        )
        
        col1, col2, col3 = st.columns([2, 1, 2])
        with col2:
            if st.button("✅ 사업개발 실적 제출", type="primary", use_container_width=True, key="submit_business_result"):
                if not business_result_content or business_result_content.strip() == business_result_template.strip():
                    st.error("❌ 사업개발 실적 내용을 입력해주세요!")
                else:
                    try:
                        df = load_data(sheet)
                        new_id = len(df) + 1
                        
                        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        report = {
                            'id': new_id,
                            '작성시간': timestamp,
                            '팀': team,
                            '작성자': writer,
                            '구분': '사업개발-실적',
                            '내용': business_result_content
                        }
                        save_data(sheet, report)
                        st.success("✅ 사업개발 실적 제출 완료!")
                        st.balloons()
                    except Exception as e:
                        st.error(f"제출 실패: {e}")
        
        st.markdown("---")
        
        # 사업개발 - 계획
        st.markdown("### 📌 [사업개발 - 계획]")
        st.caption("3자검사를 제외한 계획 내용 작성")
        
        business_plan_template = """ㅇ 주제~작성
  - (일시 및 장소) 
  - (참석자) 
  - (내용) 
"""
        
        business_plan_content = st.text_area(
            "사업개발 계획 내용 입력",
            value=business_plan_template,
            height=200,
            key="business_plan_input"
        )
        
        col1, col2, col3 = st.columns([2, 1, 2])
        with col2:
            if st.button("✅ 사업개발 계획 제출", type="primary", use_container_width=True, key="submit_business_plan"):
                if not business_plan_content or business_plan_content.strip() == business_plan_template.strip():
                    st.error("❌ 사업개발 계획 내용을 입력해주세요!")
                else:
                    try:
                        df = load_data(sheet)
                        new_id = len(df) + 1
                        
                        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        report = {
                            'id': new_id,
                            '작성시간': timestamp,
                            '팀': team,
                            '작성자': writer,
                            '구분': '사업개발-계획',
                            '내용': business_plan_content
                        }
                        save_data(sheet, report)
                        st.success("✅ 사업개발 계획 제출 완료!")
                        st.balloons()
                    except Exception as e:
                        st.error(f"제출 실패: {e}")
        
        st.markdown("---")
        
        # 3자검사 - 실적
        st.markdown("### 🔍 [3자검사 - 실적]")
        st.caption("3자검사, 용접, RISE 사업 관련 실적 내용 작성")
        
        inspection_result_template = """ㅇ 주제~ 작성
  - (일시 및 장소) 
  - (참석자) 
  - (내용) 
"""
        
        inspection_result_content = st.text_area(
            "3자검사 실적 내용 입력",
            value=inspection_result_template,
            height=200,
            key="inspection_result_input"
        )
        
        col1, col2, col3 = st.columns([2, 1, 2])
        with col2:
            if st.button("✅ 3자검사 실적 제출", type="primary", use_container_width=True, key="submit_inspection_result"):
                if not inspection_result_content or inspection_result_content.strip() == inspection_result_template.strip():
                    st.error("❌ 3자검사 실적 내용을 입력해주세요!")
                else:
                    try:
                        df = load_data(sheet)
                        new_id = len(df) + 1
                        
                        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        report = {
                            'id': new_id,
                            '작성시간': timestamp,
                            '팀': team,
                            '작성자': writer,
                            '구분': '3자검사-실적',
                            '내용': inspection_result_content
                        }
                        save_data(sheet, report)
                        st.success("✅ 3자검사 실적 제출 완료!")
                        st.balloons()
                    except Exception as e:
                        st.error(f"제출 실패: {e}")
        
        st.markdown("---")
        
        # 3자검사 - 계획
        st.markdown("### 🔍 [3자검사 - 계획]")
        st.caption("3자검사, 용접, RISE 사업 관련 계획 내용 작성")
        
        inspection_plan_template = """ㅇ 주제~ 작성
  - (일시 및 장소) 
  - (참석자) 
  - (내용) 
"""
        
        inspection_plan_content = st.text_area(
            "3자검사 계획 내용 입력",
            value=inspection_plan_template,
            height=200,
            key="inspection_plan_input"
        )
        
        col1, col2, col3 = st.columns([2, 1, 2])
        with col2:
            if st.button("✅ 3자검사 계획 제출", type="primary", use_container_width=True, key="submit_inspection_plan"):
                if not inspection_plan_content or inspection_plan_content.strip() == inspection_plan_template.strip():
                    st.error("❌ 3자검사 계획 내용을 입력해주세요!")
                else:
                    try:
                        df = load_data(sheet)
                        new_id = len(df) + 1
                        
                        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        report = {
                            'id': new_id,
                            '작성시간': timestamp,
                            '팀': team,
                            '작성자': writer,
                            '구분': '3자검사-계획',
                            '내용': inspection_plan_content
                        }
                        save_data(sheet, report)
                        st.success("✅ 3자검사 계획 제출 완료!")
                        st.balloons()
                    except Exception as e:
                        st.error(f"제출 실패: {e}")
        
        st.markdown("---")
        
        # 월간보고 섹션
        st.markdown("### 📅 [월간보고]")
        st.caption("월간 업무 보고 내용 작성")
        
        monthly_template = """ㅇ 주제~ 작성
  - (일시 및 장소) 
  - (참석자) 
  - (내용) 
"""
        
        monthly_content = st.text_area(
            "월간보고 내용 입력",
            value=monthly_template,
            height=200,
            key="monthly_input"
        )
        
        col1, col2, col3 = st.columns([2, 1, 2])
        with col2:
            if st.button("✅ 월간보고 제출", type="primary", use_container_width=True, key="submit_monthly"):
                if not monthly_content or monthly_content.strip() == monthly_template.strip():
                    st.error("❌ 월간보고 내용을 입력해주세요!")
                else:
                    try:
                        df = load_data(sheet)
                        new_id = len(df) + 1
                        
                        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        report = {
                            'id': new_id,
                            '작성시간': timestamp,
                            '팀': team,
                            '작성자': writer,
                            '구분': '월간보고',
                            '내용': monthly_content
                        }
                        save_data(sheet, report)
                        st.success("✅ 월간보고 제출 완료!")
                        st.balloons()
                    except Exception as e:
                        st.error(f"제출 실패: {e}")
    
    # 탭 2: 전체 보기
    with tab2:
        st.subheader("📊 전체 보고서 현황")
        
        try:
            df = load_data(sheet)
            
            if df.empty:
                st.info("아직 제출된 보고서가 없습니다.")
            else:
                # 필터
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    selected_team = st.selectbox(
                        "팀 필터",
                        ["전체"] + list(df['팀'].unique()),
                        key="team_filter"
                    )
                
                with col2:
                    selected_type = st.selectbox(
                        "구분 필터",
                        ["전체"] + list(df['구분'].unique()),
                        key="type_filter"
                    )
                
                with col3:
                    st.metric("총 보고 건수", len(df))
                
                with col4:
                    if st.button("🔄 새로고침"):
                        st.rerun()
                
                # 필터링
                filtered_df = df.copy()
                if selected_team != "전체":
                    filtered_df = filtered_df[filtered_df['팀'] == selected_team]
                if selected_type != "전체":
                    filtered_df = filtered_df[filtered_df['구분'] == selected_type]
                
                st.markdown("---")
                
                # 구분별로 그룹화하여 표시
                for report_type in filtered_df['구분'].unique():
                    type_df = filtered_df[filtered_df['구분'] == report_type]
                    
                    st.markdown(f"## {report_type}")
                    
                    for team in type_df['팀'].unique():
                        team_df = type_df[type_df['팀'] == team]
                        
                        with st.expander(f"**{team}** ({len(team_df)}건)", expanded=True):
                            for idx, row in team_df.iterrows():
                                col1, col2 = st.columns([6, 1])
                                
                                with col1:
                                    st.markdown(f"{row['내용']}")
                                
                                with col2:
                                    if st.button("🗑️", key=f"delete_{idx}_{row['id']}", type="secondary", help="삭제"):
                                        try:
                                            delete_data(sheet, row['id'])
                                            st.success("삭제되었습니다!")
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"삭제 실패: {e}")
                                
                                st.markdown("---")
                
                # CSV 다운로드
                st.markdown("### 📥 데이터 다운로드")
                csv = df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📄 CSV 다운로드",
                    data=csv,
                    file_name=f"주간보고_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
                
        except Exception as e:
            st.error(f"데이터 로드 실패: {e}")

if __name__ == "__main__":
    main()
