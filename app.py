import streamlit as st
import pandas as pd

# 1. 페이지 설정 (스마트폰/PC 겸용)
st.set_page_config(
    page_title="AXIS TECH - 계정 및 거래처 관리",
    page_icon="💼",
    layout="wide"
)

# 2. 타이틀 영역
st.title("💼 AXIS TECH 관리 시스템")
st.caption("모바일 및 PC 겸용 통합 대시보드")

# 3. 사이드바 메뉴 (모바일에서는 좌측 상단 > 버튼으로 열림)
menu = st.sidebar.selectbox(
    "📌 메뉴 선택",
    ["계정/거래처 목록", "신규 데이터 등록", "시스템 정보"]
)

# --- [메뉴 1: 계정/거래처 목록] ---
if menu == "계정/거래처 목록":
    st.subheader("📋 계정 및 거래처 현황")
    
    # 검색 및 필터 (모바일 환경에 맞춘 배치)
    search_term = st.text_input("🔍 거래처명 또는 계정 검색", "")
    
    # 예시 데이터 (실제 데이터베이스 연동 영역)
    sample_data = [
        {"ID": 1, "구분": "매출처", "거래처명": "(주)에이시스", "대표자": "홍길동", "연락처": "010-1234-5678", "상태": "정상"},
        {"ID": 2, "구분": "매입처", "거래처명": "텍스엔지니어링", "대표자": "이몽룡", "연락처": "010-9876-5432", "상태": "정상"},
        {"ID": 3, "구분": "매출처", "거래처명": "한국스틸", "대표자": "성춘향", "연락처": "031-111-2222", "상태": "대기"},
    ]
    df = pd.DataFrame(sample_data)
    
    # 검색어 필터링
    if search_term:
        df = df[df["거래처명"].str.contains(search_term) | df["구분"].str.contains(search_term)]
    
    # 모바일/PC 반응형 표 출력
    st.dataframe(df, use_container_width=True, hide_index=True)

# --- [메뉴 2: 신규 데이터 등록] ---
elif menu == "신규 데이터 등록":
    st.subheader("📝 신규 계정/거래처 등록")
    
    with st.form("account_form"):
        col1, col2 = st.columns(2)
        with col1:
            category = st.selectbox("구분", ["매출처", "매입처", "기타"])
            company_name = st.text_input("거래처명 *")
            ceo_name = st.text_input("대표자명")
        with col2:
            phone = st.text_input("연락처")
            biz_num = st.text_input("사업자등록번호")
            address = st.text_input("주소")
            
        submitted = st.form_submit_button("💾 저장하기", use_container_width=True)
        if submitted:
            if company_name:
                st.success(f"'{company_name}' 등록 완료되었습니다.")
            else:
                st.warning("거래처명을 입력해 주세요.")

# --- [메뉴 3: 시스템 정보] ---
elif menu == "시스템 정보":
    st.subheader("⚙️ 시스템 상태")
    st.info("클라우드 서버 24시간 가동 중 (스마트폰 접속 지원)")
    st.write("© AXIS TECH. All rights reserved.")


