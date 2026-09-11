import streamlit as st

# 페이지 기본 설정
st.set_page_config(
    page_title="AXIS TECH - 시스템 관리",
    page_icon="⚙️",
    layout="wide"
)

# 헤더 영역
st.title("⚙️ AXIS TECH 웹 시스템")
st.caption("Streamlit Cloud 클라우드 호스팅 서비스")

st.divider()

# 메인 콘텐츠 영역
col1, col2 = st.columns([2, 1])

with col1:
    st.subheader("📌 시스템 상태")
    st.success("클라우드 서버가 정상적으로 작동 중입니다.")
    
    st.info("기존 데스크톱 UI(Tkinter)를 웹 기반 Streamlit 대시보드로 전환하는 중입니다.")

with col2:
    st.subheader("💡 빠른 메뉴")
    st.button("데이터 새로고침", use_container_width=True)
    st.button("로그 보기", use_container_width=True)

st.divider()
st.write("© AXIS TECH. All rights reserved.")

