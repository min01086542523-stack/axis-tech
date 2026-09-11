# 엑스테크 생산 MES

데스크톱(CustomTkinter)과 웹(Streamlit)을 함께 제공합니다.

## 웹으로 실행

```
pip install -r requirements.txt
streamlit run app.py
```

Streamlit Community Cloud: Main file path는 `app.py` 입니다.

클라우드에서는 App settings → Secrets에 `DATABASE_URL`을 넣으면 PC MES와 같은 Supabase 테이블을 사용합니다. 형식은 `.streamlit/secrets.toml.example` 을 참고하세요.

## 데스크톱 실행

```
python main.py
```
