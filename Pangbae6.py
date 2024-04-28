import streamlit as st
import time
import pandas as pd
import openpyxl
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import plotly.express as px
import io  # 메모리 내 파일 작업을 위해 필요한 모듈 임포트

# Set Streamlit page to wide mode
st.set_page_config(layout="wide")

#사이드바에 이미지와 기타 요소를 추가합니다.
st.sidebar.markdown("""
            
    <style>
        .sidebar .sidebar-content {
            background-color: White;   
        }
        img.sidebar-image {
            display: block;
            margin: auto;  /* 상하 좌우 자동 마진을 적용하여 이미지를 가운데 정렬 */
            width: 200px;
        }
    </style>
    """, unsafe_allow_html=True)

st.sidebar.image("C:\\Elmes\\imgs\\elmes.png", width=200, use_column_width=False, output_format='PNG')

link = 'http://elmes.co.kr'
text = 'ELMES Korea Corp. 방문하기'
st.sidebar.markdown(
    f"<a href='{link}' target='_blank'>{text}</a>", unsafe_allow_html=True
)

# 메인 페이지에 헤더를 추가합니다.
with st.container():
    st.markdown("""
        <style>
            /* 상단 바 스타일 설정 */
            .top-bar {
                font-size: 30px;
                font-weight: normal;
                background-color: white; /* 배경색 설정 */
                padding: 10px;
                text-align: center;
                border: 2px solid grey; /* 테두리 설정 */
                box-shadow: 3px 3px 3px LightGrey; /* 그림자 효과 추가 */
                margin: 20px 20px; /* 위아래 여백 설정 */
                width: 650px; /* 너비 설정 */
                margin-left: 180px; /* 왼쪽 여백 설정 */
                margin-right: auto; /* 오른쪽 여백 자동 */
                margin-bottom: 50px; /* 아래쪽 여백 설정 */
            }
        </style>
        <div class='top-bar'>
            방배 6구역 지하수위 통합 관리 시스템
        </div>
        """, unsafe_allow_html=True)

# Function to create an Excel file in memory and provide a download link.
def generate_excel(df, start_date):
    columns_to_exclude = ['water_03', 'volt_01', 'volt_02', 'volt_03', 'volt_04', 'volt_05', 'volt_06', 'volt_07', 'volt_08', 'volt_09', 'volt_10', 'volt_11', 'volt_12', 'volt_13', 'volt_14', 'volt_15', 'volt_16', 'DateTime']
    df = df[(df['DateTime'] >= start_date) & (df['DateTime'] <= datetime.today())]
    df_filtered = df.drop(columns=columns_to_exclude, errors='ignore')
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_filtered.to_excel(writer, index=False)
    output.seek(0)
    return output

# 파일 저장 함수
def save_to_excel(df, start_date, filename):
    try:
        output = generate_excel(df, start_date)
        with open(filename, 'wb') as f:
            f.write(output.getvalue())
        st.sidebar.success(f"Saved to:\n{filename}")
    except Exception as e:
        st.sidebar.error(f"Failed to save the file: {str(e)}")

# # 오늘 날짜로 초기 파일명 생성
default_filename = f"c:\\Elmes\\downloads\\Pangbae_water_{datetime.now().strftime('%Y%m%d')}.xlsx"
value = default_filename

# Function to convert date and time strings into datetime objects
def convert_datetime(date_str, time_str):
    try:
        if pd.isna(date_str) or pd.isna(time_str) or not date_str.strip() or not time_str.strip():
            return pd.NaT
        date_object = pd.to_datetime(date_str, format='%Y. %m. %d', errors='raise')
        time_object = pd.to_datetime(time_str, format='%H:%M:%S', errors='raise').time()
        return pd.to_datetime(f"{date_object.date()} {time_object}")
    except ValueError as e:
        st.error(f"ValueError converting datetime: {e} with date: '{date_str}' and time: '{time_str}'")
        return pd.NaT
    except Exception as e:
        st.error(f"Other error converting datetime: {e}")
        return pd.NaT

# Authenticate with Google Sheets
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(r'c:\Elmes\keys\pangbae-water-8f0f0bec873b.json', scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_key('19FU_a8Rgy2iIkdqFW8ut4V0FsuFWKq6SsJ7UfrCkc18')
sheet = spreadsheet.get_worksheet(0)

# Define a function to fetch and prepare data
def load_data():
    data = sheet.get_all_records()
    df = pd.DataFrame(data)
    df = df.dropna(subset=['Date', 'Time'])
    df['DateTime'] = df.apply(lambda row: convert_datetime(row['Date'], row['Time']), axis=1)
    return df.dropna(subset=['DateTime']).sort_values(by='DateTime')

# Ensure df is loaded into session state
if 'df' not in st.session_state:
    st.session_state.df = load_data()

df = st.session_state.df

# Initialize page in session state if it's not already set
if 'page' not in st.session_state:
    st.session_state.page = 0

def change_page(delta):
    st.session_state.page += delta

# Sidebar configuration
st.sidebar.header("기간 선택:")
today = datetime.today()
date_options = {f" {i} 일간 테이터": today - timedelta(days=i) for i in [7, 30, 180, 365]}
selected_option = st.sidebar.selectbox('    7,30,180,365 일', options=list(date_options.keys()), index=0)
start_date = date_options[selected_option]

filtered_df = df[(df['DateTime'] >= start_date) & (df['DateTime'] <= today)].copy()
filtered_df['TimeOnly'] = filtered_df['DateTime'].dt.strftime('%H:%M:%S')

#st.sidebar.header("Data Settings")
if st.sidebar.button('업데이트', key='update_data_button'):
    st.session_state.df = load_data()
    st.rerun()

# No need to check if a button was pressed
st.sidebar.image(r'C:\Elmes\imgs\exc.png', width=40)  # 이미지 경로와 너비 조절
data_to_download = generate_excel(df, start_date)
st.sidebar.download_button(
    label="엑셀 다운로드",
    data=data_to_download,
    file_name=f"Pangbae_water_{datetime.now().strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="엑셀 파일이 '다운로드' 폴더에 저장됩니다"
)

# if st.sidebar.button('엑셀 다운로드'):
#     data_to_download = generate_excel(df, start_date)
#     output_bytes = data_to_download.getvalue()  # Retrieve bytes buffer from io.BytesIO
#     btn = st.sidebar.download_button(
#         label="Download Excel",
#         data=output_bytes,
#         file_name=f"Pangbae_water_{datetime.now().strftime('%Y%m%d')}.xlsx",
#         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         on_click=lambda: st.sidebar.success('File downloading...')  # Optional feedback message
#     )


# if st.sidebar.button('엑셀 저장'):
#     save_to_excel(df, start_date, value)

# Define total pages and manage pagination
total_pages = 4

plot_columns = {
    0: ['water_01', 'water_02', 'water_04', 'water_05'],
    1: ['water_06', 'water_07', 'water_08', 'water_09'],
    2: ['water_10', 'water_11', 'water_12', 'water_13'],
    3: ['water_14', 'water_15', 'water_16']
}[st.session_state.page]

for i, col in enumerate(plot_columns):
    column_series = pd.to_numeric(filtered_df[col], errors='coerce')
    min_val = column_series.min(skipna=True)
    max_val = column_series.max(skipna=True)
    if column_series.isnull().all():
        continue

    plot_data = filtered_df[column_series.notna()]
    fig = px.line(
        plot_data, x='DateTime', y=col,
        title=f"{col}:  GL (-)",
        labels={'DateTime': 'Date and Time', col: 'Measurement Level'},
        render_mode='webgl'
    )

    fig.update_traces(
        hovertemplate='<b>Date:</b> %{x|%Y-%m-%d}<br><b>Time:</b> %{customdata}<br><b>Level:</b> %{y:.2f}',
        customdata=plot_data['TimeOnly']
    )

    colors = ['blue', 'green', 'red', 'purple', 'orange', 'yellow', 'cyan', 'magenta', 'pink']
    fig.update_traces(mode='markers+lines', marker=dict(size=5), line=dict(color=colors[i % len(colors)], width=2))

    fig.update_layout(
        width=1100,
        height=300,
        yaxis_range=[min_val - 0.1, max_val + 0.1],
        xaxis=dict(
            title='Date and Time',
            title_font=dict(
                color='Black',
                size=18
            ),
            tickformat='%Y-%m-%d', 
            tickangle=35, 
            gridcolor='LightGrey',
            showline=True,
            linecolor='DarkGrey', 
            linewidth=2
        ),
        yaxis=dict(
            title='Level (Meter)',
            title_font=dict(
                color='Black',
                size=18
            ),
            gridcolor='LightGrey',
            showline=True, 
            linecolor='DarkGrey',
            linewidth=2, 
            tickformat='.2f'
        ),
        plot_bgcolor='white',
        paper_bgcolor='white',
        margin=dict(l=20, r=20, t=40, b=20),
        hovermode='x unified',
        showlegend=False
    )
    st.plotly_chart(fig)

# Pagination buttons
if st.session_state.page > 0:
    st.button('Previous', on_click=lambda: change_page(-1))

if st.session_state.page < total_pages - 1:
    st.button('Next', on_click=lambda: change_page(1))

# Display current page number
st.write(f"Current page: {st.session_state.page + 1}/{total_pages}")
