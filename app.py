import streamlit as st
import pandas as pd
from datetime import datetime
import io

# --- データ処理のコアとなる関数 ---
def create_travel_form_df(template_path, data):
    """テンプレートCSVを読み込み、ユーザー入力データでDataFrameを更新する関数"""
    try:
        df = pd.read_csv(template_path, header=None)
    except FileNotFoundError:
        st.error(f"エラー: テンプレートファイル '{template_path}' が見つかりません。")
        return None

    df.iat[6, 4] = data["selected_title"]
    df.iat[8, 13] = datetime.now().strftime('%Y-%m-%d')
    df.iat[24, 2] = data["trip_purpose"]
    df.iat[25, 2] = data["main_destination"]
    df.iat[26, 2] = data["start_date_trip"].strftime('%Y-%m-%d')
    df.iat[26, 4] = data["end_date_trip"].strftime('%Y-%m-%d')
    df.iat[11, 13] = data["applicant_name"]
    df.iat[32, 7] = data["emergency_contact"]

    schedule_start_row = 37
    num_template_schedule_rows = 6
    schedule_data = []
    for item in data["schedule"]:
        row = [
            item["date"].strftime('%Y-%m-%d'), item["dep_county"], item["dep_town"],
            item["arr_county"], item["arr_town"], item["destination_detail"],
            item["transport"], "", item["dep_time"].strftime('%H:%M:%S'),
            item["arr_time"].strftime('%H:%M:%S'), item["hotel_name_tel"], "",
            item["hotel_map_link"], "", "", "", ""
        ]
        schedule_data.append(row)
    new_schedule_df = pd.DataFrame(schedule_data)
    if not new_schedule_df.empty:
        new_schedule_df.columns = df.columns[:new_schedule_df.shape[1]]
    df_top = df.iloc[:schedule_start_row]
    df_bottom = df.iloc[schedule_start_row + num_template_schedule_rows:]
    final_df = pd.concat([df_top, new_schedule_df, df_bottom], ignore_index=True)
    return final_df

# --- Streamlit UI部分 ---
st.set_page_config(layout="wide")
st.title('国内移動届 自動作成ツール ✈️')

# --- セッションステートの初期化 ---
if 'schedule' not in st.session_state:
    st.session_state.schedule = []
# スケジュール入力欄の状態を管理
if 'dep_county' not in st.session_state: st.session_state.dep_county = "Nairobi"
if 'dep_town' not in st.session_state: st.session_state.dep_town = "CBD"
if 'arr_county' not in st.session_state: st.session_state.arr_county = "Kiambu"
if 'arr_town' not in st.session_state: st.session_state.arr_town = "Thika"

# --- メインフォーム ---
with st.form("travel_form"):
    st.header("必要事項を入力してください")
    
    # ★★★【変更点】様式の選択肢を3つに ★★★
    # 注：3つ目の選択肢の名称を書き換えてください
    title_options = ['Application for Official Trip', 'Order of Official Trip', '【３つ目の選択肢をここに入力】']
    selected_title = st.selectbox('様式の種類を選択してください', title_options)

    st.subheader("1. 目的、2. 目的地、3. 期間")
    purpose_options = ['Field Trip', 'shopping', 'transfer']
    trip_purpose = st.selectbox('1. Purpose of Official Trip (目的)', purpose_options)
    main_destination = st.text_input("2. Destination (目的地)", "Muranga county Gatanga")
    col1, col2 = st.columns(2)
    start_date_trip = col1.date_input("3. Period (期間) - 開始日")
    end_date_trip = col2.date_input("3. Period (期間) - 終了日")

    st.subheader("その他")
    applicant_name = st.text_input("申請者氏名", "Seiichiro Harauma")
    emergency_contact = st.text_input("緊急連絡先の電話番号", "254704387792")

    # --- スケジュール入力 ---
    st.header("6. Schedule for All")
    st.write("↓ 下のフォームで行程を1つずつ追加してください。")
    
    # --- ★★★【新機能】出発地と到着地の入力＆入れ替え機能 ---
    col_dep, col_swap, col_arr = st.columns([5, 1, 5])
    with col_dep:
        st.text_input("出発カウンティ", key="dep_county")
        st.text_input("出発タウン", key="dep_town")
    with col_swap:
        st.write("") # スペース調整
        st.write("") # スペース調整
        swap_button = st.button("🔁 入れ替え")
    with col_arr:
        st.text_input("到着カウンティ", key="arr_county")
        st.text_input("到着タウン", key="arr_town")
    
    if swap_button:
        # 値を入れ替える
        st.session_state.dep_county, st.session_state.arr_county = st.session_state.arr_county, st.session_state.dep_county
        st.session_state.dep_town, st.session_state.arr_town = st.session_state.arr_town, st.session_state.dep_town
        # 画面を再描画して、入力ボックスの表示を更新
        st.rerun()

    schedule_cols_1 = st.columns((2, 3))
    date = schedule_cols_1[0].date_input("日付")
    destination_detail = schedule_cols_1[1].text_input("目的地/詳細", "Matatu Stage")

    time_cols = st.columns((2, 1, 1, 3, 3))
    transport = time_cols[0].text_input("移動手段", "Taxi, Matatu")
    dep_time = time_cols[1].time_input("出発時間")
    arr_time = time_cols[2].time_input("到着時間")
    hotel_name_tel = time_cols[3].text_input("ホテル名と電話番号", "The Kiama River hotel +254725200665")
    hotel_map_link = time_cols[4].text_input("ホテルのGoogle Mapリンク (任意)")

    # --- フォームのボタン ---
    add_clicked = st.form_submit_button("＋ 行程を追加")
    submitted = st.form_submit_button("✅ 移動届を生成する")

# --- ボタンが押された後の処理 ---
if add_clicked:
    st.session_state.schedule.append({
        "date": date, "dep_county": st.session_state.dep_county, "dep_town": st.session_state.dep_town,
        "arr_county": st.session_state.arr_county, "arr_town": st.session_state.arr_town,
        "destination_detail": destination_detail, "transport": transport,
        "dep_time": dep_time, "arr_time": arr_time,
        "hotel_name_tel": hotel_name_tel, "hotel_map_link": hotel_map_link
    })
    st.success("行程を追加しました！")

if st.session_state.schedule:
    st.header("追加された行程リスト")
    st.dataframe(pd.DataFrame(st.session_state.schedule).astype(str), use_container_width=True)

if submitted:
    if not st.session_state.schedule:
        st.warning("スケジュールに少なくとも1つの行程を追加してください。")
    else:
        user_data = {
            "selected_title": selected_title, "applicant_name": applicant_name,
            "trip_purpose": trip_purpose, "main_destination": main_destination,
            "start_date_trip": start_date_trip, "end_date_trip": end_date_trip,
            "emergency_contact": emergency_contact, "schedule": st.session_state.schedule
        }
        template_file = '国内移動届.xlsx - 申請様式（New）.csv'
        final_df = create_travel_form_df(template_file, user_data)
        if final_df is not None:
            output = io.StringIO()
            final_df.to_csv(output, index=False, header=False, encoding='utf-8-sig')
            csv_data = output.getvalue()
            st.balloons()
            st.success("移動届が正常に生成されました！下のボタンからダウンロードしてください。")
            st.download_button(
                label="📥 完成したCSVファイルをダウンロード",
                data=csv_data,
                file_name=f"国内移動届_{datetime.now().strftime('%Y%m%d')}.csv",
                mime='text/csv',
            )