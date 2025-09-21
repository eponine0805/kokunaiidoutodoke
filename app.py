import streamlit as st
import pandas as pd
from datetime import datetime
import io

# --- データ処理のコアとなる関数 ---
def create_travel_form_df(template_path, data):
    """テンプレートCSVを読み込み、ユーザー入力データでDataFrameを更新する関数"""
    try:
        # テンプレートファイルを読み込む（ヘッダーなし）
        df = pd.read_csv(template_path, header=None)
    except FileNotFoundError:
        st.error(f"エラー: テンプレートファイル '{template_path}' が見つかりません。")
        st.info("app.pyと同じフォルダに '国内移動届.xlsx - 申請様式（New）.csv' があるか確認してください。")
        return None

    # --- 指定された項目を自動入力 ---
    # .iat[行インデックス, 列インデックス] を使って特定セルを更新します

    # Date (申請日) - セル(I, 9)付近
    df.iat[8, 13] = datetime.now().strftime('%Y-%m-%d')

    # 1. Purpose of Official Trip (目的) - セル(C, 25)付近
    df.iat[24, 2] = data["trip_purpose"]

    # 2. Destination (目的地) - セル(C, 26)付近
    df.iat[25, 2] = data["main_destination"]

    # 3. Period (期間) - セル(C, 27)と(E, 27)付近
    df.iat[26, 2] = data["start_date_trip"].strftime('%Y-%m-%d')
    df.iat[26, 4] = data["end_date_trip"].strftime('%Y-%m-%d')

    # 氏名や緊急連絡先なども更新
    df.iat[11, 13] = data["applicant_name"]
    df.iat[32, 7] = data["emergency_contact"]


    # 6. Schedule for All (スケジュール)
    schedule_start_row = 37  # スケジュールが始まる行のインデックス
    num_template_schedule_rows = 6 # テンプレートに存在するスケジュール行数

    schedule_data = []
    for item in data["schedule"]:
        row = [
            item["date"].strftime('%Y-%m-%d'),
            item["dep_county"],
            item["dep_town"],
            item["arr_county"],
            item["arr_town"],
            item["destination_detail"],
            item["transport"],
            "", # Flight No.
            item["dep_time"].strftime('%H:%M:%S'),
            item["arr_time"].strftime('%H:%M:%S'),
            item["hotel_name_tel"],
            "", # Spacer
            item["hotel_map_link"],
            "", "", "", "" # 残りの空セル
        ]
        schedule_data.append(row)

    # 新しいスケジュールをDataFrameに変換
    new_schedule_df = pd.DataFrame(schedule_data)
    new_schedule_df.columns = df.columns[:new_schedule_df.shape[1]]

    # 元のDataFrameを上部、スケジュール、下部に分割して結合
    df_top = df.iloc[:schedule_start_row]
    df_bottom = df.iloc[schedule_start_row + num_template_schedule_rows:]
    final_df = pd.concat([df_top, new_schedule_df, df_bottom], ignore_index=True)

    return final_df

# --- Streamlit UI部分 (Webページの表示) ---
st.set_page_config(layout="wide")
st.title('国内移動届 自動作成ツール ✈️')

# セッション状態でスケジュールリストを初期化
if 'schedule' not in st.session_state:
    st.session_state.schedule = []

# --- 入力フォーム ---
with st.form("travel_form"):
    st.header("必要事項を入力してください")

    # --- 基本情報セクション ---
    st.subheader("1. 目的、2. 目的地、3. 期間")
    trip_purpose = st.text_input("1. Purpose of Official Trip (目的)", "Field Trip")
    main_destination = st.text_input("2. Destination (目的地)", "Muranga county Gatanga")

    col1, col2 = st.columns(2)
    with col1:
        start_date_trip = st.date_input("3. Period (期間) - 開始日")
    with col2:
        end_date_trip = st.date_input("3. Period (期間) - 終了日")

    st.subheader("その他")
    applicant_name = st.text_input("申請者氏名", "Seiichiro Harauma")
    emergency_contact = st.text_input("緊急連絡先の電話番号", "254704387792")

    # --- スケジュールセクション ---
    st.header("6. Schedule for All")
    st.write("↓ 下のフォームで行程を1つずつ追加してください。")

    schedule_cols = st.columns((2, 2, 2, 2, 2, 3))
    date = schedule_cols[0].date_input("日付")
    dep_county = schedule_cols[1].text_input("出発カウンティ", "Nairobi")
    dep_town = schedule_cols[2].text_input("出発タウン", "CBD")
    arr_county = schedule_cols[3].text_input("到着カウンティ", "Kiambu")
    arr_town = schedule_cols[4].text_input("到着タウン", "Thika")
    destination_detail = schedule_cols[5].text_input("目的地/詳細", "Matatu Stage")

    time_cols = st.columns((2, 1, 1, 3, 3))
    transport = time_cols[0].text_input("移動手段", "Taxi, Matatu")
    dep_time = time_cols[1].time_input("出発時間")
    arr_time = time_cols[2].time_input("到着時間")
    hotel_name_tel = time_cols[3].text_input("ホテル名と電話番号", "The Kiama River hotel +254725200665")
    hotel_map_link = time_cols[4].text_input("ホテルのGoogle Mapリンク (任意)")

    # 「行程を追加」ボタン
    if st.form_submit_button("＋ 行程を追加"):
        st.session_state.schedule.append({
            "date": date, "dep_county": dep_county, "dep_town": dep_town,
            "arr_county": arr_county, "arr_town": arr_town, "destination_detail": destination_detail,
            "transport": transport, "dep_time": dep_time, "arr_time": arr_time,
            "hotel_name_tel": hotel_name_tel, "hotel_map_link": hotel_map_link
        })
        st.success("行程を追加しました！リストはフォームの下に表示されます。")

    # 「生成」ボタン
    submitted = st.form_submit_button("✅ 移動届を生成する")


# --- 追加されたスケジュールの表示 ---
if st.session_state.schedule:
    st.header("追加された行程リスト")
    schedule_df_display = pd.DataFrame(st.session_state.schedule)
    st.dataframe(schedule_df_display.astype(str), use_container_width=True)


# --- 生成ボタンが押された後の処理 ---
if submitted:
    if not st.session_state.schedule:
        st.warning("スケジュールに少なくとも1つの行程を追加してください。")
    else:
        user_data = {
            "applicant_name": applicant_name, "trip_purpose": trip_purpose,
            "main_destination": main_destination, "start_date_trip": start_date_trip,
            "end_date_trip": end_date_trip, "emergency_contact": emergency_contact,
            "schedule": st.session_state.schedule
        }

        # テンプレートファイルの名前を定義
        template_file = '国内移動届.xlsx - 申請様式（New）.csv'
        final_df = create_travel_form_df(template_file, user_data)

        if final_df is not None:
            # CSVをメモリ上で作成
            output = io.StringIO()
            final_df.to_csv(output, index=False, header=False, encoding='utf-8-sig')
            csv_data = output.getvalue()

            st.balloons()
            st.success("移動届が正常に生成されました！下のボタンからダウンロードしてください。")

            # ダウンロードボタン
            st.download_button(
                label="📥 完成したCSVファイルをダウンロード",
                data=csv_data,
                file_name=f"国内移動届_{datetime.now().strftime('%Y%m%d')}.csv",
                mime='text/csv',
            )