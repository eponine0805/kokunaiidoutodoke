import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os

# --- データ処理のコアとなる関数 ---
def create_travel_form_df(template_path, data):
    """テンプレートCSVを読み込み、ユーザー入力データでDataFrameを更新する関数"""
    try:
        # ファイルの先頭36行をスキップしてデータ部分のみ読み込む
        df = pd.read_csv(template_path, header=None, skiprows=36)
    except FileNotFoundError:
        st.error(f"エラー: テンプレートファイル '{template_path}' が見つかりません。")
        # サーバー上のファイル一覧を表示して、ユーザーが確認できるようにする
        st.warning("サーバー上に存在するファイル:")
        # 現在のディレクトリのファイル/フォルダをリストする
        # (Streamlit Cloudでは権限により動作が異なる場合があります)
        try:
            files_in_directory = os.listdir('.')
            for f in files_in_directory:
                st.code(f)
        except Exception as list_e:
            st.error(f"ディレクトリのリスト取得中にエラー発生: {list_e}")
        return None
    except Exception as e:
        st.error(f"ファイルの読み込み中に予期せぬエラーが発生しました: {e}")
        return None

    # ヘッダー部分(先頭36行)を別途読み込む
    header_df = pd.read_csv(template_path, header=None, nrows=36)

    # --- ヘッダー部分にデータを書き込む ---
    header_df.iat[6, 4] = data["selected_title"]
    header_df.iat[8, 13] = datetime.now().strftime('%Y-%m-%d')
    header_df.iat[11, 13] = data["applicant_name"]
    header_df.iat[24, 2] = data["trip_purpose"]
    header_df.iat[25, 2] = data["main_destination"]
    header_df.iat[26, 2] = data["start_date_trip"].strftime('%Y-%m-%d')
    header_df.iat[26, 4] = data["end_date_trip"].strftime('%Y-%m-%d')
    header_df.iat[32, 7] = data["emergency_contact"]


    # --- 新しいスケジュールを作成 ---
    schedule_data = []
    # DataFrameの列数に合わせて空のリストを作成
    num_columns = header_df.shape[1]
    for item in data["schedule"]:
        # 基本的なデータをリストに追加
        row_data = [
            item["date"].strftime('%Y-%m-%d'), 
            item["dep_county"], 
            item["dep_town"],
            item["arr_county"], 
            item["arr_town"], 
            item["destination_detail"],
            item["transport"], 
            "", 
            item["dep_time"].strftime('%H:%M'), # 秒は不要な場合が多い
            item["arr_time"].strftime('%H:%M'), # 秒は不要な場合が多い
            item["hotel_name_tel"], 
            "",
            item["hotel_map_link"]
        ]
        # 残りの列を空文字列で埋める
        row = row_data + [""] * (num_columns - len(row_data))
        schedule_data.append(row)

    new_schedule_df = pd.DataFrame(schedule_data)
    
    # --- ヘッダー、新しいスケジュール、元のフッター部分を結合 ---
    # 元のデータから空のスケジュール行とヘッダーを除いた部分をフッターとする
    footer_start_row = 6 # 元のファイルのデータ部分のスケジュールは6行あると仮定
    footer_df = df.iloc[footer_start_row:] 

    final_df = pd.concat([header_df, new_schedule_df, footer_df], ignore_index=True)
    return final_df


# --- Streamlit UI部分 ---
st.set_page_config(layout="wide")
st.title('国内移動届 自動作成ツール ✈️')

# --- セッションステートの初期化 ---
if 'schedule' not in st.session_state:
    st.session_state.schedule = []
if 'dep_county' not in st.session_state: st.session_state.dep_county = "Muranga"
if 'dep_town' not in st.session_state: st.session_state.dep_town = "Gatanga"
if 'arr_county' not in st.session_state: st.session_state.arr_county = "Kiambu"
if 'arr_town' not in st.session_state: st.session_state.arr_town = "Thika"
# rerunしても入力値が保持されるようにキーを分ける
if 'dep_county_input' not in st.session_state: st.session_state.dep_county_input = st.session_state.dep_county
if 'dep_town_input' not in st.session_state: st.session_state.dep_town_input = st.session_state.dep_town
if 'arr_county_input' not in st.session_state: st.session_state.arr_county_input = st.session_state.arr_county
if 'arr_town_input' not in st.session_state: st.session_state.arr_town_input = st.session_state.arr_town


# --- UIの定義 ---
st.header("行程の出発地・到着地")
st.write("↓ デフォルトの出発地・到着地を入力してください。🔁ボタンで入れ替えも可能です。")

col_dep, col_swap, col_arr = st.columns([5, 1, 5])
with col_dep:
    dep_county = st.text_input("出発カウンティ", value=st.session_state.dep_county, key="dep_county_default")
    dep_town = st.text_input("出発タウン", value=st.session_state.dep_town, key="dep_town_default")
with col_swap:
    st.write("") 
    st.write("") 
    if st.button("🔁"):
        st.session_state.dep_county, st.session_state.arr_county = st.session_state.arr_county, st.session_state.dep_county
        st.session_state.dep_town, st.session_state.arr_town = st.session_state.arr_town, st.session_state.dep_town
        st.rerun()
with col_arr:
    arr_county = st.text_input("到着カウンティ", value=st.session_state.arr_county, key="arr_county_default")
    arr_town = st.text_input("到着タウン", value=st.session_state.arr_town, key="arr_town_default")

# ユーザーが入力したデフォルト値を更新
st.session_state.dep_county = dep_county
st.session_state.dep_town = dep_town
st.session_state.arr_county = arr_county
st.session_state.arr_town = arr_town


with st.form("travel_form", clear_on_submit=True):
    st.header("基本情報")
    title_options = ['Application for Official Trip', 'Order of Official Trip', 'Application for Private Trip']
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

    st.header("6. Schedule for All - 行程詳細")
    st.write("↓ 行程を入力し、「＋ 行程を追加」ボタンで行程リストに追加してください。")
    
    # 行程入力部分
    date = st.date_input("日付")
    
    # 出発地と到着地の入力フィールドをスケジュールフォーム内に移動
    dep_county_input = st.text_input("出発カウンティ", value=st.session_state.dep_county)
    dep_town_input = st.text_input("出発タウン", value=st.session_state.dep_town)
    arr_county_input = st.text_input("到着カウンティ", value=st.session_state.arr_county)
    arr_town_input = st.text_input("到着タウン", value=st.session_state.arr_town)
    
    destination_detail = st.text_input("目的地/詳細", "Matatu Stage")
    transport = st.text_input("移動手段", "Taxi, Matatu")
    time_cols = st.columns(2)
    dep_time = time_cols[0].time_input("出発時間")
    arr_time = time_cols[1].time_input("到着時間")
    hotel_name_tel = st.text_input("ホテル名と電話番号", "The Kiama River hotel +254725200665")
    hotel_map_link = st.text_input("ホテルのGoogle Mapリンク (任意)")

    add_clicked = st.form_submit_button("＋ 行程を追加")
    submitted = st.form_submit_button("✅ 移動届を生成する")

if add_clicked:
    st.session_state.schedule.append({
        "date": date, 
        "dep_county": dep_county_input, 
        "dep_town": dep_town_input,
        "arr_county": arr_county_input, 
        "arr_town": arr_town_input,
        "destination_detail": destination_detail, 
        "transport": transport,
        "dep_time": dep_time, 
        "arr_time": arr_time,
        "hotel_name_tel": hotel_name_tel, 
        "hotel_map_link": hotel_map_link
    })
    # フォームがクリアされてしまうので、デフォルト値をセッションステートに保持
    st.session_state.dep_county = arr_county_input
    st.session_state.dep_town = arr_town_input
    st.rerun()

if st.session_state.schedule:
    st.header("追加された行程リスト")
    # 不要な列を削除し、見やすいように整形
    df_display = pd.DataFrame(st.session_state.schedule)
    display_cols = ["date", "dep_county", "dep_town", "arr_county", "arr_town", "transport", "dep_time", "arr_time", "hotel_name_tel"]
    st.dataframe(df_display[display_cols].astype(str), use_container_width=True)

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
        # ★★★【修正点】アップロードされた正しいファイル名を参照 ★★★
        template_file = '国内移動届.xlsx - 申請様式（New）.csv'
        final_df = create_travel_form_df(template_file, user_data)
        
        if final_df is not None:
            output = io.StringIO()
            final_df.to_csv(output, index=False, header=False) # encodingはStringIOでは不要
            csv_data = output.getvalue().encode('utf-8-sig') # ダウンロード時にBOM付きUTF-8にエンコード
            
            st.balloons()
            st.success("移動届が正常に生成されました！下のボタンからダウンロードしてください。")
            st.download_button(
                label="📥 完成したCSVファイルをダウンロード",
                data=csv_data,
                file_name=f"国内移動届_{datetime.now().strftime('%Y%m%d')}.csv",
                mime='text/csv',
            )
            # 生成後にリストをクリア
            st.session_state.schedule = []