import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os

# --- データ処理のコアとなる関数 (Excel対応版) ---
def create_travel_form_df(template_path, data):
    """テンプレートExcelを読み込み、ユーザー入力データで更新する関数"""
    try:
        # openpyxl をエンジンとして指定してExcelファイルを読み込む
        # ヘッダーがないファイルとして読み込む
        df = pd.read_excel(template_path, header=None, engine='openpyxl', sheet_name=0)
    except FileNotFoundError:
        st.error(f"エラー: テンプレートファイル '{template_path}' が見つかりません。")
        st.warning("サーバー上に存在するファイル:")
        try:
            files_in_directory = os.listdir('.')
            for f in files_in_directory:
                st.code(f)
        except Exception as list_e:
            st.error(f"ディレクトリのリスト取得中にエラー発生: {list_e}")
        return None
    except Exception as e:
        st.error(f"Excelファイルの読み込み中に予期せぬエラーが発生しました: {e}")
        return None

    # 元のDataFrameをコピーして、書き込み用のDataFrameを作成
    write_df = df.copy()

    # --- ヘッダー部分にデータを書き込む (iatからilocに変更) ---
    # .iatは高速ですが、予期せぬ型変換を避けるためilocで安全に値を設定します
    write_df.iloc[6, 4] = data["selected_title"]
    write_df.iloc[8, 13] = datetime.now().strftime('%Y-%m-%d')
    write_df.iloc[11, 13] = data["applicant_name"]
    write_df.iloc[24, 2] = data["trip_purpose"]
    write_df.iloc[25, 2] = data["main_destination"]
    write_df.iloc[26, 2] = data["start_date_trip"].strftime('%Y-%m-%d')
    write_df.iloc[26, 4] = data["end_date_trip"].strftime('%Y-%m-%d')
    write_df.iloc[32, 7] = data["emergency_contact"]

    # --- 新しいスケジュールを作成 ---
    schedule_data = []
    num_columns = write_df.shape[1]
    for item in data["schedule"]:
        row_data = [
            item["date"].strftime('%Y-%m-%d'), 
            item["dep_county"], 
            item["dep_town"],
            item["arr_county"], 
            item["arr_town"], 
            item["destination_detail"],
            item["transport"], 
            "", 
            item["dep_time"].strftime('%H:%M'),
            item["arr_time"].strftime('%H:%M'),
            item["hotel_name_tel"], 
            "",
            item["hotel_map_link"]
        ]
        # 元のExcelの列数に合わせて空のデータを追加
        row = row_data + [""] * (num_columns - len(row_data))
        schedule_data.append(row)

    new_schedule_df = pd.DataFrame(schedule_data)

    # --- データフレームのパーツを結合 ---
    # 1. ヘッダー部分 (スケジュール開始行まで)
    header_part = write_df.iloc[:36]
    # 2. フッター部分 (元のスケジュールの終わりから)
    footer_part = write_df.iloc[36 + 6:] # 元のテンプレートのスケジュールが6行あると仮定

    # ヘッダー、新しいスケジュール、フッターを結合
    final_df = pd.concat([header_part, new_schedule_df, footer_part], ignore_index=True)

    return final_df


# --- Streamlit UI部分 (変更なし) ---
st.set_page_config(layout="wide")
st.title('国内移動届 自動作成ツール ✈️ (Excel対応版)')

if 'schedule' not in st.session_state:
    st.session_state.schedule = []
if 'dep_county' not in st.session_state: st.session_state.dep_county = "Muranga"
if 'dep_town' not in st.session_state: st.session_state.dep_town = "Gatanga"
if 'arr_county' not in st.session_state: st.session_state.arr_county = "Kiambu"
if 'arr_town' not in st.session_state: st.session_state.arr_town = "Thika"

st.header("行程の出発地・到着地")
col_dep, col_swap, col_arr = st.columns([5, 1, 5])
with col_dep:
    dep_county = st.text_input("出発カウンティ", st.session_state.dep_county)
    dep_town = st.text_input("出発タウン", st.session_state.dep_town)
with col_swap:
    st.write("") 
    st.write("") 
    if st.button("🔁 入れ替え"):
        dep_county, arr_county = arr_county, dep_county
        dep_town, arr_town = arr_town, dep_town
        st.session_state.dep_county = dep_county
        st.session_state.dep_town = dep_town
        st.session_state.arr_county = arr_county
        st.session_state.arr_town = arr_town
        st.rerun()
with col_arr:
    arr_county = st.text_input("到着カウンティ", st.session_state.arr_county)
    arr_town = st.text_input("到着タウン", st.session_state.arr_town)

# フォームの外でセッションステートを更新
st.session_state.dep_county = dep_county
st.session_state.dep_town = dep_town
st.session_state.arr_county = arr_county
st.session_state.arr_town = arr_town

with st.form("travel_form"):
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

    st.header("スケジュール詳細")
    date = st.date_input("日付")
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
        "date": date, "dep_county": st.session_state.dep_county, "dep_town": st.session_state.dep_town,
        "arr_county": st.session_state.arr_county, "arr_town": st.session_state.arr_town,
        "destination_detail": destination_detail, "transport": transport,
        "dep_time": dep_time, "arr_time": arr_time,
        "hotel_name_tel": hotel_name_tel, "hotel_map_link": hotel_map_link
    })
    # 次の入力のために、到着地を出発地として設定
    st.session_state.dep_county = st.session_state.arr_county
    st.session_state.dep_town = st.session_state.arr_town
    st.rerun()

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
        
        # ★★★【修正点】元のExcelファイル名を参照 ★★★
        template_file = '国内移動届.xlsx' 
        final_df = create_travel_form_df(template_file, user_data)
        
        if final_df is not None:
            # ★★★【修正点】Excelファイルとして出力する処理 ★★★
            output = io.BytesIO()
            # pandas DataFrame を Excel 形式で BytesIO オブジェクトに書き込む
            # indexとheaderは元のExcelに合わせて出力しない
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, header=False, sheet_name='申請様式（New）')
            
            excel_data = output.getvalue()
            
            st.balloons()
            st.success("移動届が正常に生成されました！下のボタンからダウンロードしてください。")
            st.download_button(
                label="📥 完成したExcelファイルをダウンロード",
                data=excel_data,
                file_name=f"国内移動届_{datetime.now().strftime('%Y%m%d')}.xlsx",
                # ExcelファイルのMIMEタイプを指定
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
            # 生成後にリストをクリア
            st.session_state.schedule = []