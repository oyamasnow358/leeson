import streamlit as st
import pandas as pd
import gspread # Google Sheets API連携用
from oauth2client.service_account import ServiceAccountCredentials # 認証情報
import json # JSONキーファイル読み込み用
from io import BytesIO # Excelアップロード/ダウンロード用
from openpyxl import load_workbook # Excel読み書き用

st.set_page_config(
    page_title="GoogleフォームからExcel授業カード作成",
    page_icon="📝",
    layout="centered"
)

st.title("📝 GoogleフォームからExcel授業カードを作成")
st.markdown("---")

# Google Sheets APIからデータを取得する関数
@st.cache_data(ttl=3600)
def load_data_from_google_sheet(spreadsheet_name, worksheet_name):
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        
        creds_json_string = st.secrets.get("GOOGLE_SHEETS_CREDENTIALS")
        
        if creds_json_string is None:
            st.error(
                "Streamlitのsecretsに 'GOOGLE_SHEETS_CREDENTIALS' キーが見つかりません。\n"
                "サービスアカウントのJSONキーを `secrets.toml` または Streamlit Cloudの設定で `GOOGLE_SHEETS_CREDENTIALS` "
                "という名前で設定してください。"
            )
            st.stop()
        
        creds_info = json.loads(creds_json_string)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
        client = gspread.authorize(creds)
        
        spreadsheet = client.open(spreadsheet_name)
        worksheet = spreadsheet.worksheet(worksheet_name)
        
        data = worksheet.get_all_values()
        
        # ヘッダー行を特定し、その後のデータを辞書のリストとして処理
        if not data:
            return []

        headers = data[0]
        records = data[1:]
        
        processed_records = []
        for row in records:
            if any(cell.strip() for cell in row): # 空行をスキップ
                row_dict = {}
                for i, header in enumerate(headers):
                    if i < len(row): # 列のインデックスが範囲内であることを確認
                        value = str(row[i]).strip()
                        # リストとして扱う項目 (セミコロン区切り)
                        if header in ['導入の流れ', '活動の流れ', '振り返りの流れ', '指導のポイント', '教材写真URL']:
                            row_dict[header] = [item.strip() for item in value.split(';') if item.strip()] if value else []
                        # ハッシュタグ (カンマ区切り)
                        elif header == 'ハッシュタグ':
                            row_dict[header] = [item.strip() for item in value.split(',') if item.strip()] if value else []
                        # 数値に変換する項目 (エラーハンドリング付き)
                        elif header == '単元内での並び順':
                            try:
                                row_dict[header] = int(value)
                            except (ValueError, TypeError):
                                row_dict[header] = 9999 # デフォルト値
                        # ICT活用有無 (ブール値または文字列)
                        elif header == 'ICT活用有無':
                            val_lower = value.lower()
                            if val_lower == 'true' or val_lower == 'はい':
                                row_dict[header] = 'あり'
                            elif val_lower == 'false' or val_lower == 'いいえ':
                                row_dict[header] = 'なし'
                            else:
                                row_dict[header] = value # その他の場合はそのまま
                        else:
                            row_dict[header] = value
                    else:
                        row_dict[header] = '' # データがない場合、空文字列を設定

                # ユニークなIDを付与 (Google Sheets由来であることを示す)
                # タイムスタンプをIDの一部として使うことで、よりユニーク性を高める
                timestamp = row_dict.get('タイムスタンプ', f"no_timestamp_{len(processed_records)}")
                row_dict['generated_id'] = f"gs_{timestamp}_{len(processed_records)}"
                processed_records.append(row_dict)
        
        return processed_records

    except KeyError as e:
        st.error(f"Google Sheets APIの認証情報が見つかりません。`secrets.toml`またはStreamlit Cloudの設定を確認してください: {e}")
        st.stop()
    except Exception as e:
        st.error(f"Googleスプレッドシートからのデータ読み込み中に予期せぬエラーが発生しました: {e}")
        st.exception(e)
        st.stop()
    return []

# Excelテンプレートにフォームデータを書き込む関数
def generate_excel_from_form_data(form_data):
    output_excel = BytesIO()
    try:
        # 既存のExcelテンプレートファイル（授業カード.xlsm）を読み込む
        # openpyxlはxlsmファイルを扱えますが、VBAマクロは保持するものの実行はできません。
        # VBAマクロを完全に機能させたい場合は、Pythonから外部ツールを呼び出すなど複雑な対応が必要です。
        # 今回はデータを書き込むことのみに焦点を当てます。
        with open("授業カード.xlsm", "rb") as f:
            workbook_data = BytesIO(f.read())
        
        wb = load_workbook(workbook_data, read_only=False, keep_vba=True)
            
# 変更前
# ws = wb.active # アクティブなシート

# 変更後
        ws = wb['授業カード'] # シート名「授業カード」を指定

  

        # Googleフォームの項目名（スプレッドシートのヘッダー名）とExcelのセル位置を対応させる
        # このマッピングは、あなたの「授業カード.xlsm」のレイアウトに合わせて調整してください。
        # 例: Googleフォームの項目名 -> Excelのセル
        cell_mappings = {

            '知的段階及び発達段階': 'B2',
            '単元名': 'B3',
            'キャッチコピー': 'B5',
            '授業のねらい': 'B9',
            '学部学年': 'A5',
            '障害種': 'A2',
            '授業時間': 'A4',
            '準備物': 'B15',
            '導入の内容': 'B11',
            '展開の内容': 'B12',
            'まとめの内容': 'B13',
            '授業のPOINT': 'B10',
            '検索ワード': 'B23',
            'ICT活用': 'B21',
            '教科': 'A3',
            '学習形態': 'B7',
            '授業タイトル': 'B4',
            '単元内で何回目の授業か': 'B8', # 新規追加、必要に応じて
        }

        for form_field, excel_cell in cell_mappings.items():
            value = form_data.get(form_field, '')
            if isinstance(value, list):
                # リストの場合はセミコロンまたはカンマで結合（用途による）
                if form_field == 'ハッシュタグ':
                    ws[excel_cell] = ','.join(value)
                else:
                    ws[excel_cell] = ';'.join(value)
            else:
                ws[excel_cell] = str(value)
        
        wb.save(output_excel)
        output_excel.seek(0)
        return output_excel.getvalue()

    except FileNotFoundError:
        st.error("⚠️ '授業カード.xlsm' テンプレートファイルが見つかりません。アプリケーションと同じ階層に配置してください。")
        return None
    except Exception as e:
        st.error(f"Excelファイルの生成中にエラーが発生しました: {e}")
        st.exception(e)
        return None

# --- Streamlit UI ---

st.info("""
このアプリでは、Googleフォームで入力された回答をもとに、個別の授業カードExcelファイルを生成・ダウンロードできます。
利用するためには、`secrets.toml` ファイルまたはStreamlit Cloudのシークレット設定が必要です。
""")

# Google Sheets API の設定を secrets から取得
GOOGLE_SHEET_SPREADSHEET_NAME = st.secrets.get("google_sheet_spreadsheet_name", "あなたのスプレッドシート名")
GOOGLE_SHEET_WORKSHEET_NAME = st.secrets.get("google_sheet_worksheet_name", "フォームの回答 1")

# GoogleフォームのURL (ユーザーに表示するためのダミーリンク)
# 実際に使用するGoogleフォームのURLに置き換えてください。
google_form_input_link = st.secrets.get("google_form_url", "https://forms.gle/YOUR_ACTUAL_GOOGLE_FORM_LINK")

st.markdown(
    f"""
    <p style="font-size:1.1em;">入力用のGoogleフォームはこちら: <a href="{google_form_input_link}" target="_blank">📝 Googleフォームを開く</a></p>
    """, unsafe_allow_html=True
)

if google_form_input_link == "https://forms.gle/YOUR_ACTUAL_GOOGLE_FORM_LINK":
    st.warning("⚠️ GoogleフォームのURLを、Streamlitのsecrets (`google_form_url` キー) またはコード内で実際のURLに更新してください。")

# 最新のフォーム回答を読み込むボタン
if st.button("🔄 最新のフォーム回答を読み込む"):
    load_data_from_google_sheet.clear() # キャッシュをクリアして再読み込みを強制
    sheet_lesson_data_records = load_data_from_google_sheet(
        spreadsheet_name=GOOGLE_SHEET_SPREADSHEET_NAME,
        worksheet_name=GOOGLE_SHEET_WORKSHEET_NAME
    )
    if sheet_lesson_data_records:
        st.session_state.google_form_records = sheet_lesson_data_records
        st.success(f"{len(sheet_lesson_data_records)}件のGoogleフォーム回答を読み込みました！")
    else:
        st.info("Googleフォームからの回答が見つかりませんでした。")

# Googleフォームから取得したデータがある場合、選択・ダウンロードUIを表示
if 'google_form_records' in st.session_state and st.session_state.google_form_records:
    
    st.markdown("---")
    st.subheader("⬇️ Excel授業カード生成")

    # ドロップダウンリストに表示するオプションを作成
    # 例: "{単元名} - {タイムスタンプ}"
    selection_options = [
        f"[{entry.get('タイムスタンプ', '日時不明')}] {entry.get('単元名', '単元名なし')} - {entry.get('単元内の授業タイトル', '授業タイトルなし')}"
        for entry in st.session_state.google_form_records
    ]
    
    selected_index = st.selectbox(
        "Excel化するフォーム回答を選択してください",
        options=range(len(selection_options)),
        format_func=lambda x: selection_options[x],
        key="selected_form_entry_for_excel"
    )

    if selected_index is not None:
        selected_form_entry = st.session_state.google_form_records[selected_index]
        
        # Excel生成ボタン
        if st.button(f"「{selected_form_entry.get('単元名', '選択された回答')}」のExcelをダウンロード", key="download_generated_excel"):
            excel_data = generate_excel_from_form_data(selected_form_entry)
            if excel_data:
                # ファイル名に単元名とタイムスタンプを含める
                unit_name_for_filename = selected_form_entry.get('単元名', '授業カード').replace(' ', '_').replace('/', '_')
                timestamp_for_filename = selected_form_entry.get('タイムスタンプ', '').split(' ')[0].replace('-', '') # 日付のみを使用
                download_filename = f"{unit_name_for_filename}_授業カード_{timestamp_for_filename}.xlsm"
                
                st.download_button(
                    label="✅ Excelファイルをダウンロード",
                    data=excel_data,
                    file_name=download_filename,
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                    help="Googleフォームの回答から生成されたExcelファイルをダウンロードします。"
                )
                st.success("指定されたフォーム回答からExcelファイルを生成しました！")
            else:
                st.error("Excelファイルの生成に失敗しました。上記のエラーメッセージを確認してください。")
    else:
        st.info("選択可能なフォーム回答がありません。")
else:
    st.info("まず「最新のフォーム回答を読み込む」ボタンを押して、Googleフォームのデータを取得してください。")

st.markdown("---")
