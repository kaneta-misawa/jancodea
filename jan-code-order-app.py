import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import cv2
from pyzbar.pyzbar import decode
import numpy as np
from datetime import datetime
import time
import os
from PIL import Image
import io

# タイトルとアプリの説明
st.title('発注管理アプリ')
st.subheader('JANコードを読み取って発注履歴を記録')

# セッション状態の初期化
if 'product_info' not in st.session_state:
    st.session_state.product_info = None
if 'jan_code' not in st.session_state:
    st.session_state.jan_code = None

# サイドバーにファイルアップローダーを設置（サービスアカウントのJSONキー用）
st.sidebar.header("Google認証設定")
uploaded_file = st.sidebar.file_uploader("サービスアカウントのJSONキーをアップロード", type=['json'])

# 認証関数
def authenticate_google_sheets(json_file):
    """
    Googleスプレッドシートに接続するための認証を行う関数
    
    Parameters:
    json_file: アップロードされたサービスアカウントのJSONキーファイル
    
    Returns:
    client: 認証済みのgspreadクライアント
    """
    try:
        # JSONデータを読み込む
        json_content = json_file.getvalue().decode()
        
        # 一時ファイルに保存
        temp_json_path = "temp_key.json"
        with open(temp_json_path, "w") as f:
            f.write(json_content)
        
        # APIの認証スコープ
        scope = ['https://spreadsheets.google.com/feeds',
                'https://www.googleapis.com/auth/drive']
        
        # 認証情報の取得
        credentials = ServiceAccountCredentials.from_json_keyfile_name(temp_json_path, scope)
        
        # gspreadクライアントの作成
        client = gspread.authorize(credentials)
        
        # 一時ファイルの削除
        os.remove(temp_json_path)
        
        return client
    
    except Exception as e:
        st.error(f"認証エラー: {e}")
        return None

# スプレッドシートからデータを取得する関数
def get_master_data(client, spreadsheet_name, worksheet_name):
    """
    マスターデータをスプレッドシートから取得する関数
    
    Parameters:
    client: 認証済みのgspreadクライアント
    spreadsheet_name: スプレッドシートの名前
    worksheet_name: ワークシートの名前
    
    Returns:
    df: 取得したデータのDataFrame
    """
    try:
        # スプレッドシートを開く
        spreadsheet = client.open(spreadsheet_name)
        
        # ワークシートを選択
        worksheet = spreadsheet.worksheet(worksheet_name)
        
        # データを取得
        data = worksheet.get_all_records()
        
        # DataFrameに変換
        df = pd.DataFrame(data)
        
        return df
    
    except Exception as e:
        st.error(f"データ取得エラー: {e}")
        return None

# JANコードから商品情報を検索する関数
def find_product_by_jan(master_df, jan_code):
    """
    JANコードから商品情報を検索する関数
    
    Parameters:
    master_df: マスターデータのDataFrame
    jan_code: 検索するJANコード
    
    Returns:
    product_info: 商品情報の辞書（見つからない場合はNone）
    """
    if master_df is None or jan_code is None:
        return None
    
    # JANコードで検索（文字列として扱う）
    jan_code = str(jan_code).strip()
    master_df['JANコード'] = master_df['JANコード'].astype(str)
    
    # マッチする行を検索
    matched_products = master_df[master_df['JANコード'] == jan_code]
    
    if len(matched_products) == 0:
        return None
    
    # 最初のマッチを取得
    product = matched_products.iloc[0]
    
    # 必要な情報を辞書に格納
    product_info = {
        'jan_code': jan_code,
        'product_name': product['商品名'],
        'unit_price': product['単価'],
        'min_order': product['最低発注単位']
    }
    
    return product_info

# 発注履歴をスプレッドシートに書き込む関数
def write_order_history(client, spreadsheet_name, worksheet_name, order_data):
    """
    発注履歴をスプレッドシートに書き込む関数
    
    Parameters:
    client: 認証済みのgspreadクライアント
    spreadsheet_name: スプレッドシートの名前
    worksheet_name: ワークシートの名前
    order_data: 書き込む発注データの辞書
    
    Returns:
    success: 書き込み成功したかどうかのブール値
    """
    try:
        # スプレッドシートを開く
        spreadsheet = client.open(spreadsheet_name)
        
        # ワークシートを選択
        worksheet = spreadsheet.worksheet(worksheet_name)
        
        # 発注データを行として追加
        row = [
            order_data['date_time'],
            order_data['jan_code'],
            order_data['product_name'],
            order_data['quantity'],
            order_data['total_price']
        ]
        
        worksheet.append_row(row)
        
        return True
    
    except Exception as e:
        st.error(f"発注履歴書き込みエラー: {e}")
        return False

# カメラからJANコードを読み取る関数
def scan_jan_code():
    """
    カメラからJANコードを読み取る関数
    
    Returns:
    jan_code: 読み取ったJANコード（読み取れない場合はNone）
    """
    st.write("カメラが起動します...")
    
    # カメラの設定
    cap = cv2.VideoCapture(0)  # 0はデフォルトのカメラ
    
    if not cap.isOpened():
        st.error("カメラにアクセスできません。")
        return None
    
    st.write("JANコードをカメラにかざしてください。")
    stframe = st.empty()
    
    # 5秒間スキャンを試みる
    start_time = time.time()
    timeout = 10  # 10秒のタイムアウト
    
    while time.time() - start_time < timeout:
        ret, frame = cap.read()
        
        if not ret:
            st.error("フレームの取得に失敗しました。")
            break
        
        # フレームをRGBに変換（Streamlitでの表示用）
        rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        
        # ストリーミング表示
        stframe.image(rgb_frame, channels="RGB", caption="カメラプレビュー")
        
        # JANコードの検出
        decoded_objects = decode(frame)
        
        for obj in decoded_objects:
            # 見つかったコードの種類とデータ
            barcode_type = obj.type
            barcode_data = obj.data.decode('utf-8')
            
            # JANコードが見つかった場合
            if barcode_type in ['EAN13', 'EAN8', 'UPCA', 'UPCE']:
                cap.release()
                return barcode_data
    
    # タイムアウトしたらカメラを解放
    cap.release()
    st.warning("JANコードを検出できませんでした。")
    return None

# 画像ファイルからJANコードを読み取る関数（カメラ代替用）
def scan_jan_from_image(uploaded_image):
    """
    アップロードされた画像からJANコードを読み取る関数
    
    Parameters:
    uploaded_image: アップロードされた画像ファイル
    
    Returns:
    jan_code: 読み取ったJANコード（読み取れない場合はNone）
    """
    try:
        # 画像を読み込む
        image = Image.open(uploaded_image)
        image_array = np.array(image)
        
        # 画像を表示
        st.image(image, caption="アップロードされた画像", use_column_width=True)
        
        # JANコードの検出
        decoded_objects = decode(image_array)
        
        for obj in decoded_objects:
            # 見つかったコードの種類とデータ
            barcode_type = obj.type
            barcode_data = obj.data.decode('utf-8')
            
            # JANコードが見つかった場合
            if barcode_type in ['EAN13', 'EAN8', 'UPCA', 'UPCE']:
                return barcode_data
        
        st.warning("画像からJANコードを検出できませんでした。")
        return None
        
    except Exception as e:
        st.error(f"画像処理エラー: {e}")
        return None

# メイン処理
if uploaded_file:
    # スプレッドシートの設定
    st.sidebar.subheader("スプレッドシート設定")
    spreadsheet_name = st.sidebar.text_input("スプレッドシート名", "発注管理")
    master_sheet_name = st.sidebar.text_input("マスターデータのシート名", "商品マスター")
    history_sheet_name = st.sidebar.text_input("発注履歴のシート名", "発注履歴")
    
    # Google認証
    client = authenticate_google_sheets(uploaded_file)
    
    if client:
        st.sidebar.success("Google認証に成功しました！")
        
        # タブ作成
        tab1, tab2 = st.tabs(["JANコード読み取り", "手動入力"])
        
        with tab1:
            st.header("JANコード読み取り")
            
            # 読み取り方法の選択
            scan_method = st.radio(
                "JANコードの読み取り方法を選択:",
                ["カメラで読み取る", "画像をアップロード"]
            )
            
            jan_code = None
            
            if scan_method == "カメラで読み取る":
                if st.button("カメラを起動"):
                    jan_code = scan_jan_code()
                    if jan_code:
                        st.session_state.jan_code = jan_code
                        st.success(f"JANコード {jan_code} を読み取りました！")
            
            else:  # 画像をアップロード
                uploaded_image = st.file_uploader("JANコードが写っている画像をアップロード", type=['jpg', 'jpeg', 'png'])
                if uploaded_image:
                    jan_code = scan_jan_from_image(uploaded_image)
                    if jan_code:
                        st.session_state.jan_code = jan_code
                        st.success(f"JANコード {jan_code} を読み取りました！")
        
        with tab2:
            st.header("JANコード手動入力")
            manual_jan = st.text_input("JANコードを入力してください:")
            
            if st.button("検索"):
                if manual_jan:
                    st.session_state.jan_code = manual_jan
                    st.success(f"JANコード {manual_jan} を入力しました！")
        
        # JANコードが取得できたらマスターデータから商品情報を取得
        if st.session_state.jan_code:
            # マスターデータの取得
            master_df = get_master_data(client, spreadsheet_name, master_sheet_name)
            
            if master_df is not None:
                # 商品情報の検索
                product_info = find_product_by_jan(master_df, st.session_state.jan_code)
                
                if product_info:
                    st.session_state.product_info = product_info
                    
                    # 商品情報の表示
                    st.subheader("商品情報")
                    st.write(f"**商品名**: {product_info['product_name']}")
                    st.write(f"**単価**: {product_info['unit_price']}円")
                    st.write(f"**最低発注単位**: {product_info['min_order']}個")
                    
                    # 数量入力フォーム
                    st.subheader("発注数量")
                    quantity = st.number_input(
                        "数量を入力してください:",
                        min_value=product_info['min_order'],
                        step=product_info['min_order'],
                        value=product_info['min_order']
                    )
                    
                    # 数量が最低発注単位の倍数かチェック
                    is_valid_quantity = quantity % product_info['min_order'] == 0
                    
                    if not is_valid_quantity:
                        st.warning(f"数量は最低発注単位({product_info['min_order']}個)の倍数にしてください。")
                    
                    # 合計金額の計算と表示
                    total_price = quantity * product_info['unit_price']
                    st.subheader("合計金額")
                    st.write(f"**合計**: {total_price:,}円")
                    
                    # 発注ボタン
                    if st.button("発注する"):
                        if is_valid_quantity:
                            # 発注データの準備
                            order_data = {
                                'date_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                'jan_code': product_info['jan_code'],
                                'product_name': product_info['product_name'],
                                'quantity': quantity,
                                'total_price': total_price
                            }
                            
                            # 発注履歴に書き込み
                            if write_order_history(client, spreadsheet_name, history_sheet_name, order_data):
                                st.success("発注が完了しました！")
                                
                                # 状態をリセット
                                st.session_state.product_info = None
                                st.session_state.jan_code = None
                                
                                # 発注完了メッセージ
                                st.balloons()
                            else:
                                st.error("発注の記録に失敗しました。もう一度お試しください。")
                        else:
                            st.error("数量が最低発注単位の倍数ではありません。")
                else:
                    st.error(f"JANコード {st.session_state.jan_code} に該当する商品が見つかりませんでした。")
            else:
                st.error("マスターデータの取得に失敗しました。")
    else:
        st.sidebar.error("Google認証に失敗しました。JSONキーを確認してください。")
else:
    st.info("サイドバーからサービスアカウントのJSONキーをアップロードしてください。")
    
    # 使い方の説明
    st.subheader("アプリの使い方")
    st.markdown("""
    1. サイドバーからGoogleサービスアカウントのJSONキーをアップロードします。
    2. スプレッドシートの名前とシート名を設定します。
    3. JANコードをカメラで読み取るか、画像をアップロードするか、手動で入力します。
    4. 商品情報が表示されたら、発注数量を入力します。
    5. 「発注する」ボタンをクリックして発注を確定します。
    
    **事前準備**
    - Googleサービスアカウントを作成し、JSONキーをダウンロードしておく必要があります。
    - スプレッドシートに以下のシートを用意しておきます：
      - 商品マスター: JANコード、商品名、単価、最低発注単位のカラムを持つシート
      - 発注履歴: 日時、JANコード、商品名、数量、金額のカラムを持つシート
    """)
