import streamlit as st
import pandas as pd
import datetime
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="ポイ活通帳Web", layout="wide")

# ▼ご自身のスプレッドシートURLに書き換えてください
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1hcBElwsxmWfjU7Zhxqwuy4kMDDJm_yVnI6m2f1z2rAM/edit?gid=1564950317#gid=1564950317"

def check_password():
    # ▼ご自身の好きなパスワードに書き換えてください
    MY_PASSWORD = "tkhr_WD&030118"

    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        st.title("🔒 ワイのポイポイポイぽポイポイぽぴー ログイン")
        pwd = st.text_input("パスワードを入力してください", type="password")
        if st.button("ログイン"):
            if pwd == MY_PASSWORD:
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("パスワードが違います")
        return False
    return True

if not check_password():
    st.stop()

st.title("💳 ポイ活通帳")

USES = ["買い物", "貯", "運用推移", "キャンペーン", "ポイント還元", "ポイント利用", "チャージ", "失効", "発券・交換など"]
STATUSES = ["確定", "処理中", "拒否", "調査"]

# --- Google Sheets 接続設定 ---
@st.cache_resource
def init_connection():
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )
    return gspread.authorize(creds)

try:
    gc = init_connection()
    sh = gc.open_by_url(SPREADSHEET_URL)
    ws_ledger = sh.worksheet("通帳データ")
    # 今回追加した2つのシートを読み込む
    ws_summary = sh.worksheet("獲得集計表")
    ws_balance = sh.worksheet("ポイント残高表")
    ws_point = sh.worksheet("ポイント定義")
except Exception as e:
    st.error("スプレッドシートの読み込みに失敗しました。「通帳データ」「ポイント定義」「獲得集計表」「ポイント残高表」の4つのタブがすべて作成されているか確認してください。")
    st.stop()

# --- データの読み書き関数 ---
def load_ledger():
    try:
        records = ws_ledger.get_all_records()
        if records:
            df = pd.DataFrame(records)
            df['日付'] = pd.to_datetime(df['日付'])
            return df
        else:
            return pd.DataFrame(columns=["取引通番", "日付", "用途", "ポイント名", "取得", "利用", "状態", "メモ"])
    except Exception:
        # シートが完全に空の場合のエラーを回避
        return pd.DataFrame(columns=["取引通番", "日付", "用途", "ポイント名", "取得", "利用", "状態", "メモ"])

def load_point_defs():
    try:
        records = ws_point.get_all_records()
        if records:
            return pd.DataFrame(records)
        else:
            raise ValueError("シートが空です")
    except Exception:
        # エラー時や空の場合は初期データを作成して返す
        default_df = pd.DataFrame({
            "ポイント名": ["V", "V期間限定", "V運用", "楽天", "d", "PayPay"],
            "種類": ["一般", "期間限定", "運用", "一般", "一般", "一般"],
            "初期残高": [0, 0, 0, 0, 0, 0]
        })
        save_points(default_df)
        return default_df

def save_ledger(df):
    df_save = df.copy()
    df_save['日付'] = pd.to_datetime(df_save['日付']).dt.strftime('%Y-%m-%d')
    df_save = df_save.fillna("")
    for col in ["取引通番", "取得", "利用"]:
        if col in df_save.columns:
            df_save[col] = pd.to_numeric(df_save[col], errors='coerce').fillna(0).astype(int)
    ws_ledger.clear()
    ws_ledger.update(values=[df_save.columns.values.tolist()] + df_save.values.tolist(), range_name="A1")

def save_points(df):
    df_save = df.copy().fillna("")
    if "初期残高" in df_save.columns:
        df_save["初期残高"] = pd.to_numeric(df_save["初期残高"], errors='coerce').fillna(0).astype(int)
    ws_point.clear()
    ws_point.update(values=[df_save.columns.values.tolist()] + df_save.values.tolist(), range_name="A1")

# --- 集計表と残高表の保存関数（今回追加） ---
def save_summary(pivot_df):
    df_save = pivot_df.copy()
    # 列名が2段構え（取得・2026-06など）になっているのを平らにする
    if isinstance(df_save.columns, pd.MultiIndex):
        df_save.columns = [f"{col[0]}_{col[1]}" if col[1] else col[0] for col in df_save.columns]
    df_save = df_save.reset_index().fillna(0)
    
    # スプレッドシートエラー防止のため、数字をすべて標準の数値型に変換
    for col in df_save.columns:
        if pd.api.types.is_numeric_dtype(df_save[col]):
            df_save[col] = pd.to_numeric(df_save[col], errors='coerce').fillna(0).astype(int)
            
    ws_summary.clear()
    ws_summary.update(values=[df_save.columns.values.tolist()] + df_save.values.tolist(), range_name="A1")

def save_balance(df_balance):
    df_save = df_balance.copy().fillna(0)
    for col in ["残高", "現金換算額"]:
        if col in df_save.columns:
            df_save[col] = pd.to_numeric(df_save[col], errors='coerce').fillna(0).astype(int)
            
    ws_balance.clear()
    ws_balance.update(values=[df_save.columns.values.tolist()] + df_save.values.tolist(), range_name="A1")

# --- メイン画面構築 ---
df = load_ledger()
point_df = load_point_defs()
POINT_NAMES = point_df["ポイント名"].tolist()

menu = st.sidebar.radio("メニュー", ["通帳入力・履歴", "獲得集計・増減推移表", "ポイント残高表", "ポイント定義・初期残高設定"])

if menu == "通帳入力・履歴":
    st.header("📝 通帳データ入力")
    if not POINT_NAMES:
        st.warning("設定画面でポイントを登録してください。")
    else:
        with st.form("ledger_form", clear_on_submit=True):
            col1, col2, col3, col4 = st.columns(4)
            date = col1.date_input("日付", datetime.date.today())
            use = col2.selectbox("用途", USES)
            point_name = col3.selectbox("ポイント名", POINT_NAMES)
            status = col4.selectbox("状態", STATUSES)
            
            col5, col6, col7 = st.columns([1, 1, 2])
            acquired_str = col5.text_input("取得", value="0")
            used_str = col6.text_input("利用", value="0")
            memo = col7.text_input("メモ")
            submit = st.form_submit_button("通帳に追記する")
            
            if submit:
                try:
                    acquired = abs(int(acquired_str))
                    used = abs(int(used_str))
                    new_id = 1 if df.empty else df['取引通番'].max() + 1
                    new_row = pd.DataFrame({
                        "取引通番": [new_id], "日付": [pd.to_datetime(date)], "用途": [use],
                        "ポイント名": [point_name], "取得": [acquired], "利用": [used],
                        "状態": [status], "メモ": [memo]
                    })
                    new_df = pd.concat([df, new_row], ignore_index=True)
                    save_ledger(new_df)
                    st.success("追加しました！")
                    st.rerun()
                except ValueError:
                    st.error("半角数字を入力してください。")

    st.markdown("---")
    st.header("📖 履歴一覧（編集・削除）")
    if not df.empty:
        df['年月'] = df['日付'].dt.to_period('M').astype(str)
        month_list = ["すべて"] + sorted(df['年月'].unique().tolist(), reverse=True)
        selected_month = st.selectbox("表示する月を選択してください", month_list)

        display_df = df.copy() if selected_month == "すべて" else df[df['年月'] == selected_month].copy()
        display_df['日付'] = display_df['日付'].dt.strftime('%Y-%m-%d')
        display_df = display_df.drop(columns=['年月']).sort_values("取引通番", ascending=False)
        
        st.write(f"**{selected_month} のデータ一覧**")
        edited_df = st.data_editor(
            display_df, num_rows="dynamic", use_container_width=True,
            column_config={"取引通番": st.column_config.NumberColumn(disabled=True)}
        )

        if st.button("変更を保存する"):
            edited_df['日付'] = pd.to_datetime(edited_df['日付'])
            edited_df['取得'] = pd.to_numeric(edited_df['取得']).abs()
            edited_df['利用'] = pd.to_numeric(edited_df['利用']).abs()

            if selected_month != "すべて":
                other_months_df = df[df['年月'] != selected_month].drop(columns=['年月'])
                final_df = pd.concat([other_months_df, edited_df], ignore_index=True)
            else:
                final_df = edited_df

            save_ledger(final_df)
            st.success("保存しました！")
            st.rerun()

elif menu == "獲得集計・増減推移表":
    st.header("📊 月別 獲得集計 ＆ 増減推移統括表")
    if not df.empty:
        df['年月'] = df['日付'].dt.to_period('M').astype(str)
        
        # クリックで展開するグラフ
        with st.expander("📈 月別の獲得・利用グラフを表示する"):
            graph_df = df.groupby('年月')[['取得', '利用']].sum()
            col1, col2 = st.columns([1, 1])
            with col1:
                st.bar_chart(graph_df)
                
        st.subheader("① 獲得・利用 集計表")
        pivot = pd.pivot_table(
            df, index="ポイント名", columns="年月", values=["取得", "利用"], 
            aggfunc="sum", fill_value=0, margins=True, margins_name="合計"
        )
        idx = pivot.index.tolist()
        if '合計' in idx:
            idx.remove('合計')
            pivot = pivot.reindex(['合計'] + idx)
        st.dataframe(pivot, use_container_width=True)
        
        # --- スプレッドシート保存ボタン ---
        if st.button("💾 この獲得・利用集計表をスプレッドシートに保存"):
            save_summary(pivot)
            st.success("スプレッドシートの「獲得集計表」タブに出力しました！")

        st.markdown("---")
        st.subheader("② 増減推移統括表（月間トータル）")
        df_net = df.copy()
        df_net['増減'] = df_net['取得'] - df_net['利用']
        pivot_net = pd.pivot_table(
            df_net, index="ポイント名", columns="年月", values="増減", 
            aggfunc="sum", fill_value=0, margins=True, margins_name="合計"
        )
        idx_net = pivot_net.index.tolist()
        if '合計' in idx_net:
            idx_net.remove('合計')
            pivot_net = pivot_net.reindex(['合計'] + idx_net)
        
        def color_red(val):
            return 'color: red' if isinstance(val, (int, float)) and val < 0 else ''

        try:
            st.dataframe(pivot_net.style.map(color_red), use_container_width=True)
        except AttributeError:
            st.dataframe(pivot_net.style.applymap(color_red), use_container_width=True)

elif menu == "ポイント残高表":
    st.header("💰 現在のポイント残高")
    if not df.empty:
        ledger_bal = df.groupby("ポイント名")[["取得", "利用"]].sum().reset_index()
        balance_df = pd.merge(point_df, ledger_bal, on="ポイント名", how="left").fillna(0)
    else:
        balance_df = point_df.copy()
        balance_df["取得"] = balance_df["利用"] = 0
        
    balance_df["残高"] = balance_df["初期残高"] + balance_df["取得"] - balance_df["利用"]
    balance_df["現金換算額"] = balance_df.apply(lambda row: row["残高"] * (5.0 if "Global(三菱UFJ)" in row["ポイント名"] else 1.0), axis=1)
    
    display_df = balance_df[balance_df["残高"] != 0][["ポイント名", "残高", "現金換算額"]]
    if not display_df.empty:
        st.metric("総現金換算額", f"¥{display_df['現金換算額'].sum():,.0f}")
        sorted_display_df = display_df.sort_values("現金換算額", ascending=False)
        st.dataframe(sorted_display_df, use_container_width=True)
        
        # --- スプレッドシート保存ボタン ---
        if st.button("💾 このポイント残高表をスプレッドシートに保存"):
            save_balance(sorted_display_df)
            st.success("スプレッドシートの「ポイント残高表」タブに出力しました！")

elif menu == "ポイント定義・初期残高設定":
    st.header("⚙️ ポイント定義・初期残高設定")
    edited_point_df = st.data_editor(
        point_df, num_rows="dynamic", use_container_width=True,
        column_config={
            "ポイント名": st.column_config.TextColumn("ポイント名", required=True),
            "種類": st.column_config.SelectboxColumn("種類", options=["一般", "期間限定", "運用", "ギフトカード", "航空マイル", "電子マネー"], required=True),
            "初期残高": st.column_config.NumberColumn("初期残高", min_value=0, step=1, default=0)
        }
    )
    if st.button("設定を保存する"):
        edited_point_df = edited_point_df.dropna(subset=["ポイント名"])
        if edited_point_df["ポイント名"].duplicated().any():
            st.error("ポイント名が重複しています。")
        else:
            save_points(edited_point_df)
            st.success("保存しました！")
