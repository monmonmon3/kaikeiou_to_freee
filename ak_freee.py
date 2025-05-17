import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import base64

st.title("会計王 to Freee CSV Converter")

# ファイルアップロード
uploaded_file = st.file_uploader("会計王の仕訳ファイルをアップロード", type=["xls", "xlsx"])

if uploaded_file is not None:
    # データ読み込み
    df = pd.read_excel(uploaded_file, header=1)

    # 必要な列を残して不要な列を削除
    df = df.drop(['行番号','借方部門コード','借方部門名称','貸方部門コード','貸方部門名称','借方事業分類コード','貸方事業分類コード',
                  '補助摘要','メモ','付箋１','付箋２','伝票種別','通番','貸方経過措置名称','取引摘要コード','補助摘要コード',
                  '借方科目コード','貸方科目コード','借方補助コード','貸方補助コード','借方事業分類コード','貸方事業分類コード'], axis=1)

    # 固定ヘッダー
    fixed_headers = ["日付","伝票番号","決算整理仕訳","借方勘定科目","借方補助科目","借方金額","借方税区分","借方税額",
                     "貸方勘定科目","貸方補助科目","貸方金額","貸方税区分","貸方税額","摘要"]

    df_freee = pd.DataFrame(columns=fixed_headers)

    # 列名のマッピング
    mapping = {'伝票番号':'伝票番号',
               '伝票日付':'日付',
               '借方科目名称':'借方勘定科目',
               '借方補助科目名称':'借方補助科目',
               '借方金額':'借方金額',
               '借方消費税':'借方税額',
               '貸方科目名称':'貸方勘定科目',
               '貸方補助科目名称':'貸方補助科目',
               '貸方金額':'貸方金額',
               '貸方消費税':'貸方税額',
               '取引摘要':'摘要'}

    # 転記処理
    for src_col, dest_col in mapping.items():
        if src_col in df.columns and dest_col in df_freee.columns:
            df_freee[dest_col] = df[src_col]

    # パターン辞書を作成
    patterns = {
        ('対象外', '0%', 80): '対象外',
        ('対象外', '0%', None): '対象外',
        ('対象外', '10%', 80): '対象外',
        ('対象外', '10%', None): '対象外',
        ('対象外', '8%', 80): '対象外',
        ('対象外', '8%', None): '対象外',
        ('対象外', '8%軽', 80): '対象外',
        ('対象外', '8%軽', None): '対象外',
        ('課売仕入', '10%', 80): '課対仕入（控80）10%',
        ('課売仕入', '10%', None): '課対仕入10%',
        ('課売仕入', '8%', 80): '課対仕入（控80）8%',
        ('課売仕入', '8%', None): '課対仕入8%',
        ('課売仕入', '8%軽', 80): '課対仕入（控80）8%（軽）',
        ('課売仕入', '8%軽', None): '課対仕入8%（軽）',
        ('非課仕入', '0%', 80): '非課仕入',
        ('非課仕入', '0%', None): '非課仕入',
        ('非課仕入', '10%', 80): '非課仕入',
        ('非課仕入', '10%', None): '非課仕入',
        ('非課仕入', '8%', 80): '非課仕入',
        ('非課仕入', '8%', None): '非課仕入',
        ('非課仕入', '8%軽', 80): '非課仕入',
        ('非課仕入', '8%軽', None): '非課仕入',
        ('課売返還', '10%', 80): '課売返五（控80）10%',
        ('課売返還', '10%', None): '課売返五10%',
        ('課売返還', '8%', 80): '課売返五（控80）8%',
        ('課売返還', '8%', None): '課売返五8%',
        ('課売返還', '8%軽', 80): '課売返五（控80）8%（軽）',
        ('課売返還', '8%軽', None): '課売返五8%（軽）',
        ('課税売上', '10%', 80): '課税売上（控80）10%',
        ('課税売上', '10%', None): '課税売上10%',
        ('課税売上', '8%', 80): '課税売上（控80）8%',
        ('課税売上', '8%', None): '課税売上8%',
        ('課税売上', '8%軽', 80): '課税売上（控80）8%（軽）',
        ('課税売上', '8%軽', None): '課税売上8%（軽）',
        ('非課売上', '0%', 80): '非課売上',
        ('非課売上', '0%', None): '非課売上',
        ('非課売上', '10%', 80): '非課売上',
        ('非課売上', '10%', None): '非課売上',
        ('非課売上', '8%', 80): '非課売上',
        ('非課売上', '8%', None): '非課売上',
        ('非課売上', '8%軽', 80): '非課売上',
        ('非課売上', '8%軽', None): '非課売上',
        ('課売仕返', '10%', 80): '課対仕返（控80）10%',
        ('課売仕返', '10%', None): '課対仕返10%',
        ('課売仕返', '8%', 80): '課対仕返（控80）8%',
        ('課売仕返', '8%', None): '課対仕返8%',
        ('課売仕返', '8%軽', 80): '課対仕返（控80）8%（軽）',
        ('課売仕返', '8%軽', None): '課対仕返8%（軽）'
    }

    # 任意のキー列を受け取ってC列の値を決定する関数
    def determine_d_value(row, key_col1, key_col2, key_col3, patterns):
        return patterns.get((row[key_col1], row[key_col2], row[key_col3]), 'その他')

    # nanをNoneに置き換える
    df['借方経過措置コード'] = df['借方経過措置コード'].replace({np.nan: None})
    df['貸方経過措置コード'] = df['貸方経過措置コード'].replace({np.nan: None})

    # 借方税区分と貸方税区分の決定
    df_freee['借方税区分'] = df.apply(determine_d_value, axis=1, key_col1='借方課税区分名称', key_col2='借方税率', key_col3= '借方経過措置コード', patterns=patterns)
    df_freee['貸方税区分'] = df.apply(determine_d_value, axis=1, key_col1='貸方課税区分名称', key_col2='貸方税率', key_col3= '貸方経過措置コード', patterns=patterns)

    # すべてのスペース（半角スペースおよび全角スペース）を削除する関数
    def remove_spaces(text):
        if isinstance(text, str):
            return text.replace(' ', '').replace('　', '')
        else:
            return text

    # 借方勘定科目と貸方勘定科目の値を文字列に変換し、余分なスペースを削除
    df_freee['借方勘定科目'] = df_freee['借方勘定科目'].astype(str).apply(remove_spaces)
    df_freee['貸方勘定科目'] = df_freee['貸方勘定科目'].astype(str).apply(remove_spaces)

    # 'nan'という文字列をNaNに変換
    df_freee['借方勘定科目'] = df_freee['借方勘定科目'].replace('nan', None)
    df_freee['貸方勘定科目'] = df_freee['貸方勘定科目'].replace('nan', None)

    # '借方勘定科目'が欠損値の場合、その行の'借方税区分'の値を削除（NaNに置き換え）
    df_freee.loc[df_freee['借方勘定科目'].isna(), '借方税区分'] = None
    df_freee.loc[df_freee['貸方勘定科目'].isna(), '貸方税区分'] = None

    # CSVファイルのダウンロードリンクを作成
    def convert_df_to_csv(df):
        return df.to_csv(encoding='cp932', index=False)

    csv = convert_df_to_csv(df_freee)
    b64 = base64.b64encode(csv.encode()).decode()  # CSVファイルをbase64でエンコード

    st.download_button(
        label="Download CSV",
        data=BytesIO(csv.encode('cp932')),
        file_name='freee_import.csv',
        mime='text/csv'
    )
