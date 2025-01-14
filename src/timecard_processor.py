import pandas as pd
from datetime import datetime, timedelta
import numpy as np
from openpyxl import load_workbook
from tqdm import tqdm

class TimecardProcessor:
    def __init__(self):
        self.exception_codes = {
            '00': '通常', '01': '早出', '02': '遅刻', '03': '外出',
            '04': '再入', '05': '早退', '06': '残業', '07': '徹夜',
            '09': '深夜', '10': '休日出勤', '11': '休日早出',
            '12': '休日遅刻', '13': '休日外出', '14': '休日再入',
            '15': '休日早退', '16': '休日残業', '17': '休日徹夜',
            '19': '休日深夜'
        }
        self.name_mapping = {}

    def load_name_mapping(self, mapping_file):
        """カード番号と名前の対応表を読み込む"""
        try:
            print("名前マッピングファイルを読み込んでいます...")
            # CSVファイルを読み込み（カード番号と名前の2列を想定）
            mapping_df = pd.read_csv(mapping_file)
            
            # カラム名を確認し、必要に応じて調整
            card_col = mapping_df.columns[0]  # カード番号のカラム
            name_col = mapping_df.columns[1]  # 名前のカラム
            
            # マッピング辞書を作成
            self.name_mapping = dict(zip(mapping_df[card_col].astype(str), mapping_df[name_col]))
            print(f"名前マッピングの読み込みが完了しました（{len(self.name_mapping)}件）")
            
        except Exception as e:
            print(f"名前マッピングファイルの読み込み中にエラーが発生しました: {str(e)}")
            raise

    def convert_duration_to_minutes(self, duration_str):
        """時数文字列を分単位に変換"""
        if not duration_str or duration_str.strip() == '':
            return 0
        try:
            hours, minutes = map(int, duration_str.split(':'))
            return hours * 60 + minutes
        except:
            return 0

    def convert_minutes_to_duration(self, minutes):
        """分を時数文字列に変換"""
        if minutes == 0:
            return ''
        hours = minutes // 60
        remaining_minutes = minutes % 60
        return f'{hours:03d}:{remaining_minutes:02d}'

    def process_csv(self, input_file, output_file, mapping_file=None):
        """CSVファイルを処理してExcelファイルに出力"""
        print("タイムカードデータ処理を開始します...")

        # 名前マッピングファイルが指定されている場合は読み込む
        if mapping_file:
            self.load_name_mapping(mapping_file)
        
        # CSVファイルを読み込み
        print("CSVファイルを読み込んでいます...")
        df = pd.read_csv(input_file, skiprows=1, header=None,
                        names=['カード番号', '区分', '年月日',
                              '入1時刻', '入1異例', '退1時刻', '退1異例',
                              '入2時刻', '入2異例', '退2時刻', '退2異例',
                              '時数1', '時数2', '空白'],
                        skipinitialspace=True)
        
        # 不要な列を削除
        df = df.drop('空白', axis=1)

        # カード番号を名前に変換
        if self.name_mapping:
            print("カード番号を名前に変換中...")
            # カード番号を文字列として扱う
            df['カード番号'] = df['カード番号'].astype(str).str.zfill(4)
            # 名前カラムを追加
            df['名前'] = df['カード番号'].map(self.name_mapping)
            # マッピングできなかった場合はカード番号を維持
            df['名前'] = df.apply(lambda row: row['名前'] if pd.notna(row['名前']) else f"未登録(カード番号:{row['カード番号']})", axis=1)
            # カード番号カラムを削除し、名前カラムを先頭に
            df = df.drop('カード番号', axis=1)
            df = df.rename(columns={'名前': 'カード番号'})

        # データのクリーニング
        print("データのクリーニングを実行中...")
        with tqdm(total=len(df.columns), desc="列の処理") as pbar:
            for col in df.columns:
                if df[col].dtype == object:
                    df[col] = df[col].str.strip()
                pbar.update(1)

        # 年月日を日付のみに変換
        print("日付データを変換中...")
        df['年月日'] = pd.to_datetime(df['年月日'].astype(str), format='%y/%m/%d').dt.date
        
        # 時刻データを処理
        print("時刻データを処理中...")
        time_columns = ['入1時刻', '退1時刻', '入2時刻', '退2時刻']
        with tqdm(total=len(time_columns), desc="時刻の処理") as pbar:
            for col in time_columns:
                df[col] = df[col].apply(lambda x: x if pd.notna(x) and str(x).strip() else '')
                pbar.update(1)

        # 時数を処理
        print("時数データを処理中...")
        def format_duration(duration_str):
            if pd.isna(duration_str) or str(duration_str).strip() == '':
                return ''
            try:
                hours, minutes = map(int, duration_str.split(':'))
                return f'{hours:03d}:{minutes:02d}'
            except:
                return ''

        duration_columns = ['時数1', '時数2']
        with tqdm(total=len(duration_columns), desc="時数の処理") as pbar:
            for col in duration_columns:
                df[col] = df[col].apply(format_duration)
                pbar.update(1)

        # 時数1と時数2の合計を計算
        print("合計時数を計算中...")
        with tqdm(total=len(df), desc="合計時数の計算") as pbar:
            df['合計時数'] = df.apply(
                lambda row: self.convert_minutes_to_duration(
                    self.convert_duration_to_minutes(row['時数1']) +
                    self.convert_duration_to_minutes(row['時数2'])
                ),
                axis=1
            )
            pbar.update(len(df))

        # 異例コードを変換
        print("異例コードを変換中...")
        exception_columns = ['入1異例', '退1異例', '入2異例', '退2異例']
        with tqdm(total=len(exception_columns), desc="異例コードの変換") as pbar:
            for col in exception_columns:
                df[col] = df[col].astype(str).map(self.exception_codes).fillna('')
                pbar.update(1)

        # Excelファイルとして出力
        print("Excelファイルに出力中...")
        df.to_excel(output_file, sheet_name='勤怠データ', index=False)

        # Excelファイルを開いてフォーマットを設定
        print("Excelファイルのフォーマットを設定中...")
        wb = load_workbook(output_file)
        ws = wb['勤怠データ']

        # 列のインデックスを取得
        date_col_idx = df.columns.get_loc('年月日') + 1
        time_col_indices = [df.columns.get_loc(col) + 1 for col in time_columns]

        # セルの書式設定
        total_rows = ws.max_row - 1  # ヘッダーを除く
        with tqdm(total=total_rows, desc="セル書式の設定") as pbar:
            for row in range(2, ws.max_row + 1):
                # 日付書式の設定
                ws.cell(row=row, column=date_col_idx).number_format = 'YYYY/MM/DD'
                
                # 時刻書式の設定
                for col_idx in time_col_indices:
                    cell = ws.cell(row=row, column=col_idx)
                    if cell.value:
                        cell.number_format = 'HH:MM'
                
                pbar.update(1)

        # 変更を保存
        print("ファイルを保存中...")
        wb.save(output_file)
        
        print("処理が完了しました！")
        return df

# 使用例
if __name__ == "__main__":
    processor = TimecardProcessor()
    try:
        # 名前マッピングファイルを指定して処理を実行
        df = processor.process_csv(
            input_file='input.csv',
            output_file='output.xlsx',
            mapping_file='name_mapping.csv'  # カード番号と名前の対応表CSV
        )
        print("正常に処理が完了しました。")
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")