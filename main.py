from src.TimecardProcessor import TimecardProcessor

# プロセッサーのインスタンス化
processor = TimecardProcessor()

# CSVファイルの処理とExcel出力
df = processor.process_csv('dakoku_all.csv', 'output.xlsx')