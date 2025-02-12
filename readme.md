# Calcolo100 CSV2Excel

![app_image](/images/スクリーンショット%202025-01-14%20222035.png)

Nipo製計算タイムレコーダー カルコロ100向けのCSV変換ツールです。

本ツールは、有志による制作であり、公式に配布しているものではありません。

## 機能概要

- カルコロ100の打刻データ（CSVファイル）をExcelファイルに変換
- カード番号と従業員名のマッピング機能
- 日本語CSVファイルの自動エンコーディング検出（UTF-8, CP932, Shift-JIS対応）
- グラフィカルユーザーインターフェース（GUI）による簡単操作
- 時間データの合計計算機能

## 必要要件

- Python 3.8以上(パッケージ版の場合は不要)

## インストール方法

パッケージ版は、リリースよりダウンロードしてください

1. リポジトリをクローン:
   ```bash
   git clone https://github.com/Mega-Gorilla/Calcolo100_CSV2Excel.git
   cd Calcolo100_CSV2Excel
   ```

2. 依存ライブラリをインストール:
   ```bash
   pip install -r requirements.txt
   ```

## 使用方法

### GUIアプリケーションの起動

```bash
python app.py
```

### 基本的な操作手順

1. マッピングファイルの設定
   - 「マッピングファイル選択」ボタンでCSVファイルを選択
   - テーブル上で直接編集可能
   - 「マッピングを保存」ボタンで変更を保存

2. 入力ファイルの選択
   - 「入力ファイル選択」ボタンでカルコロ100のCSVファイルを選択

3. 出力ファイルの設定
   - 「出力ファイル選択」ボタンで保存先のExcelファイルを指定

4. 変換の実行
   - 「変換開始」ボタンをクリック
   - 処理の進捗状況がログウィンドウに表示

## ファイルフォーマット

### マッピングファイル（CSV）

カード番号と従業員名の対応を定義するCSVファイル。

```csv
カード番号,名前
0001,山田太郎
0002,鈴木花子
```

### 入力CSVファイル（カルコロ100）

カルコロ100の標準CSVフォーマット:

| No | フィールド名 | 説明                  | フォーマット |
|----|------------|----------------------|-------------|
| 1  | カード番号  | 従業員ID/カード番号     | 数値4桁     |
| 2  | 区分       | 記録種別              | 数値1桁     |
| 3  | 年月日     | 日付                  | YY/MM/DD   |
| 4  | 入1時刻    | 1回目の出勤時刻        | HH:MM      |
| 5  | 入1異例    | 1回目の出勤例外コード   | 数値2桁     |
| 6  | 退1時刻    | 1回目の退勤時刻        | HH:MM      |
| 7  | 退1異例    | 1回目の退勤例外コード   | 数値2桁     |
| 8  | 入2時刻    | 2回目の出勤時刻        | HH:MM      |
| 9  | 入2異例    | 2回目の出勤例外コード   | 数値2桁     |
| 10 | 退2時刻    | 2回目の退勤時刻        | HH:MM      |
| 11 | 退2異例    | 2回目の退勤例外コード   | 数値2桁     |
| 12 | 時数1      | 1回目の勤務時間        | HHH:MM     |
| 13 | 時数2      | 2回目の勤務時間        | HHH:MM     |

### 異例コード一覧

| 区分   | 平日 | 休日 | 説明               |
|--------|------|------|-------------------|
| 出勤   | 00   | 10   | 通常の出勤        |
| 退勤   | 00   | 10   | 通常の退勤        |
| 早出   | 01   | 11   | 通常より早い出勤   |
| 遅刻   | 02   | 12   | 遅刻での出勤      |
| 外出   | 03   | 13   | 一時的な外出      |
| 再入   | 04   | 14   | 外出からの戻り    |
| 早退   | 05   | 15   | 通常より早い退勤   |
| 残業   | 06   | 16   | 残業での勤務      |
| 深夜   | 09   | 19   | 深夜勤務          |
| 徹夜   | 07   | 17   | 終夜勤務          |

## 注意事項

- 本ツールは非公式のツールです
- 処理前に必ずデータのバックアップを取ってください
- 大量のデータを処理する場合は、処理に時間がかかる場合があります
- エラーが発生した場合は、ログウィンドウの内容を確認してください

## トラブルシューティング

1. ファイルが読み込めない場合
   - ファイルの文字エンコーディングを確認
   - CSVファイルのフォーマットが正しいか確認
   - ファイルが他のプログラムで開かれていないか確認

2. マッピングが正しく機能しない場合
   - マッピングファイルのフォーマットを確認
   - カード番号が4桁になっているか確認
   - カラム名が正しいか確認

## 貢献

バグ報告や機能改善の提案は、Issueやプルリクエストでお願いします。