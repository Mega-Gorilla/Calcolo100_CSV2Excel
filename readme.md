# Calcolo100 CSV2Excel

Nipo製計算タイムレコーダー カルコロ100向けのCSV変換スクリプトです。
本スクリプトは、有志による制作であり、公式に配布しているものではありません。
「データ」キーの打刻データにて保存された独自フォーマットCSVを、編集可能なExcelデータに変換します。

# Calcolo100 CSVフォーマット仕様

| No | フィールド名    | 説明                  | フォーマット | 例        |
|----|----------------|----------------------|-------------|-----------|
| 1  | カード No.     | 従業員ID/カード番号     | 数値4桁     | 0001      |
| 2  | 区分           | 記録種別              | 数値1桁     | 1         |
| 3  | 年月日         | 日付                  | YY/MM/DD   | 13/05/16  |
| 4  | 入1時刻        | 1回目の出勤時刻        | HH:MM      | 09:00     |
| 5  | 異例コード     | 1回目の出勤例外コード   | 数値2桁     | 00        |
| 6  | 退1時刻        | 1回目の退勤時刻        | HH:MM      | 18:00     |
| 7  | 異例コード     | 1回目の退勤例外コード   | 数値2桁     | 00        |
| 8  | 入2時刻        | 2回目の出勤時刻        | HH:MM      | 18:30     |
| 9  | 異例コード     | 2回目の出勤例外コード   | 数値2桁     | 04        |
| 10 | 退2時刻        | 2回目の退勤時刻        | HH:MM      | 23:00     |
| 11 | 異例コード     | 2回目の退勤例外コード   | 数値2桁     | 09        |
| 12 | 時数1          | 1回目の勤務時間        | HHH:MM     | 009:00    |
| 13 | 時数2          | 2回目の勤務時間        | HHH:MM     | 006:30    |

## 異例コード一覧

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

## データ例

```csv
0001,1,13/05/16,09:00,00,18:00,00,18:30,04,23:00,09,009:00,006:30,
```

このデータを分解すると：

| フィールド | 値      | 説明                                    |
|-----------|---------|----------------------------------------|
| 1         | 0001    | 従業員番号/カード番号                    |
| 2         | 1       | 区分                                    |
| 3         | 13/05/16| 2013年5月16日                          |
| 4         | 09:00   | 1回目の出勤時刻                         |
| 5         | 00      | 通常出勤（平日）                        |
| 6         | 18:00   | 1回目の退勤時刻                         |
| 7         | 00      | 通常退勤（平日）                        |
| 8         | 18:30   | 2回目の出勤時刻                         |
| 9         | 04      | 再入（休憩後の再入室）                   |
| 10        | 23:00   | 2回目の退勤時刻                         |
| 11        | 09      | 深夜勤務                               |
| 12        | 009:00  | 1回目の勤務時間（9時間）                |
| 13        | 006:30  | 2回目の勤務時間（6時間30分）            |