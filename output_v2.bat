@echo off
SET DETAILFILE=%OneDrive%\広報チーム\イベントカレンダー\イベント詳細一覧表.xlsx
echo HTMLを出力しています...
python .\src\main.py "%1" -e "%DETAILFILE%" -t doc -o out.html
echo Googleカレンダー用CSVファイルを出力しています...
python .\src\main.py "%1" -e "%DETAILFILE%" -t googlecsv -o googlecsv.csv
echo 処理が完了しました