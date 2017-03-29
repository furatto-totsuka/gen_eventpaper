@echo off
SET BASEPATH=%OneDrive%\広報チーム\イベントカレンダー
python .\src\main.py "%BASEPATH%\%1月ふらっとイベント表.xlsx" -e "%BASEPATH%\イベント詳細一覧表.xlsx" -t doc> .\out.html 
python .\src\main.py "%BASEPATH%\%1月ふらっとイベント表.xlsx" -e "%BASEPATH%\イベント詳細一覧表.xlsx" -t googlecsv> .\out.csv 
