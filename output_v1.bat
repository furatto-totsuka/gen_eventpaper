@echo off
SET BASEPATH=%OneDrive%\�L��`�[��\�C�x���g�J�����_�[
python .\src\main.py "%BASEPATH%\%1���ӂ���ƃC�x���g�\.xlsx" -e "%BASEPATH%\�C�x���g�ڍ׈ꗗ�\.xlsx" -t doc> .\out.html 
python .\src\main.py "%BASEPATH%\%1���ӂ���ƃC�x���g�\.xlsx" -e "%BASEPATH%\�C�x���g�ڍ׈ꗗ�\.xlsx" -t googlecsv> .\out.csv 
