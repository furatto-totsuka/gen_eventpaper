@echo off
SET DETAILFILE=%OneDrive%\�L��`�[��\�C�x���g�J�����_�[\�C�x���g�ڍ׈ꗗ�\.xlsx
echo HTML���o�͂��Ă��܂�...
python .\src\main.py "%1" -e "%DETAILFILE%" -t doc -o .\out\out.html
echo Google�J�����_�[�pCSV�t�@�C�����o�͂��Ă��܂�...
python .\src\main.py "%1" -e "%DETAILFILE%" -t googlecsv -o .\out\googlecsv.csv
echo �������������܂���