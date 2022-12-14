@echo off
setlocal enabledelayedexpansion

cd /d %~dp0

del *.csv

xcopy /S /D:12/9/22 X:\mold_info\mold_yokohama\nissei\MONDAT C:\Users\sbytm\OneDrive\Documents\GitHub\pythonLineChart\data

rem ファイルはShift-JISが前提なので、文字コードがUTF-8の場合は設定
rem chcp 65001

 
rem ヘッダーを判定するための行カウント
set /a cnt=0
 
rem ファイル名を設定
set FILE_NAME=matome.csv
 
rem 移動したフォルダにあるCSVファイルを１つずつ処理
for /f %%a in ('dir /b *.csv') do (
    set TARGET_FILE=%%a
    
    rem ファイル名を出力
    echo TARGET_FILE=!TARGET_FILE!
 
    rem 1ファイル目であればヘッダーを出力
    if !cnt!==0 (
        rem 1行目を取得して出力
        set /p header=<!TARGET_FILE!
        echo !header!>>"!FILE_NAME!"
    )
    
    rem カウントを追加
    set /a cnt=!cnt!+1
    
    rem ファイルの1行目を飛ばして1行ずつ出力
    for /f "usebackq delims= skip=1" %%b in ("!TARGET_FILE!") do (
       echo %%b>>"!FILE_NAME!"
    )
)
 
endlocal