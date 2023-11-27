@echo off

echo 今月分の日報を自動で作成します。
echo このファイルはショートカットを作成して他の場所に設置可能です。
echo .
echo .
echo .

:input0
set /p arg0=（必須）苗字を入力し、Enterを押してください。 : 
if "%arg0%"=="" (
    echo ※入力してください。
    goto input0
)

:input1
set /p arg1=（必須）名前を入力し、Enterを押してください。 : 
if "%arg1%"=="" (
    echo ※入力してください。
    goto input1
)

:input2
set /p arg2=（オプション）保存先ディレクトリのパスを入力してください。例「C:\Users\user_name\Desktop\日報」※入力を省略した場合は「日報作成ツール」内に作成されます。 : 
if "%arg2%"=="" (
    set arg2=%~dp0
)

powershell.exe -ExecutionPolicy Bypass -File "create_daily_reports.ps1" %arg0% %arg1% %arg2%
