@echo off

echo �������̓���������ō쐬���܂��B
echo ���̃t�@�C���̓V���[�g�J�b�g���쐬���đ��̏ꏊ�ɐݒu�\�ł��B
echo .
echo .
echo .

:input0
set /p arg0=�i�K�{�j�c������͂��AEnter�������Ă��������B : 
if "%arg0%"=="" (
    echo �����͂��Ă��������B
    goto input0
)

:input1
set /p arg1=�i�K�{�j���O����͂��AEnter�������Ă��������B : 
if "%arg1%"=="" (
    echo �����͂��Ă��������B
    goto input1
)

:input2
set /p arg2=�i�I�v�V�����j�ۑ���f�B���N�g���̃p�X����͂��Ă��������B��uC:\Users\user_name\Desktop\����v�����͂��ȗ������ꍇ�́u����쐬�c�[���v���ɍ쐬����܂��B : 
if "%arg2%"=="" (
    set arg2=%~dp0
)

powershell.exe -ExecutionPolicy Bypass -File "create_daily_reports.ps1" %arg0% %arg1% %arg2%
