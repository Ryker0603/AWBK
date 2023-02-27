chcp 65001
@setlocal DisableDelayedExpansion
@echo off
: Get the admin
: %1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit
: Keep running in the current directory
: cd /d "%~dp0"

::========================================================================================================================================
: Basic settings 
: KMS Server
set K1=kms.catqu.com
set K2=kms.myds.cloud
set K3=kms.wxlost.com
set K4=kms.ddns.net
set K5=kms.luody.info

: Windows activation key
: Windows Vista
set WVB=Windows Vista Business
set WVE=Windows Vista Enterprise

: Windows 7
set W7P=Windows 7 Pro
set W7E=Windows 7 Enterprise

: Windows 8.1
set W8.1P=Windows 8.1 Pro
set W8.1E=Windows 8.1 Enterprise

: Windows 10
set W10H=Windows 10 Home
set W10ED=Windows 10 Education 
set W10P=Windows 10 Pro
set W10PW=Windows 10 Pro for Workstations
set W10E=Windows 10 Enterprise

: Windows 11
set W11H=Windows 11 Home
set W11ED=Windows 11 Education 
set W11P=Windows 11 Pro
set W11PW=Windows 11 Pro for Workstations
set W11E=Windows 11 Enterprise

: Save check active status content file name
set TXT=CheckActivationStatus.txt

::========================================================================================================================================
: Language
title  Activate Windows By KMS
mode con cols=98 lines=30
set EN=:
set TC=:
set SC=:
set EXIT=:
set ENCK=:
set TCCK=:
set SCCK=:

echo:
echo:
echo                   _______________________________________________________________
echo                  ^|      Choose Your Language                                     ^| 
echo                  ^|                                                               ^| 
echo                  ^|      [1] English                                              ^|
echo                  ^|                                                               ^|
echo                  ^|      [2] 繁体中文                                             ^|
echo                  ^|                                                               ^|
echo                  ^|      [3] 简体中文                                             ^|
echo                  ^|                                                               ^|
echo                  ^|      [4] Exit                                                 ^|
echo                  ^|_______________________________________________________________^|
echo:          
choice /C:1234 /N /M ">                   Enter Your Choice in the Keyboard [1,2,3...] : "

if errorlevel  4 exit /b
if errorlevel  3 cls & call :SC & goto :MainMenu
if errorlevel  2 cls & call :TC & goto :MainMenu
if errorlevel  1 cls & call :EN & goto :MainMenu 
:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:EN
set EN=echo
set EXIT=Exit
set ENCK=choice
set KSE=KMS server error
exit /b

:TC
set TC=echo
set EXIT=退出
set TCCK=choice
set KSE=KMS 服務器錯誤
exit /b

:SC
set SC=echo
set EXIT=退出
set SCCK=choice
set KSE=KMS 服务器错误
exit /b


::========================================================================================================================================
:MainMenu

cls
title  Activate Windows By KMS
mode con cols=98 lines=30

echo:
echo:
echo                   _______________________________________________________________
echo                  ^|                                                               ^| 
%EN%                  ^|      [1] About                                                ^|
%TC%                  ^|      [1] 關於                                                 ^|
%SC%                  ^|      [1] 关于                                                 ^|
echo                  ^|                                                               ^|
%EN%                  ^|      [2] Activate Windows                                     ^|
%TC%                  ^|      [2] 激活Windows系統                                      ^|
%SC%                  ^|      [2] 激活Windows系统                                      ^|
echo                  ^|                                                               ^|
%EN%                  ^|      [3] Uninstall Key                                        ^|
%TC%                  ^|      [3] 卸載密鑰                                             ^|
%SC%                  ^|      [3] 卸载密钥                                             ^|
echo                  ^|                                                               ^|
%EN%                  ^|      [4] Check Activation Status                              ^|
%TC%                  ^|      [4] 檢查激活狀態                                         ^|
%SC%                  ^|      [4] 查看激活状态                                         ^|
echo                  ^|                                                               ^|
echo                  ^|      [5] %EXIT%                                                 ^|
echo                  ^|_______________________________________________________________^|
echo:          
%ENCK% /C:12345 /N /M ">                   Enter Your Choice in the Keyboard [1,2,3...] : "
%TCCK% /C:12345 /N /M ">                   在鍵盤中輸入您的選擇 [1,2,3...] : "
%SCCK% /C:12345 /N /M ">                   在键盘中输入您的选择 [1,2,3...] : "

if errorlevel  5 exit /b
if errorlevel  4 setlocal & call :CheckStatus & cls & endlocal & cls & goto :MainMenu
if errorlevel  3 setlocal & call :Uninstall Key & cls & endlocal & goto :MainMenu
if errorlevel  2 setlocal & call :Activate Windows & cls & endlocal & goto :MainMenu
if errorlevel  1 setlocal & call :About & cls & endlocal & goto :MainMenu

::========================================================================================================================================

:About
cls

echo ************************************************************
echo ***                 About AWBK Debug                     ***
echo ************************************************************  
echo #This script was developed by Ryker Lim
echo #Github:https://github.com/Ryker0603
echo #This script uses from https://github.com/LulzSecToolkit/KMS-activator-V-X/blob/master/Check-Activation-Status-Alternative.cmd as a script for Check Activation Status Alternative
echo #This script uses KMS to activate Windows Vista above Windows systems.
echo #Version: 1.1
echo #Version number:202302270
echo.
pause 
exit /b
::========================================================================================================================================

:Activate Windows
cls
mode con cols=98 lines=30

echo:
echo:
echo                   _______________________________________________________________
echo                  ^|                                                               ^|
echo                  ^|      [1] Windows Vista                                        ^|
echo                  ^|                                                               ^|
echo                  ^|      [2] Windows 7                                            ^|
echo                  ^|                                                               ^| 
echo                  ^|      [3] Windows 8.1                                          ^|
echo                  ^|                                                               ^|
echo                  ^|      [4] Windows 10                                           ^|
echo                  ^|                                                               ^|
echo                  ^|      [5] Windows 11                                           ^|
echo                  ^|                                                               ^|
echo                  ^|      [6] %EXIT%                                                 ^|
echo                  ^|_______________________________________________________________^|
echo:          
%ENCK% /C:123456 /N /M ">                   Enter Your Choice in the Keyboard [1,2,3...] : "
%TCCK% /C:123456 /N /M ">                   在鍵盤中輸入您的選擇 [1,2,3...] : "
%SCCK% /C:123456 /N /M ">                   在键盘中输入您的选择 [1,2,3...] : "

if errorlevel  6 exit /b
if errorlevel  5 setlocal & call :Win11 & cls & endlocal & goto :MainMenu
if errorlevel  4 setlocal & call :Win10 & cls & endlocal & goto :MainMenu
if errorlevel  3 setlocal & call :Win8 & cls & endlocal & goto :MainMenu
if errorlevel  2 setlocal & call :Win7 & cls & endlocal & goto :MainMenu
if errorlevel  1 setlocal & call :WinV & cls & endlocal & goto :MainMenu
:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
:WinV
cls
mode con cols=98 lines=30

echo:
echo:
echo                   _______________________________________________________________
echo                  ^|                                                               ^|
%EN%                  ^|      [1] Windows Vista Business                               ^|
%TC%                  ^|      [1] Windows Vista 商業版                                 ^|
%SC%                  ^|      [1] Windows Vista 商业版                                 ^|
echo                  ^|                                                               ^|
%EN%                  ^|      [2] Windows Vista Enterprise                             ^|
%TC%                  ^|      [2] Windows Vista 企業版                                 ^|
%SC%                  ^|      [2] Windows Vista 企业版                                 ^|
echo                  ^|                                                               ^|
echo                  ^|      [3] %EXIT%                                                 ^|
echo                  ^|_______________________________________________________________^|
echo:          
%ENCK% /C:123 /N /M ">                   Enter Your Choice in the Keyboard [1,2,3...] : "
%TCCK% /C:123 /N /M ">                   在鍵盤中輸入您的選擇 [1,2,3...] : "
%SCCK% /C:123 /N /M ">                   在键盘中输入您的选择 [1,2,3...] : "

if errorlevel  3 exit /b
if errorlevel  2 cls & echo %WVE% &ping 127.0.0.1 -n 5 >nul & goto :KMS test
if errorlevel  1 cls & echo %WVB%  &ping 127.0.0.1 -n 5 >nul & goto :KMS test
:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
:Win7
cls
mode con cols=98 lines=30

echo:
echo:
echo                   _______________________________________________________________
echo                  ^|                                                               ^|
%EN%                  ^|      [1] Windows 7 Pro                                        ^|
%TC%                  ^|      [1] Windows 7 專業版                                     ^|
%SC%                  ^|      [1] Windows 7 专业版                                     ^|
echo                  ^|                                                               ^|
%EN%                  ^|      [2] Windows 7 Enterprise                                 ^|
%TC%                  ^|      [2] Windows 7 企業版                                     ^|
%SC%                  ^|      [2] Windows 7 企业版                                     ^|
echo                  ^|                                                               ^|
echo                  ^|      [3] %EXIT%                                                 ^|
echo                  ^|_______________________________________________________________^|
echo:          
%ENCK% /C:123 /N /M ">                   Enter Your Choice in the Keyboard [1,2,3...] : "
%TCCK% /C:123 /N /M ">                   在鍵盤中輸入您的選擇 [1,2,3...] : "
%SCCK% /C:123 /N /M ">                   在键盘中输入您的选择 [1,2,3...] : "

if errorlevel  3 exit /b
if errorlevel  2 cls & echo %W7E% &ping 127.0.0.1 -n 5 >nul & goto :KMS test
if errorlevel  1 cls & echo %W7P%  &ping 127.0.0.1 -n 5 >nul & goto :KMS test
:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
:Win8
cls
mode con cols=98 lines=30

echo:
echo:
echo                   _______________________________________________________________
echo                  ^|                                                               ^|
%EN%                  ^|      [1] Windows 8.1 Pro                                      ^|
%TC%                  ^|      [1] Windows 8.1 專業版                                    ^|
%SC%                  ^|      [1] Windows 8.1 专业版                                    ^|
echo                  ^|                                                               ^|
%EN%                  ^|      [2] Windows 8.1 Enterprise                               ^|
%TC%                  ^|      [2] Windows 8.1 企業版                                    ^|
%SC%                  ^|      [2] Windows 8.1 企业版                                    ^|
echo                  ^|                                                               ^|
echo                  ^|      [3] %EXIT%                                                 ^|
echo                  ^|_______________________________________________________________^|
echo:          
%ENCK% /C:123 /N /M ">                   Enter Your Choice in the Keyboard [1,2,3...] : "
%TCCK% /C:123 /N /M ">                   在鍵盤中輸入您的選擇 [1,2,3...] : "
%SCCK% /C:123 /N /M ">                   在键盘中输入您的选择 [1,2,3...] : "

if errorlevel  3 exit /b
if errorlevel  2 cls & echo %W8.1E% &ping 127.0.0.1 -n 5 >nul & goto :KMS test
if errorlevel  1 cls & echo %W8.1P%  &ping 127.0.0.1 -n 5 >nul & goto :KMS test
:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
:Win10
cls
mode con cols=98 lines=30

echo:
echo:
echo                   _______________________________________________________________
echo                  ^|                                                               ^| 
%EN%                  ^|      [1] Windows 10 Home                                      ^|
%TC%                  ^|      [1] Windows 10 家庭版                                    ^|
%SC%                  ^|      [1] Windows 10 家庭版                                    ^|
echo                  ^|                                                               ^|
%EN%                  ^|      [2] Windows 10 Education                                 ^|
%TC%                  ^|      [2] Windows 10 教育版                                    ^|
%SC%                  ^|      [2] Windows 10 教育版                                    ^|
echo                  ^|                                                               ^|
%EN%                  ^|      [3] Windows 10 Pro                                       ^|
%TC%                  ^|      [3] Windows 10 專業版                                    ^|
%SC%                  ^|      [3] Windows 10 专业版                                    ^|
echo                  ^|                                                               ^|
%EN%                  ^|      [4] Windows 10 Pro for Workstations                      ^|
%TC%                  ^|      [4] Windows 10 工作站專業版                              ^|
%SC%                  ^|      [4] Windows 10 工作站专业版                              ^|
echo                  ^|                                                               ^|
%EN%                  ^|      [5] Windows 10 Enterprise                                ^|
%TC%                  ^|      [5] Windows 10 企業版                                    ^|
%SC%                  ^|      [5] Windows 10 企业版                                    ^|
echo                  ^|                                                               ^|
echo                  ^|      [6] %EXIT%                                                 ^|
echo                  ^|_______________________________________________________________^|
echo:          
%ENCK% /C:123456 /N /M ">                   Enter Your Choice in the Keyboard [1,2,3...] : "
%TCCK% /C:123456 /N /M ">                   在鍵盤中輸入您的選擇 [1,2,3...] : "
%SCCK% /C:123456 /N /M ">                   在键盘中输入您的选择 [1,2,3...] : "
if errorlevel  6 exit /b
if errorlevel  5 cls & echo %W10E% &ping 127.0.0.1 -n 5 >nul & goto :KMS test
if errorlevel  4 cls & echo %W10PW% &ping 127.0.0.1 -n 5 >nul & goto :KMS test
if errorlevel  3 cls & echo %W10P% &ping 127.0.0.1 -n 5 >nul & goto :KMS test
if errorlevel  2 cls & echo %W10ED% &ping 127.0.0.1 -n 5 >nul & goto :KMS test
if errorlevel  1 cls & echo %W10H% &ping 127.0.0.1 -n 5 >nul & goto :KMS test


:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
:Win11
cls
mode con cols=98 lines=30

echo:
echo:
echo                   _______________________________________________________________
echo                  ^|                                                               ^| 
%EN%                  ^|      [1] Windows 11 Home                                      ^|
%TC%                  ^|      [1] Windows 11 家庭版                                    ^|
%SC%                  ^|      [1] Windows 11 家庭版                                    ^|
echo                  ^|                                                               ^|
%EN%                  ^|      [2] Windows 11 Education                                 ^|
%TC%                  ^|      [2] Windows 11 教育版                                    ^|
%SC%                  ^|      [2] Windows 11 教育版                                    ^|
echo                  ^|                                                               ^|
%EN%                  ^|      [3] Windows 11 Pro                                       ^|
%TC%                  ^|      [3] Windows 11 專業版                                    ^|
%SC%                  ^|      [3] Windows 11 专业版                                    ^|
echo                  ^|                                                               ^|
%EN%                  ^|      [4] Windows 11 Pro for Workstations                      ^|
%TC%                  ^|      [4] Windows 11 工作站專業版                              ^|
%SC%                  ^|      [4] Windows 11 工作站专业版                              ^|
echo                  ^|                                                               ^|
%EN%                  ^|      [5] Windows 11 Enterprise                                ^|
%TC%                  ^|      [5] Windows 11 企業版                                    ^|
%SC%                  ^|      [5] Windows 11 企业版                                    ^|
echo                  ^|                                                               ^|
echo                  ^|      [6] %EXIT%                                                 ^|
echo                  ^|_______________________________________________________________^|
echo:          
%ENCK% /C:123456 /N /M ">                   Enter Your Choice in the Keyboard [1,2,3...] : "
%TCCK% /C:123456 /N /M ">                   在鍵盤中輸入您的選擇 [1,2,3...] : "
%SCCK% /C:123456 /N /M ">                   在键盘中输入您的选择 [1,2,3...] : "

if errorlevel  6 exit /b
if errorlevel  5 cls & echo %W11E% &ping 127.0.0.1 -n 5 >nul & goto :KMS test
if errorlevel  4 cls & echo %W11PW% &ping 127.0.0.1 -n 5 >nul & goto :KMS test
if errorlevel  3 cls & echo %W11P% &ping 127.0.0.1 -n 5 >nul & goto :KMS test
if errorlevel  2 cls & echo %W11ED% &ping 127.0.0.1 -n 5 >nul & goto :KMS test
if errorlevel  1 cls & echo %W11H% &ping 127.0.0.1 -n 5 >nul & goto :KMS test

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
:KMS test


set KMS=%K1%

ping %KMS%
echo %KMS%
if %errorlevel%==0 (echo slmgr.vbs /skms %KMS% & goto :ato) else (set KMS=%K2%  )
ping %KMS%
echo %KMS%
if %errorlevel%==0 (echo slmgr.vbs /skms %KMS% & goto :ato) else (set KMS=%K3% )
ping %KMS%
echo %KMS%
if %errorlevel%==0 (echo slmgr.vbs /skms %KMS% & goto :ato) else (set KMS=%K4% )
ping %KMS%
echo %KMS%
if %errorlevel%==0 (echo slmgr.vbs /skms %KMS% & goto :ato) else (set KMS=%K5% )
ping %KMS%
echo %KMS%
if %errorlevel%==0 (echo slmgr.vbs /skms %KMS% & goto :ato) else (echo %KSE% & goto :end)

:ato
echo slmgr.vbs /ato
ping 127.0.0.1 -n 5 >nul  
echo slmgr.vbs /xpr
%EN% Windows activation process complete
%TC% Windows 激活過程完成
%SC% Windows 激活过程完成

:end
pause 
exit 

::========================================================================================================================================
:Uninstall Key
cls

echo slmgr /upk
%EN% The key has been uninstalled
%TC% 密鑰已卸載
%SC% 密钥已卸载
echo slmgr /cpky
%EN% The registry information of the key has been removed
%TC% 該密鑰的註冊表信息已被刪除
%SC% 该密钥的注册表信息已被删除

pause 
exit /b 

::========================================================================================================================================

:CheckStatus
cls
title Check Activation Status
:: ==================================================
::  Check-Activation-Status-Alternative.cmd
::  https://github.com/LulzSecToolkit/KMS-activator-V-X/blob/master/Check-Activation-Status-Alternative.cmd
:: ==================================================

echo ************************************************************  
%EN% ***                   Windows Status                     *** 
%TC% ***                   Windows 狀態                       ***
%SC% ***                   Windows 状态                       ***  
echo ************************************************************  
copy /y %systemroot%\System32\slmgr.vbs "%temp%\slmgr.vbs" >nul 2>&1 
cscript //nologo "%temp%\slmgr.vbs" /dli  
cscript //nologo "%temp%\slmgr.vbs" /xpr  
echo ____________________________________________________________________________  
echo. 
ping 127.0.0.1 -n 5 >nul

set verb=0
set spp=SoftwareLicensingProduct
for /f "tokens=2 delims==" %%G in ('"wmic path %spp% where (PartialProductKey is not null) get ID /value"') do (set app=%%G&call :chk)
del /f /q "%temp%\slmgr.vbs" >nul 2>&1
echo.
%ENCK% /C:yn /N /M "> Do you need to save information? [Y/N] : "
%TCCK% /C:yn /N /M "> 您需要保存信息嗎？ [Y/N] : "
%SCCK% /C:yn /N /M "> 您需要保存信息吗？ [Y/N] : "
if errorlevel 2 goto :end2
if errorlevel 1 goto :CASO

:chk
wmic path %spp% where ID='%app%' get Name /value | findstr /i "Windows" 1>nul && (exit /b)
if %verb%==0 (
set verb=1
echo ************************************************************
%EN% ***                   Office Status                      ***
%TC% ***                   Office 狀態                        ***
%SC% ***                   Office 状态                        ***  
echo ************************************************************
)
cscript //nologo "%temp%\slmgr.vbs" /dli %app%
cscript //nologo "%temp%\slmgr.vbs" /xpr %app%
echo ____________________________________________________________________________
ping 127.0.0.1 -n 5 >nul
echo.
exit /b
 
:CASO
%EN% Start export %TXT%
%TC% 開始導出 %TXT%
%SC% 开始导出 %TXT%

echo ************************************************************ >%TXT% 
%EN% ***                        Windows Status                              *** >>%TXT% 
%TC% ***                        Windows 狀態                                *** >>%TXT%
%SC% ***                        Windows 状态                                *** >>%TXT%
echo ************************************************************ >>%TXT%
copy /y %systemroot%\System32\slmgr.vbs "%temp%\slmgr.vbs" >nul 2>&1  
cscript //nologo "%temp%\slmgr.vbs" /dli  >>%TXT%
cscript //nologo "%temp%\slmgr.vbs" /xpr  >>%TXT%
echo ____________________________________________________________________________  >>%TXT%
echo. >>%TXT%

set verb=0
set spp=SoftwareLicensingProduct
for /f "tokens=2 delims==" %%G in ('"wmic path %spp% where (PartialProductKey is not null) get ID /value"') do (set appt=%%G&call :chkt)
del /f /q "%temp%\slmgr.vbs" >nul 2>&1
echo. >>%TXT%
%EN% Already exported %TXT% & goto :end2
%TC% 已經導出 %TXT% & goto :end2
%SC% 已经导出 %TXT% & goto :end2

:chkt
wmic path %spp% where ID='%appt%' get Name /value | findstr /i "Windows" 1>nul && (exit /b)
if %verb%==0 (
set verb=1
echo ************************************************************ >>%TXT%
%EN% ***                           Office Status                                 *** >>%TXT%
%TC% ***                           Office 狀態                                   *** >>%TXT%
%SC% ***                           Office 状态                                   *** >>%TXT%
echo ************************************************************ >>%TXT%
)
cscript //nologo "%temp%\slmgr.vbs" /dli %appt% >>%TXT%
cscript //nologo "%temp%\slmgr.vbs" /xpr %appt% >>%TXT%
echo ____________________________________________________________________________ >>%TXT%
echo. >>%TXT%
exit /b

:end2
pause 
exit /b

::========================================================================================================================================



