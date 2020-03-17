# 代码
## 新建TXT文件，复制代码，保存，重命名xxxx.bat
## 也可下载office.bat,运行。
## 使用管理员身份运行

@echo off<br>
(cd /d "%~dp0")&&(NET FILE||(powershell start-process -FilePath '%0' -verb runas)&&(exit /B)) >NUL 2>&1<br>
title Office 2019 Activator r/Piracy<br>
echo Converting... & mode 40,25<br>
(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)<br>
cscript //nologo ospp.vbs /unpkey:6MWKP >nul&cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul&set i=1<br>
:server<br>
if %i%==1 set KMS_Sev=kms7.MSGuides.com<br>
if %i%==2 set KMS_Sev=kms8.MSGuides.com<br>
if %i%==3 set KMS_Sev=kms9.MSGuides.com<br>
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul<br>
echo %KMS_Sev% & echo Activating...<br>
cscript //nologo ospp.vbs /act | find /i "successful" && (echo Complete) || (echo Trying another KMS Server & set /a i+=1 & goto server)<br>
pause >nul<br>
exit<br>
