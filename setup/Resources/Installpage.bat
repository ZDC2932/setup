@echo off
::�����װ������
if exist "C:\Program Files (x86)\Zhuozhengsoft\PageOfficeClient\POBrowser.exe" goto cktray
if exist "C:\Program Files\Zhuozhengsoft\PageOfficeClient\POBrowser.exe" goto cktray
::����������ֱ�Ӱ�װ
if exist "C:%homepath%\desktop\posetup.msi" goto il1
if exist posetup.msi goto install01
powershell (new-object System.Net.WebClient).DownloadFile('https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/posetup.msi','posetup.msi')
:install01
msiexec /i "posetup.msi" /quiet
@sc config PageService start= auto>NUL 2>&1
@ping -n 1 127.1>nul
@sc start PageService>NUL 2>&1
@ping -n 2 127.1>nul
goto cktray
:il1
msiexec /i "C:%homepath%\desktop\posetup.msi" /quiet
@sc config PageService start= auto>NUL 2>&1
@ping -n 1 127.1>nul
@sc start PageService>NUL 2>&1
@ping -n 2 127.1>nul
:cktray
@sc config PageService start= auto>NUL 2>&1
@ping -n 1 127.1>nul
@sc start PageService>NUL 2>&1
@ping -n 2 127.1>nul
echo **�Ѿ����в����װ����** >>report.txt
::��������ƻ�
schtasks /query /tn pengshen-server
if errorlevel == 1 goto cresc
goto endsc
:cresc
schtasks /create /sc ONSTART /tn pengshen-server /tr "poc.bat" /RL HIGHEST
:endsc
echo **�Ѿ�ע������װ����** >>report.txt
::����Ƿ�������δ����ֱ��������
if exist "C:\Program Files\Zhuozhengsoft\PageOfficeClient\POCService.exe" (goto setin1)
set processStr="C:\Program Files (x86)\Zhuozhengsoft\PageOfficeClient\pageofficetray.exe"
goto n2
:setin1
set processStr="C:\Program Files\Zhuozhengsoft\PageOfficeClient\POCService.exe"
:n2
tasklist | find "pageofficetray.exe" > "%homepath%\Documents\start.log"
if errorlevel == 1 start "" %processStr%
echo �����װ�Ϳ���������ִ�гɹ� >>report.txt