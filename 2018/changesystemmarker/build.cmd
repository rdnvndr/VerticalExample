@echo off

rem ��������� ���� ���������� ��������� �� �������� ���������

rem ������ ��������� ���������
set libs=Ascon.Vertical.Technology.dll

rem ������������ ���������
set out=changesystemmarker

rem ������ �����������
set csc_ver=v4.0



for /f "delims=" %%a in ('where /R %windir%\Microsoft.NET\assembly\GAC_MSIL\Ascon.Integration Ascon.Integration.dll') do set "refs=%%a"
for /f "delims=" %%a in ('where /R %windir%\Microsoft.NET\Framework64\ csc.exe ^| findstr %csc_ver%') do set "CSC=%%a"

set VRT=%ProgramFiles%\ASCON\Vertical
IF EXIST "%VRT%" GOTO EXISTDIR
  set VRT=%ProgramFiles(x86)%\ASCON\Vertical

:EXISTDIR

for %%i in (%libs%) do (
set refs=%refs%;"%VRT%\%%i"
)

%CSC% /out:%out%.exe /R:%refs% *.cs



