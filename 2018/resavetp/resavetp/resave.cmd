set REPO=..\Structure\Vertical\Program Files\ASCON\Vertical
set PROG=C:\Program Files (x86)\ASCON\Vertical

copy "%REPO%\Template\assembly.vtp" "%PROG%\Template\assembly.vtp" /Y
copy "%REPO%\Template\detail.vtp"   "%PROG%\Template\detail.vtp" /Y
copy "%REPO%\Template\ttp.ttp"      "%PROG%\Template\ttp.ttp" /Y

copy "%REPO%\Samples\������.vtp"       "%PROG%\Samples\������.vtp" /Y
copy "%REPO%\Samples\��� ��室���.vtp" "%PROG%\Samples\��� ��室���.vtp" /Y
copy "%REPO%\Samples\���.001.005 ���� �����.ttp"   "%PROG%\Samples\���.001.005 ���� �����.ttp" /Y
copy "%REPO%\Samples\������\078.505.9.0100.00.vtp" "%PROG%\Samples\������\078.505.9.0100.00.vtp" /Y


resave.exe "%REPO%\RestoreFiles\Model\structure.vtp" ^
           "%PROG%\Template\assembly.vtp" ^
           "%PROG%\Template\detail.vtp" ^
           "%PROG%\Template\ttp.ttp" ^
           "%PROG%\Samples\������.vtp" ^
           "%PROG%\Samples\��� ��室���.vtp" ^
           "%PROG%\Samples\���.001.005 ���� �����.ttp" ^
           "%PROG%\Samples\������\078.505.9.0100.00.vtp"


copy "%PROG%\Template\assembly.vtp" "%REPO%\Template\assembly.vtp" /Y
copy "%PROG%\Template\detail.vtp"   "%REPO%\Template\detail.vtp" /Y
copy "%PROG%\Template\ttp.ttp"      "%REPO%\Template\ttp.ttp" /Y

copy "%PROG%\Samples\������.vtp"       "%REPO%\Samples\������.vtp"  /Y
copy "%PROG%\Samples\��� ��室���.vtp" "%REPO%\Samples\��� ��室���.vtp" /Y
copy "%PROG%\Samples\���.001.005 ���� �����.ttp"   "%REPO%\Samples\���.001.005 ���� �����.ttp" /Y
copy "%PROG%\Samples\������\078.505.9.0100.00.vtp" "%REPO%\Samples\������\078.505.9.0100.00.vtp" /Y
