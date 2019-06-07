set manager = CreateObject("Ascon.Integration.AuthenticationManager")
call manager.Authenticate()

' ������ �������� ����������� ��� ������
UNDEFINED            =    0 ' ����� ����������. �� ������������
OPEN_FULL_STR_SERVER =    1 ' ��������� ��������� (���������) ����, �������� � 
                            ' ���� ��������� ���������� (������ � ������������), 
                            ' ���������� �� �������
OPEN_FULL_STR_LOCAL  =    2 ' ��������� �� ���������� ����� � ���������� � ������
OPEN_SERVER_STR_ONLY =    4 ' ��������� ����������, ���������� �� �������
OPEN_LOCAL_STR_ONLY  =    8 ' ��������� ����������, ���������� ��������
FILES_EXTRACT        =   16 ' ��� ������ ����� ����������� ����������� �������������� 
                            ' � ���� ����� �������� � �.�. ����������� ��������� 
                            ' � OPEN_FULL_STR_SERVER ��� OPEN_FULL_STR_LOCAL.
STR_SERVER_CHECKIN   =   32 ' ������� � V3 �� ������������
COMPACT_METADATA     =   64 ' ��� ������ ����� ����������� ������� ������, 
                            ' �������������� ��������� ������ ������.
WRITE_UNICODE        =  128 ' ���� ���� ����������, �� ��� ���������� ����� ��� 
                            ' ������ ����� �������� � ��������� UNICODE, ����� � 
                            ' � ���, ������� ������������ � ����� ������� 
                            ' ��-��������� (WIN1251).
OPEN_READONLY        =  256 ' ������� �� ������ ��� ������
SAVE_MERGED          =  512 ' ��� ���������� �� ��������� ������� ���������� � 
                            ' ���� ��������� ��. ��������� ���������� �� � ������ 
                            ' �� ��������
SAVE_UNTOUCHED       = 1024 ' ��������� �� �� ������� ��������� ������ � ������, 
                            ' �.�. ������� ����������� ����� �� 

strFile = "test.vtp"
' �������� ������
set vModel = CreateObject("vkernel.VModel")
if not vModel.vrLoadModel(strFile, nothing, OPEN_FULL_STR_SERVER) then
   MsgBox("���������� ������� ����")
   quit
end if

' ��������� ����� �������, ����� ������ �� ������ �� ������
vModel.vrApplySecurity()

' ��������� ����������
Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True   
objRegEx.IgnoreCase = False
objRegEx.Pattern = "\.vtp$"
outfile = objRegEx.Replace(strFile, "_v5.vtp")
if outfile = strFile then
   objRegEx.Pattern = "\.ttp$"
   outfile = objRegEx.Replace(strFile, "_v5.ttp")
end if
call vModel.vrSaveModelVersion(outfile, nothing, 28)
MsgBox "���� ��������:" & vbCR & "    " & outfile

' ��������� ������ �����
DATAVER     = 0 ' �������� ������ ������ (��������) ���������� �����������.
METADATAVER = 1 ' �������� ������ ���������� (������� � ������������), ����������� 
                ' � ���������� �����������.
SAVEVER     = 2 ' �������� ������, � ������� ��������� ��������� ����������.
FULLVER     = 3 ' �������� ������ ����� �����������.
' MsgBox vModel.vrFileVersion(FULLVER)

call manager.Deauthenticate()