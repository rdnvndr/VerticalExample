strFile = "structureV5.vtp"

' ����������� ����� �������
set manager = CreateObject("Ascon.Integration.AuthenticationManager")
call manager.Authenticate()

' �������� ������
set vModel = CreateObject("vkernel.VModel")

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

if not vModel.vrLoadModel(strFile, nothing, OPEN_LOCAL_STR_ONLY) then 
   MsgBox("���������� ������� ���� ������")
   quit
end if

' ��������� ����� �������
vModel.vrApplySecurity()

' ������� �������� ��������
vModel.vrGetClassVector.vrLocate("root").vrnClassValueItem("after_load").vrFunctionCode = ""

' �������� ������� ����������   
Set FSO = CreateObject("Scripting.FileSystemObject")
Set File = FSO.GetFile("update.vbs")
Set TextStream = File.OpenAsTextStream(1)
vModel.vrGetClassVector.vrPatchFunction("vrPatch")=TextStream.ReadAll()
TextStream.Close

' ��������� ������
call vModel.vrResetOpenMode(OPEN_SERVER_STR_ONLY + WRITE_UNICODE)
call vModel.vrSaveModel("structure.vtp", nothing)
MsgBox "���� ��������"

' ��������� ������ �����
DATAVER     = 0 ' �������� ������ ������ (��������) ���������� �����������.
METADATAVER = 1 ' �������� ������ ���������� (������� � ������������), ����������� 
                ' � ���������� �����������.
SAVEVER     = 2 ' �������� ������, � ������� ��������� ��������� ����������.
FULLVER     = 3 ' �������� ������ ����� �����������.
' MsgBox vModel.vrFileVersion(FULLVER)

call manager.Deauthenticate()
