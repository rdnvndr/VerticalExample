' ������������ � ������������� �����������
SET m_uniref = CreateObject("UniReference.UniRefer")
if not m_uniref.GlobalVars.Logon.LogonAsParams("�������������","111","��������������") then
  MsgBox("����������� �� �����������")
end if

' ���������� � ���
SET m_techref = CreateObject("TechReference.TechRefer")
strLocation = CSTR("Material")

if not m_techref.Select(true,(strLocation),m_techref.AppHandle) then
   MsgBox("����� �� ����������")
else
   m_techref.GetObjectInfo2 (sLocation),ClassID, ObjectID
   MsgBox(CSTR(ClassID))   
end if

'MsgBox((strLocation))

set m_techref = nothing
set m_uniref = nothing