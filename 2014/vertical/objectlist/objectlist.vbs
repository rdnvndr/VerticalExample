' ������������ � ������������� �����������
SET m_uniref = CreateObject("UniReference.UniRefer")
if not m_uniref.GlobalVars.Logon.LogonAsParams("�������������","111","��������������") then
  MsgBox("����������� �� �����������")
end if

' �������� ������
set vmodel = CreateObject("vkernel.VModel")
if not vmodel.vrLoadModel("test.vtp",nothing,1) then
  MsgBox("���������� ������� ����")
end if

' �������� ������ �� root ��� ������� DSE
set m_iterator = vmodel.vrGetObjVector

dim i, str
for i=0 to m_iterator.vrObjectsCount-1 
    set vobject = m_iterator.vrItem(i)
    ' MsgBox(vobject.vrClass.vrName)
    str = str & vobject.vrClass.vrName & vbCrLf
next

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fso, f
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("class.txt", ForWriting, True)
f.Write str
f.Close

' ������ �� ����� ������
set vobject = nothing       
set m_iterator = nothing   
set vmodel = nothing
set m_uniref = nothing