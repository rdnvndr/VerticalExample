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
set m_iterator = vmodel.vrGetObjVector.vrCreateIterator("dse",vmodel.vrGetObjVector.vrItem(0),true)

' �������� ������ ������ DSE
m_iterator.vrFirst
set vobject = m_iterator.vrGetObject

if not isNull(vobject) then
   ' �������� ������� namedse � ������� �� �����
   set m_attribute = vobject.vrAttrByName("namedse")
   if not isNull(m_attribute) then
       MsgBox(m_attribute.vrValue)
   end if
end if

' ������ �� ����� ������
set m_attribute = nothing
set vobject = nothing       
set m_iterator = nothing   
set vmodel = nothing
set m_uniref = nothing