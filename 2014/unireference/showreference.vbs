
' ������ ������ � ��� 

' �����������
SET UniRef = CreateObject("UniReference.UniRefer")
if not UniRef.GlobalVars.Logon.LogonAsParams("�������������","111","��������������") then
  MsgBox("����������� �� �����������")
end if

' ��������� ��������� ������
Set ProfRef = UniRef.BOListGroup.GroupByName("������").RootClassByName("PROF")

If Not IsNull(ProfRef) Then
   ' ��������� ������ LIST
   Set ProfListRef = ProfRef.ClassByName("LIST")
end if   

' ��������� ������ ��������
Set RootObjects = UniRef.BOListObject()

' �������� �������� ������ ��� PROF.LIST 
Set ProfObject = RootObjects.GetRootObject("PROF.LIST")

' �������� ��������� �������� ����� ��� ��������� �������  
if ProfObject.GetChildCollectionObject(false,false,false) then 
   ' �������� ������ ���������� ������ 
   Set GroupObject = ProfObject.ChildObjects(0)  
   ' ����� �������� ������� ��������                                                      
    Set Attr = GroupObject.Attributes.AttrObjectByIndex(0)  
    MsgBox(Attr.NameAttr + "=" + Attr.DataAttr)
   
   if GroupObject.GetChildCollectionObject(false,false,false) then        
      Set ListObject = GroupObject.ChildObjects(0) 
      Set Attr = ListObject.Attributes.AttrObjectByIndex(0)  
      MsgBox(Attr.NameAttr + "=" + Attr.DataAttr)
   end if
end if        

set Attr = nothing
set ListObject = nothing
set GroupObject = nothing
set ProfObject = nothing
set RootObjects = nothing

set ProfListRef = nothing
set ProfRef = nothing
set UniRef = nothing