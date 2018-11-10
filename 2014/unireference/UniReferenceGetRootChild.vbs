' Пример работы с УТС 

' Авторизация
SET UniRef = CreateObject("UniReference.UniRefer")
if not UniRef.GlobalVars.Logon.LogonAsParams("Администратор","111","Администраторы") then
  MsgBox("Авторизация не произведена")
end if

' Получение корневого класса
Set ProfRef = UniRef.BOListGroup.GroupByName("Другие").RootClassByName("PROF")

If Not IsNull(ProfRef) Then
   ' Получение класса LIST
   Set ProfListRef = ProfRef.ClassByName("LIST")
end if   

' Получение списка объектов
Set RootObjects = UniRef.BOListObject()

' Получить корневой объект для PROF.LIST 
Set ProfObject = RootObjects.GetRootObject("PROF.LIST")

' Получить коллекцию объектов детей для корневого объекта  
if ProfObject.GetChildCollectionObject(false,false,false) then 
   ' Получить первый подчиненый объект 
   Set GroupObject = ProfObject.ChildObjects(0)  
   ' Вывод значение первого атрибута                                                      
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