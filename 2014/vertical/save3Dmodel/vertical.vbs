' Авторизуемся в универсальном справочнике
SET m_uniref = CreateObject("UniReference.UniRefer")
if not m_uniref.GlobalVars.Logon.LogonAsParams("Администратор","111","Администраторы") then
  MsgBox("Авторизация не произведена")
end if

' Получаем модель
set vmodel = CreateObject("vkernel.VModel")
if not vmodel.vrLoadModel("test.vtp",nothing,1) then
  MsgBox("Невозможно открыть фаил")
end if

' Получаем список из root для классов DSE
set m_iterator = vmodel.vrGetObjVector.vrCreateIterator("dse",vmodel.vrGetObjVector.vrItem(0),true)

' Получаем первый объект DSE
m_iterator.vrFirst
set vobject = m_iterator.vrGetObject

set m_iterator = vmodel.vrGetObjVector.vrCreateIterator("3d_models",vobject,true)
m_iterator.vrFirst
set vobject = m_iterator.vrGetObject


dim i
if not isNull(vobject) then
for i=1 to vobject.vrAttrCount 
    MsgBox(vobject.vrAttrByIndex(i-1).vrName)
next
   set m_attribute = vobject.vrAttrByName("kompasfile")
   if not isNull(m_attribute) then
'       MsgBox(m_attribute.vrValue)

      Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
  
     'Create text stream object
     Dim TextStream
     Set TextStream = FS.CreateTextFile("kompasfile.t3d")

     MsgBox(m_attribute.vrFile.vsDiskFullName)
     ' MsgBox(m_attribute.vrValue)
  
     'Convert binary data To text And write them To the file
     ' TextStream.Write BinaryToString(m_attribute.vrValue)
       
     
   
   end if
end if

' Чистим за собой память
set m_attribute = nothing
set vobject = nothing       
set m_iterator = nothing   
set vmodel = nothing
set m_uniref = nothing