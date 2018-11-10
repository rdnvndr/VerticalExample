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

if not isNull(vobject) then
   ' Получаем атрибут namedse и выводим на экран
   set m_attribute = vobject.vrAttrByName("namedse")
   if not isNull(m_attribute) then
       MsgBox(m_attribute.vrValue)
   end if
end if

' Чистим за собой память
set m_attribute = nothing
set vobject = nothing       
set m_iterator = nothing   
set vmodel = nothing
set m_uniref = nothing