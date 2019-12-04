const OBJ_NONE       = 0
const OBJ_LINE       = 1
const OBJ_RECT       = 2
const OBJ_TABLE      = 3
const OBJ_TEXTBLOCK  = 4
const OBJ_IMAGE      = 5
const OBJ_CELL       = 6
const OBJ_SUBST      = 7
const OBJ_ANNOT      = 8
const OBJ_GROUP      = 9
const OBJ_MULTIANNOT = 10

const FindingString = "material"

const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fso, f

' Проверка в имени класса 
function checkInClassName(name)
   checkInClassName = ""
   if Instr(name, FindingString) > 0 then
      checkInClassName = name
   end if
end function

' Авторизация через Полином
set manager = CreateObject("Ascon.Integration.AuthenticationManager")
call manager.Authenticate()

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("find_log.txt", ForWriting, True)

set rServerTemplates = CreateObject("v2Dobj.RServerTemplates")
for i = 0 to rServerTemplates.rCount-1
   set document = rServerTemplates.rItemDoc(i)
   f.Write vbCr &  document.rDocumentCaption & " (" & document.rDocumentGOST & _ 
                " форма " & document.rDocumentForm & ")" & vbCr &_
           "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCr
   for pageNum = 0 to document.rCount-1
      set page = document.rPage(pageNum)
      for itemNum = 0 to page.rItemCount-1
         set item = page.rItem(itemNum)
         select case item.rType
            case OBJ_RECT
               msg = checkInClassName(item.rAttribute.rClass)
               if not msg = "" then
                   f.Write vbCr & "Используется в классе прямоугольника:" & vbCr & msg & vbCr
               end if
            case OBJ_TEXTBLOCK
               msg = checkInClassName(item.rAttribute.rClass)
               if not msg = "" then
                   f.Write vbCr & "Используется в классе текстового блока:" & vbCr & msg & vbCr
               end if
            case OBJ_IMAGE
               msg = checkInClassName(item.rAttribute.rClass)
               if not msg = "" then
                   f.Write vbCr & "Используется в классе изображения:" & vbCr & msg & vbCr
               end if
            case OBJ_TABLE
               set table = item.rAsTable
               for row=0 to table.rRows-1 
                  for col=0 to table.rColumns-1
                     set cell = table.rCell(row, coll)
                     set attr = cell.rText.rParent.rAttribute
                     msg = checkInClassName(attr.rClass)
                     if not msg = "" then
                        f.Write vbCr & "Используется в классе ячейки таблицы:" & vbCr & msg & vbCr
                     end if
                  next
               next
            case OBJ_CELL
               msg = checkInClassName(item.rAttribute.rClass)
               if not msg = "" then
                   f.Write vbCr & "Используется в классе ячейки таблицы:" & vbCr & msg & vbCr
               end if
            case OBJ_GROUP
               msg = checkInClassName(item.rAttribute.rClass)
               if not msg = "" then
                   f.Write vbCr & "Используется в классе ячейки группы:" & vbCr & msg & vbCr
               end if
            case OBJ_SUBST
               set subst = item.rAsSubst
               for classNum = 0 to subst.rCount-1
                  set classItem = subst.rClass(classNum)
                  msg = checkInClassName(classItem.rName)
                  if not msg = "" then
                      f.Write vbCr & "Используется в классе блока подстановки:" & vbCr & msg & vbCr
                  end if
               next
         end select
      next
   next
next

f.Close
call manager.Deauthenticate()
MsgBox "Работа завершена."

set subst = nothing
set item = nothing
set page = nothing
set document = nothing
set rServerTemplates = nothing
set regEx = nothing  
