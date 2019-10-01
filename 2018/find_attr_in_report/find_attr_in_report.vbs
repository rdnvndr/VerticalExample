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

const FindingString = "CreateChildMaterial(operation)"

const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fso, f

' Проверка в скриптах документа
function checkInDocument(document)
   checkInDocument = ""
   if Instr(document.rFunctionAfter, FindingString) > 0 then
      checkInDocument = checkInDocument & vbCr & document.rFunctionAfter
   end if
   if Instr(document.rFunctionBefore, FindingString) > 0 then
      checkInDocument = checkInDocument & vbCr & document.rFunctionBefore
   end if
end function

' Проверка в скриптах страницы
function checkInPage(page)
   checkInPage = "" 
   if Instr(page.rFunctionAfter, FindingString) > 0 then
       checkInPage = page.rFunctionAfter
   end if
   if Instr(page.rFunctionBefore, FindingString) > 0 then
       checkInPage = checkInPage & vbCr & page.rFunctionBefore
   end if
end function

' Проверка в имени атрибута 
function checkInAttributeName(name)
   checkInAttributeName = ""
   if Instr(name, FindingString) > 0 then
      checkInAttributeName = name
   end if
end function

' Проверка в скриптах атрибута 
function checkInAttribute(attr)
   checkInAttribute = ""
   if Instr(attr.rFunction, FindingString) > 0 then
      checkInAttribute = attr.rFunction
   end if
end function

' Проверка в скриптах класса 
function checkInClass(classItem)
   checkInClass = "" 
   if Instr(classItem.rFunctionBefore, FindingString) > 0 then
      checkInClass = classItem.rFunctionBefore
   end if
   if Instr(classItem.rFunctionBeforeChildren, FindingString) > 0 then
      checkInClass = checkInClass & vbCr & classItem.rFunctionBeforeChildren
   end if
   if Instr(classItem.rFunctionParseAutoItems, FindingString) > 0 then
      checkInClass = checkInClass & vbCr & classItem.rFunctionParseAutoItems
   end if
   if Instr(classItem.rFunctionAfter, FindingString) > 0 then
      checkInClass = checkInClass & vbCr & classItem.rFunctionAfter
   end if
end function

' Проверка в скриптах блока подстановок
function checkInSubst(subst)
   checkInSubst = ""
   if Instr(subst.rFunctionStart, FindingString) > 0 then
      checkInSubst = subst.rFunctionStart
   end if
   if Instr(subst.rFunctionEnd , FindingString) > 0 then
      checkInSubst = checkInSubst & vbCr & subst.rFunctionEnd
   end if
end function

' Проверка в скриптах комплексного атрибута
function checkInComplexAttr(comlexAttr)
   checkInComplexAttr = ""
   if Instr(complexAttr.rcpxFunctionBefore, FindingString) > 0 then
      checkInComplexAttr = complexAttr.rcpxFunctionBefore
   end if 
   if Instr(complexAttr.rcpxFunctionBeforeObject, FindingString) > 0 then
      checkInComplexAttr = checkInComplexAttr & vbCr & complexAttr.rcpxFunctionBeforeObject
   end if
end function

function checkInAttrOfComplexAttr(comlexAttr, num)
   checkInAttrOfComplexAttr = ""
   if Instr(complexAttr.rcpxAttrFunction(num), FindingString) > 0 then
      checkInAttrOfComplexAttr = complexAttr.rcpxAttrFunction(num)
   end if 
end function

' Авторизация через Полином
set manager = CreateObject("Ascon.Integration.AuthenticationManager")
call manager.Authenticate()

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("param_log.txt", ForWriting, True)

set rServerTemplates = CreateObject("v2Dobj.RServerTemplates")
for i = 0 to rServerTemplates.rCount-1
   set document = rServerTemplates.rItemDoc(i)
   f.Write vbCr &  document.rDocumentCaption & " (" & document.rDocumentGOST & _ 
                " форма " & document.rDocumentForm & ")" & vbCr &_
           "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCr
   msg = checkInDocument(document)
   if not msg = "" then
      f.Write vbCr & "Используется в функции документа:" & vbCr & msg & vbCr
   end if
   for pageNum = 0 to document.rCount-1
      set page = document.rPage(pageNum)
      msg = checkInPage(page)
      if not msg = "" then
          f.Write vbCr & "Используется  в функции страницы:" & vbCr & msg & vbCr
      end if
      for itemNum = 0 to page.rItemCount-1
         set item = page.rItem(itemNum)
         select case item.rType
            case OBJ_RECT
               msg = checkInAttribute(item.rAttribute)
               if not msg = "" then
                   f.Write vbCr & "Используется в функции прямоугольника:" & vbCr & msg & vbCr
               end if
               msg = checkInAttributeName(item.rAttribute.rClassValue)
               if not msg = "" then
                   f.Write vbCr & "Используется в атрибуте прямоугольника:" & vbCr & msg & vbCr
               end if
            case OBJ_TEXTBLOCK
               msg = checkInAttribute(item.rAttribute)
               if not msg = "" then
                   f.Write vbCr & "Используется в функции текстового блока:" & vbCr & msg & vbCr
               end if
               msg = checkInAttributeName(item.rAttribute.rClassValue)
               if not msg = "" then
                   f.Write vbCr & "Используется в атрибуте текстового блока:" & vbCr & msg & vbCr
               end if
            case OBJ_IMAGE
               msg = checkInAttribute(item.rAttribute)
               if not msg = "" then
                   f.Write vbCr & "Используется в функции изображения:" & vbCr & msg & vbCr
               end if
               msg = checkInAttributeName(item.rAttribute.rClassValue)
               if not msg = "" then
                   f.Write vbCr & "Используется в атрибуте изображения:" & vbCr & msg & vbCr
               end if
            case OBJ_TABLE
               set table = item.rAsTable
               for row=0 to table.rRows-1 
                  for col=0 to table.rColumns-1
                     set cell = table.rCell(row, coll)
                     set attr = cell.rText.rParent.rAttribute
                     msg = msg & checkInAttribute(attr)
                     if not msg = "" then
                         f.Write vbCr & "Используется в функции ячейки таблицы:" & vbCr & msg & vbCr
                     end if 
                     msg = checkInAttributeName(attr.rClassValue)
                     if not msg = "" then
                        f.Write vbCr & "Используется в атрибуте ячейки таблицы:" & vbCr & msg & vbCr
                     end if
                  next
               next
            case OBJ_CELL
               msg = checkInAttribute(item.rAttribute)
               if not msg = "" then
                   f.Write vbCr & "Используется в функции ячейки таблицы:" & vbCr & msg & vbCr
               end if
               msg = checkInAttributeName(item.rAttribute.rClassValue)
               if not msg = "" then
                   f.Write vbCr & "Используется в атрибуте ячейки таблицы:" & vbCr & msg & vbCr
               end if
            case OBJ_GROUP
               msg = checkInAttribute(item.rAttribute)
               if not msg = "" then
                   f.Write vbCr & "Используется в функции группы:" & vbCr & msg & vbCr
               end if
               msg = checkInAttributeName(item.rAttribute.rClassValue)
               if not msg = "" then
                   f.Write vbCr & "Используется в атрибуте ячейки группы:" & vbCr & msg & vbCr
               end if
            case OBJ_SUBST
               set subst = item.rAsSubst
               msg = checkInSubst(subst)
               if not msg = "" then
                   f.Write vbCr & "Используется в функции блока подстановок:" & vbCr & msg & vbCr
               end if
               for classNum = 0 to subst.rCount-1
                  set classItem = subst.rClass(classNum)
                  msg = checkInClass(classItem)
                  if not msg = "" then
                      f.Write vbCr & "Используется в функции класса:" & vbCr & msg & vbCr
                  end if
                  for attrNum = 0 to classItem.rCount-1
                     set attr = classItem.rAttribute(attrNum)
                     msg = checkInAttribute(attr)
                     if not msg="" then
                         f.Write vbCr & "Используется в функции атрибута класса:" & vbCr & msg & vbCr
                     end if
                     msg = checkInAttributeName(attr.rName)
                     if not msg="" then
                         f.Write vbCr & "Используется в атрибуте атрибута класса:" & vbCr & msg & vbCr
                     end if
                     set complexAttr = attr.rAsComplexAttr()
                     msg = checkInComplexAttr(comlexAttr) 
                     if not msg="" then
                        f.Write vbCr & "Используется в функции комплексного атрибута:" & vbCr & msg & vbCr
                     end if 
                     for attrComplexNum = 0 to complexAttr.rcpxAttrCount-1   
                        msg = checkInAttrOfComplexAttr(complexAttr, attrComplexNum) 
                        msg = checkInAttributeName(complexAttr.rcpxAttrName(attrComplexNum)) & msg
                        if not msg = "" then
                           f.Write vbCr & "Используется в атрибуте комплексного атрибута:" & vbCr & msg & vbCr
                        end if
                     next
                  next
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
