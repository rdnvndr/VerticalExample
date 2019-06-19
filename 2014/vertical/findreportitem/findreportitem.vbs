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


const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fso, f

' Авторизуемся в универсальном справочнике
SET m_uniref = CreateObject("UniReference.UniRefer")
if not m_uniref.GlobalVars.Logon.LogonAsParams("Администратор","111","Администраторы") then
  MsgBox("Авторизация не произведена")
end if

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("param_log.txt", ForWriting, True)

set rServerTemplates = CreateObject("v2Dobj.RServerTemplates")

r_obj_line_count        = 0
r_obj_rect_count        = 0
r_obj_text_block_count  = 0
r_obj_image_count       = 0
r_obj_table_count       = 0
r_obj_table_cell_count  = 0
r_obj_cell_count        = 0
r_obj_group_count       = 0
r_obj_subst_count       = 0
r_obj_class_count       = 0
r_obj_attr_count        = 0
r_obj_complexattr_count = 0

for i = 0 to rServerTemplates.rCount-1
   set document = rServerTemplates.rItemDoc(i)
   for pageNum = 0 to document.rCount-1
      set page = document.rPage(pageNum)
      obj_line_count        = 0      
      obj_rect_count        = 0
      obj_text_block_count  = 0
      obj_image_count       = 0
      obj_table_count       = 0
      obj_table_cell_count  = 0
      obj_cell_count        = 0
      obj_group_count       = 0
      obj_subst_count       = 0
      obj_class_count       = 0
      obj_attr_count        = 0
      obj_complexattr_count = 0
      for itemNum = 0 to page.rItemCount-1
         set item = page.rItem(itemNum)
         select case item.rType
            case OBJ_LINE
                obj_line_count = obj_line_count + 1
            case OBJ_RECT
                obj_rect_count = obj_rect_count + 1  
            case OBJ_TEXTBLOCK
                obj_text_block_count = obj_text_block_count + 1
            case OBJ_IMAGE
                obj_image_count = obj_image_count + 1 
            case OBJ_TABLE
               obj_table_count = obj_table_count + 1
               set table = item.rAsTable
               if obj_table_cell_count < table.rRows * table.rColumns then
                  obj_table_cell_count = table.rRows * table.rColumns
               end if
            case OBJ_CELL
                obj_cell_count = obj_cell_count +1  
            case OBJ_GROUP
               obj_group_count = obj_group_count + 1             
            case OBJ_SUBST
               obj_subst_count = obj_subst_count + 1
               set subst = item.rAsSubst
               if obj_class_count < subst.rCount then 
                  obj_class_count = subst.rCount
               end if              
               for classNum = 0 to subst.rCount-1
                  set classItem = subst.rClass(classNum)
                  if obj_attr_count < classItem.rCount then
                      obj_attr_count = classItem.rCount
                  end if
                  for attrNum = 0 to classItem.rCount-1
                     set attr = classItem.rAttribute(attrNum)
                     set complexAttr = attr.rAsComplexAttr()
                     if obj_complexattr_count < complexAttr.rcpxAttrCount then
                         obj_complexattr_count = complexAttr.rcpxAttrCount
                     end if
                  next
               next
         end select
      next
       if r_obj_line_count        < obj_line_count then
          r_obj_line_count        = obj_line_count
      end if
      if r_obj_rect_count        < obj_rect_count then
         r_obj_rect_count        = obj_rect_count
      end if
      if r_obj_text_block_count  < obj_text_block_count then
         r_obj_text_block_count  = obj_text_block_count
      end if
      if r_obj_image_count       < obj_image_count then
         r_obj_image_count       = obj_image_count
      end if
      if r_obj_table_count       < obj_table_count then
         r_obj_table_count       = obj_table_count
      end if
      if r_obj_table_cell_count  < obj_table_cell_count then
         r_obj_table_cell_count  = obj_table_cell_count
      end if
      if r_obj_cell_count        < obj_cell_count then
         r_obj_cell_count        = obj_cell_count
      end if
      if r_obj_group_count       < obj_group_count then
         r_obj_group_count       = obj_group_count
      end if
      if r_obj_subst_count       < obj_subst_count then
         r_obj_subst_count       = obj_subst_count
      end if
      if r_obj_class_count       < obj_class_count then
         r_obj_class_count       = obj_class_count
      end if
      if r_obj_attr_count        < obj_attr_count then
         r_obj_attr_count        = obj_attr_count
      end if
      if r_obj_complexattr_count < obj_complexattr_count then
         r_obj_complexattr_count = obj_complexattr_count
      end if     
   next
next
f.Write "    " & "obj_line_count = "        & r_obj_line_count        & vbCr
f.Write "    " & "obj_rect_count = "        & r_obj_rect_count        & vbCr
f.Write "    " & "obj_text_block_count = "  & r_obj_text_block_count  & vbCr
f.Write "    " & "obj_image_count = "       & r_obj_image_count       & vbCr
f.Write "    " & "obj_table_count = "       & r_obj_table_count       & vbCr
f.Write "    " & "obj_table_cell_count = "  & r_obj_table_cell_count  & vbCr
f.Write "    " & "obj_cell_count = "        & r_obj_cell_count        & vbCr
f.Write "    " & "obj_group_count = "       & r_obj_group_count       & vbCr
f.Write "    " & "obj_subst_count = "       & r_obj_subst_count       & vbCr
f.Write "    " & "obj_class_count = "       & r_obj_class_count       & vbCr
f.Write "    " & "obj_attr_count = "        & r_obj_attr_count        & vbCr
f.Write "    " & "obj_complexattr_count = " & r_obj_complexattr_count & vbCr

f.Close
MsgBox "Работа завершена."

set subst = nothing
set item = nothing
set page = nothing
set document = nothing
set rServerTemplates = nothing
set regEx = nothing  
set m_uniref = nothing