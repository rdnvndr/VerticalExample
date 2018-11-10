sub Execute 
  ''доступ к ¬ертикали можно получить через глобальную переменную VApplication
  set vmdi = VApplication.ActiveMDIChild
  if not (vmdi is nothing)  then 
     set vModel = vmdi.Content.Model
     call vModel.vrSaveModelVersion(vModel.vrFileName, nothing, 26) 
     set vModel = nothing
  else
     MsgBox "ќтсутствуют открытые технологические процессы"
  end if
  set vmdi = nothing
end sub