Dim MeasureUnit 
MeasureUnit = "VDF643282616547778305DB3EDA1BE6A2" ' Длина, дециметры 
set MeasurementServer = nothing
set MeasurementServer = CreateObject("MeasurementServer.CoMeasurementServer")
if not(MeasurementServer is nothing) then
  set MUnit = MeasurementServer.MUnit(MeasureUnit)
  if not(MUnit is nothing) then
      ' Метры 
      MsgBox(CStr(MUnit.ConvertTo(10,"VA5801DBC0F8F4E6FADEB252E9BEC29A1")))
  end if
end if