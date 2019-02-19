' Авторизуемся в универсальном справочнике
SET m_uniref = CreateObject("UniReference.UniRefer")
if not m_uniref.GlobalVars.Logon.LogonAsParams("Технолог","111","Технологи") then
  MsgBox("Авторизация не произведена")
end if

data = InputBox("Введите строку для шифрования: ")
pass = InputBox("Введите пароль для шифрования: ")
encryptText = m_uniref.GlobalVars.Encrypt(data, pass) 
decryptText = m_uniref.GlobalVars.Decrypt(encryptText, pass)
MsgBox "Исходаный текст: "  & data & vbCr & vbCr _ 
     & "Зашифрованный  текст: "  & encryptText & vbCr _ 
     & "Расшифрованный текст: " & decryptText

set m_uniref = nothing