Const HKEY_CURRENT_USER = &H80000001

strComputer = "."
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\Microsoft\Internet Explorer\Toolbar\WebBrowser"
 
strValueName = "ITBar7Position"

Result = InputBox("Do you wish to fix the IE7 MenuBar?" & vbcrlf & vbcrlf & "Please enter Yes or No.","Fix IE7 MenuBar")   
If lcase(Result) = "yes" then
    dwValue = 1
elseIf lcase(Result) = "no" then
    dwValue = 0
End if

oReg.SetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue

msgbox "Please restart the IE7 application to view the changes selected.",,"Restart IE7"



