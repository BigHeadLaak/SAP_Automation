Option Explicit

SAPInitialize
SapExecute
vbsMacro
LogoutFromSAP
CloseSAPGUI
saveinNetwork
WScript.Quit


'open & Login on SAP, with Arguments of connection,Client,Language,UserName,Password

Sub SAPInitialize

'Use PowerShell to open SAP

Dim WshShell
Dim SAPLocation 
Dim ConnectionConfig 
Dim clientString
Dim Username 
Dim Password 
Dim languageString 
Dim SapGui
Dim Appl
Dim Connection
Dim session
Dim Args



'Change Settings Here

ConnectionConfig = "MIP"
clientString = "100"
Username = 
Password = 
languageString = "DE"


' Try to get SAP GUI object, if an error occurs start up SAP GUI (..., 4)
Set WshShell = CreateObject("WScript.Shell")
	WshShell.run """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"""

' Focus SAP GUI and send it to background
Do Until WshShell.AppActivate("SAP Logon ")
    WScript.Sleep 1000
Loop
    
WshShell.SendKeys "% n"

'Connect to MIP

Set SapGui = GetObject("SAPGUI")
Set Appl = SapGui.GetScriptingEngine
Set Connection = Appl.Openconnection(ConnectionConfig, _
    True)
Set session = Connection.Children(0)

        
'Log on with Information

If IsObject(WScript) Then
	WScript.ConnectObject session, "on"
    WScript.ConnectObject appl, "on"
End If


session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = clientString
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = Username
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Password
session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = languageString
            
session.findById("wnd[0]").sendVKey 0 'ENTER


Set Connection = appl.Children(0)

If Not IsObject(session) Then
    Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject appl, "on"
End If

session.findById("wnd[0]").iconify
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0

End Sub





Sub SapExecute()

'Decalre SAP Variables
	Dim SapGuiAuto
	Dim appl
	Dim Connection
	Dim session
	Dim objExcel
	Dim objWorkbook


'Declare Variables for easier Maintanence

Dim CompanyCode 
Dim ProductLower 
Dim ProductHigher 
Dim Language
Dim ResultFormat 
Dim SaveLocation 
Dim FileName 


'Change Below for different input

CompanyCode = ""
ProductLower = ""
ProductHigher = ""
Language = "DE"
ResultFormat = "/"
SaveLocation = "U:\"
FileName = ".xls"


    If Not IsObject(appl) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set appl = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = appl.Children(0)
    End If
    If Not IsObject(session) Then
       Set session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject appl, "on"
    End If




'SAP Executions

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = ""
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtWERK-LOW").Text = CompanyCode
session.findById("wnd[0]/usr/ctxtSP$00001-LOW").Text = ProductLower
session.findById("wnd[0]/usr/ctxtSP$00001-HIGH").Text = ProductHigher
session.findById("wnd[0]/usr/ctxtSP$00009-LOW").Text = Language
session.findById("wnd[0]/usr/ctxt%LAYOUT").Text = ResultFormat
session.findById("wnd[0]/usr/ctxt%LAYOUT").SetFocus
session.findById("wnd[0]/usr/ctxt%LAYOUT").caretPosition = 6
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "" 
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&PC"
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").SetFocus
session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]/usr/ctxtDY_PATH").SetFocus
session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 0
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[3]/usr/ctxtDY_PATH").Text = SaveLocation
session.findById("wnd[3]/usr/ctxtDY_FILENAME").Text = FileName
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press




End Sub

Sub LogoutFromSAP()
    
	Dim SapGuiAuto
	Dim appl
	Dim Connection
	Dim session
	
    ' Get SAP GUI object and application
    Set SapGuiAuto = GetObject("SAPGUI")
    Set appl = SapGuiAuto.GetScriptingEngine
    
    ' Get current session
    If Not IsObject(Connection) Then
       Set Connection = appl.Children(0)
    End If
    'If IsObject(WScript) Then
    '   WScript.ConnectObject session, "on"
    '   WScript.ConnectObject appl, "on"
    'End If
    If Not IsObject(session) Then
       Set session = Connection.Children(0)
    End If
    
    ' Bring session to background and logout via "/nex"
    session.findById("wnd[0]").iconify
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nex"
    session.findById("wnd[0]").sendVKey 0

End Sub

' Close SAP GUI

Sub CloseSAPGUI()

    Dim sKillSAP
	Dim WshShell

    ' Set command for shell execution
    sKillSAP = "TASKKILL /F /IM saplogon.exe"
    ' Execute klling the saplogon.exe
	Set WshShell = CreateObject("WScript.Shell")
	WshShell.run sKillSAP
        
    ' Do it twice, as sometimes the first try does not work
    sKillSAP = "TASKKILL /F /IM saplgpad.exe"
    WshShell.run sKillSAP
	
	'Shell sKillSAP, vbHide
    
End Sub



Sub vbsMacro ()


Dim objExcel
Dim objWorkbook
Dim FileLocation
Dim FileName
Dim NewFileName



FileLocation = ""
FileName = ".xls"
NewFileName = ".xlsx"
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(FileLocation+FileName)
objExcel.Visible = True
objExcel.DisplayAlerts = False

'delete Empty Rows & Columns

objWorkbook.Worksheets("").Rows(1).EntireRow.Delete
objWorkbook.Worksheets("").Rows(1).EntireRow.Delete
objWorkbook.Worksheets("").Rows(1).EntireRow.Delete
objWorkbook.Worksheets("").Rows(1).EntireRow.Delete
objWorkbook.Worksheets("").Rows(1).EntireRow.Delete


objWorkbook.Worksheets("").Range("A:A").EntireColumn.Delete
objWorkbook.Worksheets("").Range("B:B").EntireColumn.Delete
objWorkbook.Worksheets("").Range("C:D").EntireColumn.Delete
objWorkbook.Worksheets("").Range("I:I").EntireColumn.Delete

'Number the rows
Dim indexNu
indexNu = 1
For indexNu = 1 to 9
objWorkbook.Worksheets("").Cells(1,indexNu).Value = indexNu
next 

objWorkbook.Worksheets("").Rows(3).EntireRow.Delete

'Save Xlsx File
Const xlExclusive = 3
Const xlLocalSessionChanges = 2

objWorkbook.SaveAs FileLocation+NewFileName,51, , , , , xlExclusive, xlLocalSessionChanges
objExcel.DisplayAlerts = True
objWorkbook.Close 


end Sub 



sub saveinNetwork()


Dim fileExcel
Dim objPath
Dim fileLocation
Dim newLocation1
Dim newLocation2


fileLocation = 
newLocation1 = 
newLocation2 = 

set fileExcel=CreateObject("Scripting.FileSystemObject")
Set objPath = CreateObject("Scripting.FileSystemObject")

If fileExcel.FileExists(fileLocation) Then

	If objPath.FolderExists(newLocation1) Then
		fileExcel.CopyFile fileLocation, newLocation1,True
	Else wscript.echo "This Path does not exist:" & newLocation1
	End if

	if objPath.FolderExists(newLocation2) Then 
		fileExcel.CopyFile fileLocation, newLocation2,True
	Else wscript.echo "This Path does not exists:" & newLocation2
	End if


else wscript.echo "The excel file does not exist"
End If

end sub