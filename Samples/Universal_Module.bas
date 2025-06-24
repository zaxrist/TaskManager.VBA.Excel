Attribute VB_Name = "Universal_Module"
'Module Description: this module is for local use only. Not used in any database connection
'the purpose of this module is to simplify repeatable task
'Project start: 2-march-2023 3:48AM
'Created by: Zackris

'Module Revision:-
'1.0 : Project started

Public Type POINTAPI
X As Long
Y As Long
End Type
Public Declare PtrSafe Function GetCursorPos Lib "User32" (IpPoint As POINTAPI) As Long 'get cursor coordinate

'MD5 string encryption
Public Function GetMD5(ByVal textString As String) As String
'soruce: https://stackoverflow.com/questions/36384741/cant-use-getbytes-and-computehash-methods-on-vba
  Dim enc
  Dim textBytes() As Byte
  Dim bytes
  Dim outstr As String
  Dim pos As Long
  Dim txtMsg As String
  
  If textString = "" Then Exit Function
  
  Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
  textBytes = textString
  bytes = enc.ComputeHash_2((textBytes))
    
  For pos = 1 To LenB(bytes)
    outstr = outstr & lCase(Right("0" & Hex(AscB(MidB(bytes, pos, 1))), 2))
  Next
  GetMD5 = outstr
  Set enc = Nothing
End Function

Public Function rSpace(ByVal stg As String)
rSpace = Replace(stg, " ", "")
End Function

Function lstrow(ByVal sheetName As String, Optional ByVal StartRow As Long = 1) As Long 'count last row on any sheet
lstrow = ThisWorkbook.Sheets(sheetName).Cells(Rows.count, StartRow).End(xlUp).row
End Function
Function lstCol(ByVal sheetName As String, Optional ByVal startCol As Long = 1) As Long 'same as above but for columns
lstCol = ThisWorkbook.Sheets(sheetName).Cells(startCol, Columns.count).End(xlToLeft).Column
End Function

'2.0 Settings function_____________________________________________________________________
Public Function GetSetting(ByVal RowNumber As String) As Variant
ThisWorkbook.Activate
With ThisWorkbook.Worksheets("Settings")
    GetSetting = .Cells(RowNumber, 2).Value
End With
End Function
Public Function SetSetting(ByVal RowNumber As Integer, ByVal settingValue As String) As Boolean
Application.ScreenUpdating = False
On Error GoTo fail
ThisWorkbook.Activate
With ThisWorkbook.Worksheets("Settings")
    .Cells(RowNumber, 2).Value = settingValue
End With
SetSetting = True
Application.ScreenUpdating = True
Exit Function
fail:
SetSetting = False
Application.ScreenUpdating = True
End Function

Public Function CheckApos(ByVal theString As String) As String 'check if string is all numeric or not and retun with apostrophe for access database
If isNumeric(theString) = True Then
   CheckApos = theString
Else
    CheckApos = "'" & theString & "'"
End If
End Function

Public Function getDbSetting(ByVal id As Integer) As String
Dim ds As New dbaseC
ds.Table = "dbAdmin"
getDbSetting = ds.GetRecordset("db_ID", id, True).item(2)
End Function
Public Function setDbSetting(ByVal id As Integer, ByVal Value As String) As Boolean
Dim ds As New dbaseC
ds.Table = "dbAdmin"
ds.addColnVal "db_Value", Value
ds.UpdateRecord "db_ID", id, True
End Function

Public Function CheckForUpdates() As Integer
Dim thisVersion() As String
Dim dbWVersion() As String
Dim dbaseThis() As String
Dim dbaseServer() As String
Dim needUpdate As Boolean
Dim UpdateLvl As Integer
UpdateLvl = 0

thisVersion = Split(GetSetting(6), ".")
dbWVersion = Split(getDbSetting(1), ".")
dbaseThis = Split(GetSetting(8), ".")
dbaseServer = Split(getDbSetting(2), ".")

If CInt(thisVersion(0)) >= CInt(dbWVersion(0)) Then
    txtMsg = "Major workbook version up to date..."
Else
    txtMsg = "Major workbook version REQUIRES update"
    UpdateLvl = 3
    GoTo lst
End If
If CInt(thisVersion(1)) >= CInt(dbWVersion(1)) Then
    txtMsg = "Minor workbook version up to date..."
Else
    txtMsg = "Minor workbook version REQUIRES update"
    UpdateLvl = 2
    GoTo lst
End If
If CInt(thisVersion(2)) >= CInt(dbWVersion(2)) Then
    txtMsg = "Beta workbook version up to date..."
Else
    txtMsg = "Beta workbook version REQUIRES update"
    UpdateLvl = 1
    GoTo lst
End If

If CInt(dbaseThis(0)) >= CInt(dbaseServer(0)) Then
    txtMsg = "Major Database is up to date.."
Else
    txtMsg = "Major Database REQUIRES update"
    UpdateLvl = 3
    GoTo lst
End If
If CInt(dbaseThis(1)) >= CInt(dbaseServer(1)) Then
    txtMsg = "Minor Database is up to date.."
Else
    txtMsg = "Minor Database REQUIRES update"
    UpdateLvl = 2
    GoTo lst
End If
If CInt(dbaseThis(2)) >= CInt(dbaseServer(2)) Then
    txtMsg = "Beta Database is up to date.."
Else
    txtMsg = "Beta Database REQUIRES update"
    UpdateLvl = 1
    GoTo lst
End If

lst:
CheckForUpdates = UpdateLvl
If getDevMode = True Then
Debug.Print UpdateLvl & "--" & txtMsg
End If
End Function

Public Function askForUpdates(ByVal UpdateLvl As Integer) As Boolean
Dim ans As Integer
Dim fso As New FileSystemObject
Dim LatestVersion As String
Dim deploymentServer As String

If CheckForUpdates >= UpdateLvl Then
 ans = MsgBox("New software version is released and is required to UPDATE." & vbNewLine & "Continue?" & vbNewLine & vbNewLine & "New version is located in the same folder with current version", vbYesNo + vbApplicationModal + vbInformation)
 If ans = 6 Then
 LatestVersion = getDbSetting(1)
 deploymentServer = "\\172.23.192.20\myshare\Moduleqa\P8_EQ_DataBase\Deployment\" & LatestVersion
    If Not fso.FolderExists(deploymentServer) Then
        MsgBox "Update folder not exist. Contact your developer"
        askForUpdates = False
        Exit Function
    End If
    fso.CopyFile deploymentServer & "\CMMS v" & LatestVersion & ".xlsb", ThisWorkbook.path & "\CMMS v" & LatestVersion & ".xlsb", True
    MsgBox "New version has been downloaded into this software current directory. Application will shut down and OPEN the LATEST version of CMMS" & vbNewLine & vbNewLine & _
    "CMMS Latest version: v" & LatestVersion, vbInformation
    Application.DisplayAlerts = False
    ThisWorkbook.Saved = True
    On Error Resume Next
    Application.Quit
    End
 Else
    End
 End If
askForUpdates = True
Else
    askForUpdates = False
End If
End Function

Public Function deleteOldVersion()
On Error Resume Next
Dim file As Variant
file = Dir(ThisWorkbook.path & "\*") 'scan all files in the folder and delete old one
Do While Len(file) > 0
    If InStr(1, file, "CMMS", vbBinaryCompare) <> 0 Then
        If InStr(Mid(file, 7, 5), GetSetting(6)) = 0 Then
           If InStr(Mid(file, 6, 7), "vDeploy") <> 1 Then
            Kill (ThisWorkbook.path & "\" & file)
           End If
        End If
    End If
    file = Dir
Loop
End Function

Function GenerateID() As String
    GenerateID = Format(Now(), "ddMMyyHHmmss") & Mid(Rnd(), 4, 2)
End Function
Function NowTime() As String
NowTime = Format(Now(), "dd/MMM/yy HH:mm")
End Function
Public Function GetDocFolder() As String
    GetDocFolder = Environ$("USERPROFILE") & "\Documents"
End Function

Public Function RemoveBlankLineString(ByVal stg As String) As String
On Error GoTo lst
Dim oldStg As String
oldStg = stg
Dim RE As Object: Set RE = CreateObject("VBScript.RegExp")
With RE
    .MultiLine = True
    .Global = True
    .Pattern = "(\r\n)+"
   stg = .Replace(stg, vbCrLf)
End With
RemoveBlankLineString = stg
Exit Function
lst:
RemoveBlankLineString = oldStg
End Function

Public Function getDevMode() As Boolean
If StrComp(GetSetting(15), "FALSE", vbTextCompare) = 0 Then
    getDevMode = False
Else
    getDevMode = True
End If
End Function
Public Function SetDevModeOff()
SetSetting 15, "FALSE"
End Function


