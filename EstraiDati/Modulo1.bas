Attribute VB_Name = "Modulo1"
Option Explicit

Type RecordsType
       HeaderID As Long
       Flux_ID As Long
       ErrQ_ID As Long
       REVF_ID As Long
       REVR_ID As Long
       EMEWS_R_ID As Long
       EMEWS_F_ID As Long
       Data As Date
       DataRev As Date
       OrigF As Double
       OrigR As Double
       RevF As Double
       RevR As Double
       RL As Single
       LL As Single
       Ns As Single
       FS As Single
       dt As Single
       AcK As Double
End Type


Public Db As ADODB.Connection
Public Rec() As RecordsType
Public StationName As String
Public StationFile As String
Public FileLog As String
Public Stations(50) As String
Public nStations As Integer
Public PathName As String




Function Open_ADODB_Connection() As Boolean
Dim e As Long
Dim strVersionInfo As String
Dim fn As Long

On Error GoTo ADODB_TRAP
Open_ADODB_Connection = False
Set Db = CreateObject("ADODB.Connection")
Db.Open "DSN=WEST"

strVersionInfo = "ADO Version: " & Db.Version & vbCr & _
   "DBMS Name: " & Db.Properties("DBMS Name") & vbCr & _
   "DBMS Version: " & Db.Properties("DBMS Version") & vbCr & _
   "OLE DB Version: " & Db.Properties("OLE DB Version") & vbCr & _
   "Provider Name: " & Db.Properties("Provider Name") & vbCr & _
   "Provider Version: " & Db.Properties("Provider Version") & vbCr & _
   "Driver Name: " & Db.Properties("Driver Name") & vbCr & _
   "Driver Version: " & Db.Properties("Driver Version") & vbCr & _
   "Driver ODBC Version: " & Db.Properties("Driver ODBC Version")

'MsgBox (strVersionInfo)



Open_ADODB_Connection = True

On Error GoTo 0
Exit Function

ADODB_TRAP:
    e = Err
    'MsgBox (Error(e) + " +" + Str(e))
    On Error GoTo 0
    fn = FreeFile
    Open FileLog For Append As #fn
    Print #fn, "-------------------------------"
    Print #fn, "Open_ADODB_Connection"
    Print #fn, Now
    Print #fn, Err.Number; " "; Err.Description; " "; Err.Source
'    Print #fn, "StartDate="; fmain_click.
'    Print #fn, "StopDate="; StopDate
    Print #fn, "Station Name="; StationName
'    Print #fn, "records="; Rc
    Print #fn, "Station file="; StationFile

    Close fn
    End

    End

End Function

Function sGetAppPath() As String
'*Returns the application path with a trailing \.      *
'*To use, call the function [SomeString=sGetAppPath()] *
Dim sTemp As String
        sTemp = App.Path
        If Right$(sTemp, 1) <> "\" Then sTemp = sTemp + "\"
        sGetAppPath = sTemp
End Function

Public Function GetStationsFromFile() As Boolean
    Dim filename As String
    Dim fn As Long
    
    GetStationsFromFile = False

    
    nStations = 1
    filename = sGetAppPath + "Stations.txt"
    
    fn = FreeFile
    On Error GoTo ErrorTrap
    Open filename For Input As #fn
    On Error GoTo 0
    
    Do Until EOF(fn)
        Input #fn, Stations(nStations)
        nStations = nStations + 1
    Loop
    Close fn
    nStations = nStations - 1
    If nStations > 0 Then GetStationsFromFile = True
    
    Exit Function
ErrorTrap:
'On Error GoTo 0
fn = FreeFile
Open FileLog For Append As #fn
Print #fn, "-------------------------------"
Print #fn, Now
Print #fn, "GetStationsFromFile"
Print #fn, Err.Number; " "; Err.Description; " "; Err.Source
Close fn
End
End Function
