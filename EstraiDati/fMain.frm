VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "EstraiDati"
   ClientHeight    =   1680
   ClientLeft      =   4125
   ClientTop       =   3285
   ClientWidth     =   2760
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   2760
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "fMain.frx":0C42
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   180
      Width           =   495
   End
   Begin VB.CommandButton bTest 
      Caption         =   "&NewDir"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton bEnd 
      Caption         =   "&End"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1260
      Width           =   615
   End
   Begin VB.CommandButton bVai 
      Caption         =   "&Vai"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   660
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "INGV-PA"
      Height          =   615
      Left            =   180
      TabIndex        =   3
      Top             =   840
      Width           =   1275
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    Dim ret As Boolean
    
    INIFile = sGetAppPath + "EstraiDati.ini"
    FileLog = sGetAppPath + "EstraiDati.txt"

    ret = GetStationsFromFile
    If ret = False Then
    'allarme
    End If
    PathName = sReadINI("Setup", "SavePath", INIFile)
    Debug.Print PathName
    Label1.Caption = "INGV-PA" + vbCrLf + "Roberto Maugeri" + vbCrLf + "2002 v.1.0"
    DoEvents
    bVai.Enabled = False
    DoEvents
    bTest.Enabled = False
    DoEvents
    
End Sub

Private Sub Form_Paint()
    'bTest_Click
    bVai_Click
    DoEvents
    End
End Sub

Private Sub bEnd_Click()
    Me.Hide
    Unload Me
    End
End Sub

Private Sub bTest_Click()
    Dim pathh As String
    Dim ret As Boolean
    Dim lungo As Long
    pathh = InputBox("Immetti la cartella ove salvare i files", "EstraiDati")
    If pathh = "" Then Exit Sub
    'ret = GetStationsFromFile
    'PathName = sReadINI("Setup", "SavePath", INIFile)
    lungo = WriteINI("Setup", "SavePath", pathh, INIFile)
End Sub

Private Sub bVai_Click()
    Dim ret As Boolean
    Dim SQL As String
    Dim rs As Object
    Dim i As Long
    Dim j As Long
    Dim Rc As Long
    Dim AcK As Double
    Dim StartYear As Integer
    Dim StopYear As Integer
    Dim StartMonth As Integer
    Dim StopMonth As Integer
    Dim StartDay As Integer
    Dim StopDay As Integer
    Dim StartDate As String
    Dim StopDate As String
    Dim fn As Long
    Dim fn1 As Long
    Dim fn2 As Long
    
    
    On Error GoTo ErrorTrap
    
    DoEvents
    
    ret = Open_ADODB_Connection
    
    DoEvents
    
    If ret = False Then
        Exit Sub
    End If
    
    
    'Definizione data di stop ricerca
    StopYear = Year(Now)
    StopMonth = Month(Now)
    StopDay = Day(Now + 1)
    'Definizione data di start ricerca
    StartDay = 1
    If StopMonth = 1 Then
        StartYear = StopYear - 1
        StartMonth = 12
    Else
        StartYear = StopYear
        StartMonth = StopMonth - 1
    End If
       
    'creazione startdate e stopdate
    StopDate = Trim(Str(StopYear)) + "/" + Format(StopMonth, "00") + "/" + Format(StopDay, "00")
    StartDate = Trim(Str(StartYear)) + "/" + Format(StartMonth, "00") + "/" + Format(StartDay, "00")
    
    'aggiunta dei cancelletti #
    StartDate = "#" + StartDate + "#"
    StopDate = "#" + StopDate + "#"
    
    DoEvents
    
    'prova
'    StartDate = "#2000/11/01#"
'    StopDate = "#2000/12/31#"
    
    'Per tutte le stazioni
    For j = 1 To nStations
        StationName = "ETN01"
        StationName = Stations(j)
        
        DoEvents
        
        'crea la connessione
        Set Db = CreateObject("ADODB.Connection")
        DoEvents
        Db.Open "DSN=WEST"
        DoEvents
        Set rs = CreateObject("ADODB.Recordset")
        DoEvents
        rs.CursorType = adOpenStatic
        DoEvents
        'Prende la costante per passare da ppm/s a grammi/m2/giorno
        SQL = "SELECT STAZIONI.AccumulationChamberK FROM STAZIONI WHERE ((STAZIONI.Station_ID) ='" + StationName + "'); "
        rs.Open SQL, Db
        DoEvents
        AcK = rs("AccumulationChamberK")
        DoEvents
        rs.Close
        DoEvents
        Db.Close

        'crea la connessione
        Set Db = CreateObject("ADODB.Connection")
        DoEvents
        Db.Open "DSN=WEST"
        DoEvents
        Set rs = CreateObject("ADODB.Recordset")
        DoEvents
        rs.CursorType = adOpenStatic
        DoEvents
        
        SQL = "SELECT HEADERS.DATA_SAMP, HEADERS.DATA_REVISIONE, "
        SQL = SQL + "CHANNELS.Value as REVR,  canali.EMEWS_CH_ID, "
        SQL = SQL + "canali.Revision_FLAG, HEADERS.ID, "
        SQL = SQL + "canali_1.Revision_FLAG, CHANNELS_1.Value "
        SQL = SQL + "as REVF, HEADERS.STATION_ID FROM canali "
        SQL = SQL + "AS canali_1 INNER JOIN (CHANNELS AS "
        SQL = SQL + "CHANNELS_1 INNER JOIN ((HEADERS INNER JOIN "
        SQL = SQL + "CHANNELS ON HEADERS.ID = CHANNELS.ID_HEADER) "
        SQL = SQL + "INNER JOIN canali ON CHANNELS.ID_CANALE = "
        SQL = SQL + "canali.EMEWS_CH_ID) ON CHANNELS_1.ID_HEADER "
        SQL = SQL + "= HEADERS.ID) ON canali_1.EMEWS_CH_ID = "
        SQL = SQL + "CHANNELS_1.ID_CANALE  WHERE "
        SQL = SQL + "(((( HEADERS.DATA_SAMP BETWEEN "
        SQL = SQL + StopDate + " AND " + StartDate + " )) AND "
    
    '    SQL = SQL + "#2002/01/01# AND #1997/01/01# )) AND "
        SQL = SQL + "((canali.Revision_FLAG)='REVR') AND "
        SQL = SQL + "((canali_1.Revision_FLAG)='REVF')  AND "
        SQL = SQL + "( HEADERS.STATION_ID)='" + StationName + "')) ORDER BY "
        
    '    SQL = SQL + "( HEADERS.STATION_ID)='ETN01')) ORDER BY "
        SQL = SQL + "HEADERS.DATA_SAMP;"
    
        fn2 = FreeFile
        'Open FileLog For Append As #fn2
        DoEvents
        'Print #fn2, Now
        Me.MousePointer = vbHourglass
        
        DoEvents
        
        rs.Open SQL, Db
    
        DoEvents
        
        Me.MousePointer = vbNormal
        
        'Print #fn2, Now
        'Close fn2
        
        DoEvents
        
        If rs.EOF Then
            'No Data!!!!
            GoTo continua
'           MsgBox "No data macht your query"
'           rs.Close
'           Me.MousePointer = vbNormal
'           Exit Sub
        End If
        DoEvents
        rs.MoveFirst
        DoEvents
    
        
        Rc = rs.RecordCount
        DoEvents
        Debug.Print StationName; " records-->"; Rc
    
        fn1 = FreeFile
        StationFile = PathName + "\" + StationName + ".dat"
        'StationFile = StationName + ".txt"
        On Error GoTo ErrorTrap
        Open StationFile For Output As #fn1
        DoEvents
        On Error GoTo 0
        Print #fn1, StationFile
        Print #fn1, "3"
        Print #fn1, "date"
        Print #1, "time"
        Print #1, StationName + "_CO2_flux_grams/m2/d"
        DoEvents
    
      While Not rs.EOF
        i = i + 1
        'FileLog = rs("DATA_SAMP") & "," & rs("REVF")
        Print #fn1, Format(rs("DATA_SAMP"), "dd/mm/yyyy hh:mm:ss") & " " & Format(rs("REVF") * AcK, "0.0000")
        'Print #fn1, FileLog
    '    Rec(i).Data = rs("DATA_SAMP")
    '    If IsNull(rs("DATA_REVISIONE")) Then
    '      Rec(i).DataRev = -1
    '    Else
    '      Rec(i).DataRev = rs("DATA_REVISIONE")
    '    End If
    '
    '    Rec(i).HeaderID = rs("ID")
    '    Rec(i).OrigF = rs("FLUX")
    '    Rec(i).RevF = rs("REVF")
    '    Rec(i).OrigR = rs("ERRQ")
    '    Rec(i).RevR = rs("REVR")
    '    Rec(i).AcK = AcK
        DoEvents
        rs.MoveNext
        DoEvents
      Wend
    
    Close fn1
    

continua:
    rs.Close

Next j

Exit Sub

ErrorTrap:
  On Error GoTo 0
fn = FreeFile
Open FileLog For Append As #fn
Print #fn, "-------------------------------"
Print #fn, Now
Print #fn, "bVai_Click"
Print #fn, Err.Number; " "; Err.Description; " "; Err.Source
Print #fn, "StartDate="; StartDate
Print #fn, "StopDate="; StopDate
Print #fn, "Station Name="; StationName
Print #fn, "records="; Rc
Print #fn, "Station file="; StationFile
Close fn
End

End Sub
