VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Form1"
   ClientHeight    =   1680
   ClientLeft      =   4125
   ClientTop       =   3285
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   2760
   Begin VB.CommandButton bEnd 
      Caption         =   "&End"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton bVai 
      Caption         =   "&Vai"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bEnd_Click()
    Me.Hide
    Unload Me
    End
End Sub

Private Sub bVai_Click()
    Dim Ret As Boolean
    Dim SQL As String
    Dim rs As Object
    Dim i As Long
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
    Dim fn1 As Long
    Dim fn2 As Long
    
    FileLog = sGetAppPath + "EstraiDati.txt"
    
    
    Ret = Open_ADODB_Connection
    If Ret = False Then
    End If
    
    'Exit Sub
    
    Set Db = CreateObject("ADODB.Connection")
    Db.Open "DSN=WEST"

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorType = adOpenStatic
    
    'Definizione data di stop ricerca
    StopYear = Year(Now)
    StopMonth = Month(Now)
    StopDay = Day(Now)
    'Definizione data di start ricerca
    StartDay = 1
    If StopMonth = 1 Then
        StartYear = StopYear - 1
        StartMonth = 12
    Else
        StartYear = StopYear
        StartMonth = StopMonth - 1
    End If
       
    
    StopDate = Trim(Str(StopYear)) + "/" + Format(StopMonth, "00") + "/" + Format(StopDay, "00")
    StartDate = Trim(Str(StartYear)) + "/" + Format(StartMonth, "00") + "/" + Format(StartDay, "00")
    
    'aggiunta dei cancelletti #
    StartDate = "#" + StartDate + "#"
    StopDate = "#" + StopDate + "#"
    
    
    'prova
'    StartDate = "#2000/11/01#"
'    StopDate = "#2000/12/31#"
    
    
    StationName = "ETN01"
    
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
    SQL = SQL + "( HEADERS.STATION_ID)='ETN01')) ORDER BY "
    SQL = SQL + "HEADERS.DATA_SAMP;"

    fn2 = FreeFile
    Open FileLog For Append As #fn2
    Print #fn2, Now
    Me.MousePointer = vbHourglass
    
    DoEvents
    
    rs.Open SQL, Db

    DoEvents
    
    Me.MousePointer = vbNormal
    
    Print #fn2, Now
    Close fn2
    
    If rs.EOF Then
       MsgBox "No data macht your query"
       rs.Close
       Me.MousePointer = vbNormal
       Exit Sub
    End If
    rs.MoveFirst

    
    Rc = rs.RecordCount
    Debug.Print Rc

    fn1 = FreeFile
    StationFile = StationName + ".txt"
    Open StationFile For Output As #fn1
    Print #1, StationFile
    Print #1, "3"

  While Not rs.EOF
    i = i + 1
    'FileLog = rs("DATA_SAMP") & "," & rs("REVF")
    Print #fn1, Format(rs("DATA_SAMP"), "yyyy/mm/dd hh:mm:ss") & " " & rs("REVF")
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
    rs.MoveNext
  Wend
  
  rs.Close
  
  Close fn1
End Sub
