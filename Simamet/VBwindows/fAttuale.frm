VERSION 5.00
Begin VB.Form fAttuale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acquisizione attuale"
   ClientHeight    =   2055
   ClientLeft      =   2370
   ClientTop       =   2205
   ClientWidth     =   3645
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2055
   ScaleWidth      =   3645
   Begin VB.Timer Timer1 
      Left            =   1680
      Top             =   1080
   End
   Begin VB.CommandButton bNuovo 
      Caption         =   "&Nuovo valore"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Frame frCanale2 
      Caption         =   "Temperatura °C"
      Height          =   735
      Left            =   2040
      TabIndex        =   4
      Top             =   360
      Width           =   1335
      Begin VB.Label lCanale2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Canale2"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frCanale1 
      Caption         =   "Pioggia in mm"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1335
      Begin VB.Label lCanale1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Canale1"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton bFine 
      Caption         =   "&Fine"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Lettura attuale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "fAttuale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub PrendiDati()
    
    Dim Dummy1 As String
    Dim PosVirgola As Integer
    Dim Dummy2 As String
    fMain.MSComm1.InputLen = 1
    'fmain.Timeout.Enabled = True
    'fmain.Timeout.Interval = TmOut
    dummy = ""
    Linea = ""
    Do Until dummy = vbLf
        DoEvents
        dummy = fMain.MSComm1.Input
        Linea = Linea + dummy
    Loop
    'fmain.Timeout.Enabled = False
    PosVirgola = InStr(Linea, ",")
    If PosVirgola <> 0 Then
        Dummy2 = Mid(Linea, PosVirgola + 1, Len(Linea))
        Dummy1 = Left(Linea, PosVirgola - 1)
        'elimina lo zero iniziale se c'e'
        If Left(Dummy1, 1) = "0" Then Dummy1 = Right(Dummy1, Len(Dummy1) - 1)
    Else
        Dummy1 = "0.0"
        Dummy2 = "0.0"
    End If
    'For i = 1 To Len(Linea)
    '    dums = dums + Str(Asc(Mid(Linea, i, 1))) + " "
    'Next
        
    lCanale1.Caption = Dummy1
    lCanale2.Caption = Dummy2
    bFine.Enabled = True
End Sub

Private Sub bFine_Click()
    Timer1.Enabled = False
    Timer1.Interval = 0
    'Manda un 99
    OpenCom
    fMain.MSComm1.Output = StopPrg + vbCrLf
    CloseCom
    fAttuale.MousePointer = vbNormal
    Unload Me
    fMain.Show
End Sub

Private Sub bNuovo_Click()
    
    'Manda uno 0
    fMain.MSComm1.Output = "0" + vbCrLf
    'fmain.mscomm1.Output = vbCrLf
    'Attendi l 'eco
    Do Until dummy = vbLf
        DoEvents
        dummy = fMain.MSComm1.Input
        Linea = Linea + dummy
    Loop
    PrendiDati
    
End Sub

Private Sub Form_Load()
    fMain.Hide
    bNuovo.Enabled = False
    bNuovo.Visible = False
     'Apre la porta
    OpenCom
    fMain.MSComm1.InputLen = 1
    fMain.MSComm1.Output = Chr$(3)
    'Attende l'Ok
    'fmain.Timeout.Enabled = True
    'fmain.Timeout.Interval = TmOut
    Do Until dummy = vbLf
        DoEvents
        dummy = fMain.MSComm1.Input
        Linea = Linea + dummy
    Loop
    'fmain.Timeout.Enabled = False
    dummy = ""
    fMain.MSComm1.Output = Prova + vbCrLf
    'Attende l'eco
    'fmain.Timeout.Enabled = True
    'fmain.Timeout.Interval = TmOut
    Do Until dummy = vbLf
        DoEvents
        dummy = fMain.MSComm1.Input
        Linea = Linea + dummy
    Loop
    'fmain.Timeout.Enabled = False
    PrendiDati
    Timer1.Interval = 1000
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    bFine.Enabled = False
    fAttuale.MousePointer = vbHourglass
    bNuovo_Click
    fAttuale.MousePointer = vbNormal
End Sub
