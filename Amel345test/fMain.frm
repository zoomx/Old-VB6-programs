VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form fMain 
   Caption         =   "AMEL 345 Test"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bCond3 
      Caption         =   "Ask Cond1"
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton bCond2 
      Caption         =   "Ask Cond0"
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton bCond 
      Caption         =   "Ask Cond"
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton bAskChans 
      Caption         =   "&Ask channels"
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton bDTRstatus 
      BackColor       =   &H80000003&
      Height          =   255
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton bDTR 
      Caption         =   "&DTR"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Get data"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton bClearText 
      Caption         =   "&Clear text"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton bSelect 
      Caption         =   "&Select probe"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton bQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   4215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5400
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      InputMode       =   1
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub bCond_Click()
    AskConducibility
End Sub

Private Sub bCond2_Click()
    Dim Comando As String
    Dim iCond As Long
    Dim iScale As Long
    Dim Lungo
    Dim Cond As Double
    OpenCom
    Comando = Chr$(133) & Chr$(0) & Chr$(0) 'Scala conducibilità=0
    fMain.MSComm1.Output = Comando
    Sleep 200
    Comando = Chr$(130) & Chr$(4) & Chr$(0) 'Dammi la conducibilità
    fMain.MSComm1.Output = Comando
    GetReply 2
    
    Lungo = GetReplyb(0) + 256 * CLng(GetReplyb(1))
    iCond = Lungo And 2047
    iScale = (Lungo And 49152) / 16384
    Cond = iCond / 1000
    fMain.Text1.Text = fMain.Text1.Text + vbCrLf
    fMain.Text1.Text = fMain.Text1.Text + "Cond=" + Format(Cond, "#.####") + " "
    fMain.Text1.Text = fMain.Text1.Text + "Scala=" + Str(iScale) + " "
    fMain.Text1.Text = fMain.Text1.Text + "Value=" + Str(Lungo)
    
    
    

End Sub

Private Sub bCond3_Click()
    Dim Comando As String
    Dim iCond As Long
    Dim iScale As Long
    Dim Lungo
    Dim Cond As Double
    OpenCom
    Comando = Chr$(133) & Chr$(0) & Chr$(1) 'Scala conducibilità=0
    fMain.MSComm1.Output = Comando
    Sleep 200
    Comando = Chr$(130) & Chr$(4) & Chr$(0) 'Dammi la conducibilità
    fMain.MSComm1.Output = Comando
    GetReply 2
    Lungo = GetReplyb(0) + 256 * CLng(GetReplyb(1))
    iCond = Lungo And 2047
    iScale = (Lungo And 49152) / 16384
    Cond = iCond / 1000
    fMain.Text1.Text = fMain.Text1.Text + vbCrLf
    fMain.Text1.Text = fMain.Text1.Text + "Cond=" + Format(Cond, "#.###") + " "
    fMain.Text1.Text = fMain.Text1.Text + "Scala=" + Str(iScale) + " "
    fMain.Text1.Text = fMain.Text1.Text + "Value=" + Str(iCond)

End Sub

Private Sub bDTR_Click()
OpenCom
fMain.MSComm1.DTREnable = True
Sleep 1000
fMain.MSComm1.DTREnable = False
Sleep 1000
fMain.MSComm1.DTREnable = True
CloseCom
End Sub

Private Sub Command1_Click()
    'fMain.MSComm1.DTREnable = False
    OpenCom
    fMain.Text1.Text = fMain.Text1.Text + vbCrLf + "Sending command 129"
    fMain.MSComm1.Output = Chr$(GetDataFromProbe)
    Sleep 50
'    fMain.MSComm1.Output = GetDataFromProbe
    GetReply 16
    ParseData
    AskConducibility
End Sub
Private Sub bAskChans_Click()
    OpenCom
    fMain.Text1.Text = fMain.Text1.Text + vbCrLf + "Sending command 132"
    fMain.MSComm1.Output = Chr$(&H84)
    Sleep 50
'    fMain.MSComm1.Output = GetDataFromProbe
    GetReply 1

End Sub

Private Sub Form_Load()
    ComPort = 1
    'mscomm1.o
    'OpenCom
End Sub

Private Sub bQuit_Click()
    Unload Me
    End
End Sub

Private Sub bClearText_Click()
    Text1.Text = ""
End Sub

Private Sub bSelect_Click()
    AcquamasterOk = False
    OpenCom
    SelectProbe 1
    GetReply 8
    If AcquamasterOk = True Then
        ParseAcqReply
        
    End If
        
End Sub

