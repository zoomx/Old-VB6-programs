VERSION 5.00
Begin VB.Form fVelComModem 
   Caption         =   "Velocità COM"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2865
   Icon            =   "fVelComModem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   2865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   2760
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Modem GSM"
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "57600"
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "33600"
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "28800"
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "19200"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "14400"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "9600"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Velocità connessione con il modem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "fVelComModem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Settings As String
Const N81 As String = ",n,8,1"

Private Sub Form_Load()
    Settings = "57600,n,8,1"
End Sub

Private Sub bOk_Click()
    fModem.txtPortSettings = Settings
    Unload Me
    fModem.Show
End Sub

Private Sub Option1_Click(i As Integer)
    Select Case i
        Case 0
            Settings = "9600" + N81
        Case 1
            Settings = "14400" + N81
        Case 2
            Settings = "19200" + N81
        Case 3
            Settings = "28800" + N81
        Case 4
            Settings = "33600" + N81
        Case 5
            Settings = "57600" + N81
        Case 6
            Settings = "19200" + N81
        Case Else
            Settings = "57600" + N81
    End Select
End Sub
