VERSION 5.00
Begin VB.Form fCom 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Porta seriale"
   ClientHeight    =   3420
   ClientLeft      =   4155
   ClientTop       =   3495
   ClientWidth     =   4305
   Icon            =   "fCom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3420
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton oCom1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "COM 1"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1125
      Picture         =   "fCom.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   2055
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton bFine 
      Caption         =   "&Ok"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1668
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.OptionButton oCom2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "COM 2"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Simamet"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1095
      TabIndex        =   5
      Top             =   720
      Width           =   2130
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Selezionare la porta di comunicazione seriale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   225
      TabIndex        =   4
      Top             =   1560
      Width           =   3855
   End
End
Attribute VB_Name = "fCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Eventualmente mettere qui un test sulle com esistenti
    oCom1.value = False
    oCom2.value = False
    bFine.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        CloseCom
        Me.Hide
        Unload Me
        fMain.Show
    End If
End Sub

Private Sub bFine_Click()
If fMain.MSComm1.PortOpen = False Then
    If oCom1.value = True Then ComPort = 1
    If oCom2.value = True Then ComPort = 2
    fMain.MSComm1.CommPort = ComPort
    fMain.MSComm1.Settings = "19200,n,8,1"
    fMain.MSComm1.InBufferSize = 2048
   'Altri settaggi com
    fMain.MSComm1.Handshaking = comNone
    fMain.MSComm1.RTSEnable = False

    fMain.bProgramma.Enabled = True
    fMain.bScarica.Enabled = True
    fMain.bTestSensori.Enabled = True
    fMain.bTaraBatt.Enabled = True
End If
    
    Unload Me
    fMain.Show
End Sub

Private Sub oCom1_Click()
    bFine.Enabled = True
End Sub

Private Sub oCom2_Click()
    bFine.Enabled = True
End Sub

