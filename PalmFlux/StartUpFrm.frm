VERSION 5.00
Object = "{A54BEB34-AAB3-4A8D-B736-42CB4DA7D664}#3.0#0"; "IngotComboBoxCtl.dll"
Object = "{BAE83016-CA11-4D8A-BBFA-0AE9863B82DE}#3.0#0"; "IngotLabelCtl.dll"
Object = "{1A028BB9-0262-48D8-9B67-D74CE4C2A7E8}#3.0#0"; "IngotTextBoxCtl.dll"
Object = "{9F4EED48-8EC7-4316-A47D-F6161874E478}#3.0#0"; "IngotButtonCtl.dll"
Object = "{899CE9D8-3C9F-48DF-B418-E338294B00E3}#3.0#0"; "IngotCheckBoxCtl.dll"
Begin VB.Form StartUpFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WEST Systems 06"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2400
   BeginProperty Font 
      Name            =   "AFPalm"
      Size            =   8.25
      Charset         =   2
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   StartUpPosition =   2  'CenterScreen
   Begin IngotCheckBoxCtl.AFCheckBox SoilTChk 
      Height          =   195
      Left            =   120
      OleObjectBlob   =   "StartUpFrm.frx":0000
      TabIndex        =   17
      Top             =   315
      Width           =   1815
   End
   Begin IngotCheckBoxCtl.AFCheckBox VaisalaChk 
      Height          =   195
      Left            =   120
      OleObjectBlob   =   "StartUpFrm.frx":005F
      TabIndex        =   15
      Top             =   510
      Width           =   1935
   End
   Begin IngotCheckBoxCtl.AFCheckBox AirTChk 
      Height          =   195
      Left            =   120
      OleObjectBlob   =   "StartUpFrm.frx":00C2
      TabIndex        =   16
      Top             =   690
      Width           =   1935
   End
   Begin IngotComboBoxCtl.AFComboBox InstTypeCmb 
      Height          =   195
      Left            =   1440
      OleObjectBlob   =   "StartUpFrm.frx":0121
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin IngotLabelCtl.AFLabel Label7 
      Height          =   195
      Left            =   0
      OleObjectBlob   =   "StartUpFrm.frx":0176
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin IngotButtonCtl.AFButton SaveBtn 
      Height          =   195
      Left            =   120
      OleObjectBlob   =   "StartUpFrm.frx":01CA
      TabIndex        =   12
      Top             =   2160
      Width           =   2055
   End
   Begin IngotComboBoxCtl.AFComboBox UnitCmb 
      Height          =   195
      Left            =   1440
      OleObjectBlob   =   "StartUpFrm.frx":0213
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin IngotTextBoxCtl.AFTextBox AckFld 
      Height          =   195
      Left            =   1440
      OleObjectBlob   =   "StartUpFrm.frx":0266
      TabIndex        =   10
      Top             =   1740
      Width           =   615
   End
   Begin IngotComboBoxCtl.AFComboBox NPCmb 
      Height          =   195
      Left            =   1440
      OleObjectBlob   =   "StartUpFrm.frx":02C4
      TabIndex        =   9
      Top             =   1530
      Width           =   975
   End
   Begin IngotComboBoxCtl.AFComboBox SICmb 
      Height          =   195
      Left            =   1440
      OleObjectBlob   =   "StartUpFrm.frx":0310
      TabIndex        =   8
      Top             =   1320
      Width           =   855
   End
   Begin IngotComboBoxCtl.AFComboBox FSCmb 
      Height          =   195
      Left            =   1440
      OleObjectBlob   =   "StartUpFrm.frx":035C
      TabIndex        =   7
      Top             =   1110
      Width           =   975
   End
   Begin IngotComboBoxCtl.AFComboBox RS232Cmb 
      Height          =   195
      Left            =   1440
      OleObjectBlob   =   "StartUpFrm.frx":03AF
      TabIndex        =   6
      Top             =   900
      Width           =   975
   End
   Begin IngotLabelCtl.AFLabel Label6 
      Height          =   195
      Left            =   0
      OleObjectBlob   =   "StartUpFrm.frx":0402
      TabIndex        =   5
      Top             =   1110
      Width           =   1335
   End
   Begin IngotLabelCtl.AFLabel Label5 
      Height          =   195
      Left            =   0
      OleObjectBlob   =   "StartUpFrm.frx":045A
      TabIndex        =   4
      Top             =   1950
      Width           =   1335
   End
   Begin IngotLabelCtl.AFLabel Label4 
      Height          =   195
      Left            =   0
      OleObjectBlob   =   "StartUpFrm.frx":04A8
      TabIndex        =   3
      Top             =   1740
      Width           =   1335
   End
   Begin IngotLabelCtl.AFLabel Label3 
      Height          =   195
      Left            =   0
      OleObjectBlob   =   "StartUpFrm.frx":04FF
      TabIndex        =   2
      Top             =   1530
      Width           =   1335
   End
   Begin IngotLabelCtl.AFLabel Label2 
      Height          =   195
      Left            =   0
      OleObjectBlob   =   "StartUpFrm.frx":055A
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin IngotLabelCtl.AFLabel Label1 
      Height          =   195
      Left            =   0
      OleObjectBlob   =   "StartUpFrm.frx":05B4
      TabIndex        =   0
      Top             =   900
      Width           =   1335
   End
End
Attribute VB_Name = "StartUpFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ReadConfiguration "INFO.INF", StartUpFrm
    Old_Info_Check = Info_Check
    
    StartUpFrm.InstTypeCmb.Clear
    StartUpFrm.InstTypeCmb.AddItem Licor
    StartUpFrm.InstTypeCmb.AddItem Drager
    StartUpFrm.InstTypeCmb.AddItem Politron
    StartUpFrm.InstTypeCmb.AddItem NI6B11
    StartUpFrm.InstTypeCmb.AddItem NI6B13
    StartUpFrm.InstTypeCmb.AddItem NI6B11_N
    StartUpFrm.InstTypeCmb.AddItem Riken

    StartUpFrm.InstTypeCmb.Text = Info_Check
    ShowConfiguration
End Sub

Private Sub SaveBtn_Click()
Dim ec As Integer
 Select Case Info_Check
  Case Drager
        ec = CheckDragerInformations
  Case Licor
        ec = CheckLICORInformations
  Case Politron
        ec = CheckPolitronInformations
  Case NI6B11
        ec = CheckNI6B11Informations
  Case NI6B13
        ec = CheckNI6B13Informations
  Case NI6B11_N
       ec = CheckNI6B11_N_Informations
  Case Riken
       ec = CheckRikenInformations
  Case Else
        MsgBox "Not valid " + Info_Check
  End Select
  
  
  If IsNumeric(CInt(RS232Cmb.Text)) Then
     Info_RS232_Port = CInt(RS232Cmb.Text)
  Else
     Info_RS232_Port = 1
  End If
  If Not ec Then
     MsgBox "Please check the value, I can't save it"
     Exit Sub
  End If
  
  SaveConfiguration ("INFO.INF")
  SaveBtn.Caption = "File saved, I'll end...."
  Caption = "File saved, I'll end...."
  
  'Delay (1)
  SaveBtn.Caption = "Shutting down...."
  Caption = "Shutting down...."
  'Delay (1)
  
  End

End Sub
Private Sub InstTYPECMB_Click()
    Dim NewCheck As String
    NewCheck = InstTypeCmb.List(InstTypeCmb.ListIndex)
    If NewCheck <> Info_Check Then
       Info_Check = NewCheck
       ShowConfiguration
    End If
End Sub
