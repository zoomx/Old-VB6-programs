VERSION 5.00
Object = "{1A028BB9-0262-48D8-9B67-D74CE4C2A7E8}#3.0#0"; "IngotTextBoxCtl.dll"
Object = "{BAE83016-CA11-4D8A-BBFA-0AE9863B82DE}#3.0#0"; "IngotLabelCtl.dll"
Object = "{9F4EED48-8EC7-4316-A47D-F6161874E478}#3.0#0"; "IngotButtonCtl.dll"
Begin VB.Form SaveFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Record Informations"
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
   Begin IngotButtonCtl.AFButton CancelBTN 
      Height          =   135
      Left            =   1680
      OleObjectBlob   =   "SaveFrm.frx":0000
      TabIndex        =   15
      Top             =   2160
      Width           =   615
   End
   Begin IngotButtonCtl.AFButton SaveBtn 
      Height          =   135
      Left            =   120
      OleObjectBlob   =   "SaveFrm.frx":004C
      TabIndex        =   14
      Top             =   2160
      Width           =   615
   End
   Begin IngotTextBoxCtl.AFTextBox AirFLD 
      Height          =   180
      Left            =   1800
      OleObjectBlob   =   "SaveFrm.frx":0095
      TabIndex        =   13
      Top             =   1125
      Width           =   600
   End
   Begin IngotTextBoxCtl.AFTextBox SoilFLD 
      Height          =   180
      Left            =   600
      OleObjectBlob   =   "SaveFrm.frx":00F9
      TabIndex        =   12
      Top             =   1125
      Width           =   600
   End
   Begin IngotTextBoxCtl.AFTextBox BarFLD 
      Height          =   180
      Left            =   600
      OleObjectBlob   =   "SaveFrm.frx":015D
      TabIndex        =   11
      Top             =   945
      Width           =   600
   End
   Begin IngotTextBoxCtl.AFTextBox LongFLD 
      Height          =   180
      Left            =   600
      OleObjectBlob   =   "SaveFrm.frx":01C1
      TabIndex        =   10
      Top             =   765
      Width           =   1440
   End
   Begin IngotTextBoxCtl.AFTextBox LatFLD 
      Height          =   180
      Left            =   600
      OleObjectBlob   =   "SaveFrm.frx":0222
      TabIndex        =   9
      Top             =   585
      Width           =   1440
   End
   Begin IngotTextBoxCtl.AFTextBox SpotFLD 
      Height          =   180
      Left            =   600
      OleObjectBlob   =   "SaveFrm.frx":0283
      TabIndex        =   8
      Top             =   405
      Width           =   1440
   End
   Begin IngotTextBoxCtl.AFTextBox AreaFLD 
      Height          =   180
      Left            =   600
      OleObjectBlob   =   "SaveFrm.frx":02E4
      TabIndex        =   7
      Top             =   225
      Width           =   1440
   End
   Begin IngotLabelCtl.AFLabel AFLabel7 
      Height          =   150
      Left            =   1200
      OleObjectBlob   =   "SaveFrm.frx":0345
      TabIndex        =   6
      Top             =   1125
      Width           =   615
   End
   Begin IngotLabelCtl.AFLabel AFLabel6 
      Height          =   150
      Left            =   0
      OleObjectBlob   =   "SaveFrm.frx":0392
      TabIndex        =   5
      Top             =   1125
      Width           =   615
   End
   Begin IngotLabelCtl.AFLabel AFLabel5 
      Height          =   150
      Left            =   0
      OleObjectBlob   =   "SaveFrm.frx":03E0
      TabIndex        =   4
      Top             =   945
      Width           =   615
   End
   Begin IngotLabelCtl.AFLabel AFLabel4 
      Height          =   150
      Left            =   0
      OleObjectBlob   =   "SaveFrm.frx":042D
      TabIndex        =   3
      Top             =   765
      Width           =   615
   End
   Begin IngotLabelCtl.AFLabel AFLabel3 
      Height          =   150
      Left            =   0
      OleObjectBlob   =   "SaveFrm.frx":047B
      TabIndex        =   2
      Top             =   585
      Width           =   615
   End
   Begin IngotLabelCtl.AFLabel AFLabel2 
      Height          =   150
      Left            =   0
      OleObjectBlob   =   "SaveFrm.frx":04C8
      TabIndex        =   1
      Top             =   405
      Width           =   615
   End
   Begin IngotLabelCtl.AFLabel AFLabel1 
      Height          =   150
      Left            =   0
      OleObjectBlob   =   "SaveFrm.frx":0511
      TabIndex        =   0
      Top             =   225
      Width           =   615
   End
End
Attribute VB_Name = "SaveFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

