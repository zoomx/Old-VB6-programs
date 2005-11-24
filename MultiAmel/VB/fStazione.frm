VERSION 5.00
Begin VB.Form fStazione 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup"
   ClientHeight    =   2790
   ClientLeft      =   3615
   ClientTop       =   3900
   ClientWidth     =   4635
   Icon            =   "fStazione.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2790
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1785
      MaxLength       =   1
      TabIndex        =   7
      Text            =   ";"
      Top             =   1470
      Width           =   225
   End
   Begin VB.TextBox tDirSaveFiles 
      Height          =   285
      Left            =   105
      MaxLength       =   20
      TabIndex        =   5
      Top             =   945
      Width           =   4380
   End
   Begin VB.CommandButton bAnnulla 
      Caption         =   "&Annulla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   2520
      TabIndex        =   2
      Top             =   2205
      Width           =   1215
   End
   Begin VB.CommandButton bContinua 
      Caption         =   "&Continua"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1050
      TabIndex        =   1
      Top             =   2205
      Width           =   1215
   End
   Begin VB.TextBox tDirProgrammazione 
      Height          =   285
      Left            =   105
      MaxLength       =   20
      TabIndex        =   0
      Top             =   315
      Width           =   4380
   End
   Begin VB.Label Label3 
      Caption         =   "Separatore decimali"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   6
      Top             =   1470
      Width           =   1485
   End
   Begin VB.Label Label2 
      Caption         =   "Cartella iniziale files di salvataggio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   720
      TabIndex        =   4
      Top             =   735
      Width           =   2985
   End
   Begin VB.Label Label1 
      Caption         =   "Cartella iniziale files di programmazione"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   600
      TabIndex        =   3
      Top             =   105
      Width           =   3345
   End
End
Attribute VB_Name = "fStazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    fMain.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Me.Hide
        Unload Me
        fMain.Show
    End If
End Sub

Private Sub bAnnulla_Click()
    Unload Me
    fMain.Show
End Sub

Private Sub bContinua_Click()
    Me.Hide
    fMain.Show
End Sub
