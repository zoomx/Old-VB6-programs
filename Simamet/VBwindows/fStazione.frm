VERSION 5.00
Begin VB.Form fStazione 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nome Stazione"
   ClientHeight    =   2415
   ClientLeft      =   3615
   ClientTop       =   3900
   ClientWidth     =   4635
   Icon            =   "fStazione.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2415
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bIndietro 
      Caption         =   "< &Indietro"
      Height          =   468
      Left            =   360
      TabIndex        =   3
      Top             =   1704
      Width           =   1215
   End
   Begin VB.CommandButton bAnnulla 
      Caption         =   "&Annulla"
      Height          =   468
      Left            =   3000
      TabIndex        =   4
      Top             =   1704
      Width           =   1215
   End
   Begin VB.CommandButton bContinua 
      Caption         =   "&Continua >"
      Default         =   -1  'True
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox tStazione 
      Height          =   285
      Left            =   1350
      MaxLength       =   20
      TabIndex        =   1
      Text            =   "MH4"
      Top             =   990
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Inserire il nome della località"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   660
      TabIndex        =   0
      Top             =   240
      Width           =   3552
   End
End
Attribute VB_Name = "fStazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    tStazione.Text = Stazione
End Sub

Private Sub Form_Load()
    fMain.Hide
    tStazione.Text = Stazione
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        'CloseCom
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
    Stazione = tStazione.Text
    Me.Hide
    fPartenza.Show
End Sub

Private Sub bIndietro_Click()
     Stazione = tStazione.Text
     Me.Hide
    frmOptions.Show
End Sub

