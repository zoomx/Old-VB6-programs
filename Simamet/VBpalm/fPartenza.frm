VERSION 5.00
Begin VB.Form fPartenza 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programmazione orario di partenza"
   ClientHeight    =   2790
   ClientLeft      =   3645
   ClientTop       =   4035
   ClientWidth     =   5055
   ForeColor       =   &H00000000&
   Icon            =   "fPartenza.frx":0000
   LinkTopic       =   "fPartenza"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2790
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bFine 
      Caption         =   "A&nnulla"
      Height          =   495
      Left            =   3288
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton bIndietro 
      Caption         =   "< &Indietro"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton bData 
      Caption         =   "&Data Prefissata"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton bAdesso 
      Caption         =   "&Adesso"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "ATTENZIONE Verrà utilizzato l'orologio del PC. Controllare che sia esatto!"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Orario avvio acquisizione"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "fPartenza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        'CloseCom
        Me.Hide
        Unload Me
        fMain.Show
    End If
End Sub

Private Sub bAdesso_Click()
    PAnno = Year(Now)
    PMese = Month(Now)
    PGiorno = Day(Now)
    POra = Hour(Now)
    PMinuti = Minute(Now)

    Orario = "NOW"
    Me.Hide
    fIntervallo.Show
End Sub

Private Sub bData_Click()
    fOrario.Show 1
    Orario = "Orario"
    Me.Hide
    fIntervallo.Show
End Sub

Private Sub bFine_Click()
    Unload Me
    fMain.Show
End Sub

Private Sub bIndietro_Click()
    Me.Hide
    frmOptions.Show
End Sub
