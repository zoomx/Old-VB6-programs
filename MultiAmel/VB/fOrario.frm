VERSION 5.00
Begin VB.Form fOrario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data e orario di avvio"
   ClientHeight    =   2565
   ClientLeft      =   2085
   ClientTop       =   2490
   ClientWidth     =   4530
   Icon            =   "fOrario.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2565
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bOk 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox tMin 
      Height          =   285
      Left            =   3480
      TabIndex        =   0
      Text            =   "45"
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox tOra 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Text            =   "12"
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox tGiorno 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Text            =   "05"
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox tMese 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "01"
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox tAnno 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "1999"
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton bMinm 
      Caption         =   "-"
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton bMinp 
      Caption         =   "+"
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton bOram 
      Caption         =   "-"
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton bOrap 
      Caption         =   "+"
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton pGiornom 
      Caption         =   "-"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton PGiornop 
      Caption         =   "+"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton bMesem 
      Caption         =   "-"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton bMesep 
      Caption         =   "+"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton bAnnom 
      Caption         =   "-"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton bAnnop 
      Caption         =   "+"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Inserire data e ora di attivazione."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   21
      Top             =   360
      Width           =   3705
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Minuti"
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   20
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ore"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   19
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Giorno"
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   18
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Mese"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   17
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Anno"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   1080
      Width           =   495
   End
End
Attribute VB_Name = "fOrario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub Form_Activate()
'    If Versione = "MultiAmel" Then
'        tMin.Text = "0"
'        If Hour(Now) / 2 <> Hour(Now) \ 2 Then
'            tOra.Text = Hour(Now + 1 / 24) '4.16666666666667E-02)
'        End If
'    End If
'
'End Sub

Private Sub Form_Load()
    tAnno.Text = Year(Now)
    tMese.Text = Month(Now)
    tGiorno.Text = Day(Now)
    tOra.Text = Hour(Now)
    tMin.Text = Minute(Now)
    If Versione = "MultiAmel" Then
        tMin.Text = "0"
        'If Hour(Now) / 2 <> Hour(Now) \ 2 Then
            tOra.Text = Hour(Now + 1 / 24) '4.16666666666667E-02)
        'End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        'CloseCom
        Me.Hide
        Unload Me
        fMain.Show
    End If
End Sub

Private Sub bAnnom_Click()
    Dim Anno As Integer
    Anno = Val(tAnno.Text) - 1
    If Anno < 1980 Then Anno = 1980
    tAnno.Text = Anno
End Sub

Private Sub bAnnop_Click()
    Dim Anno As Integer
    Anno = Val(tAnno.Text) + 1
    If Anno > 2050 Then Anno = 2050
    tAnno.Text = Anno
End Sub

Private Sub bMesem_Click()
    Dim Mese As Integer
    Mese = Val(tMese.Text) - 1
    If Mese < 1 Then Mese = 1
    tMese.Text = Mese
End Sub

Private Sub bMesep_Click()
    Dim Mese As Integer
    Mese = Val(tMese.Text) + 1
    If Mese > 12 Then Mese = 12
    tMese.Text = Mese
End Sub

Private Sub bMinm_Click()
    Dim Minuti As Integer
    Minuti = Val(tMin.Text) - 1
    If Minuti < 0 Then Minuti = 59
    tMin.Text = Minuti
End Sub

Private Sub bMinp_Click()
    Dim Minuti As Integer
    Minuti = Val(tMin.Text) + 1
    If Minuti > 59 Then Minuti = 0
    tMin.Text = Minuti
End Sub

Private Sub bOk_Click()
    Dim Data As String
    Dim ok As Boolean
    
    PAnno = fOrario.tAnno.Text
    PMese = fOrario.tMese.Text
    PGiorno = fOrario.tGiorno.Text
    POra = fOrario.tOra.Text
    PMinuti = fOrario.tMin.Text
    Data = PAnno + "/" + PMese + "/" + PGiorno
    ok = IsDate(Data)
    If ok = False Then
        MsgBox ("Data non valida")
        Exit Sub
    End If
        
    Me.Hide
    DoEvents
    fIntervallo.Show
End Sub

Private Sub bOram_Click()
    Dim Ora As Integer
    Ora = Val(tOra.Text) - 1
    If Ora < 0 Then Ora = 23
    tOra.Text = Ora
End Sub

Private Sub bOrap_Click()
    Dim Ora As Integer
    Ora = Val(tOra.Text) + 1
    If Ora > 23 Then Ora = 0
    tOra.Text = Ora
End Sub

Private Sub pGiornom_Click()
    Dim Giorno As Integer
    Giorno = Val(tGiorno.Text) - 1
    If Giorno < 1 Then Giorno = 1
    tGiorno.Text = Giorno
End Sub

Private Sub PGiornop_Click()
    Dim Giorno As Integer
    If Giorno > 31 Then Giorno = 31
    If Giorno > 30 And Val(tMese.Text) = 4 Then Giorno = 30
    If Giorno > 30 And Val(tMese.Text) = 6 Then Giorno = 30
    If Giorno > 30 And Val(tMese.Text) = 9 Then Giorno = 30
    If Giorno > 30 And Val(tMese.Text) = 11 Then Giorno = 30
    If Giorno > 29 And Val(tMese.Text = 2) Then Giorno = 29
    If Giorno > 28 And Val(tMese.Text) = 2 And ((Val(tAnno.Text) / 4) - Int(Val(tAnno.Text) / 4)) <> 0 Then Giorno = 28
    If Giorno > 28 And Val(tMese.Text) = 2 And ((Val(tAnno.Text) / 4) - Int(Val(tAnno.Text) / 4)) = 0 Then Giorno = 29
    If Giorno > 28 And Val(tMese.Text) = 2 And Val(tAnno.Text) = 2000 Then Giorno = 29
    tGiorno.Text = Giorno
End Sub
