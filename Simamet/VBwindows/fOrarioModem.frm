VERSION 5.00
Begin VB.Form fOrarioModem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simamet"
   ClientHeight    =   2580
   ClientLeft      =   3855
   ClientTop       =   2715
   ClientWidth     =   4185
   Icon            =   "fOrarioModem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bEsci 
      Caption         =   "&Esci"
      Height          =   375
      Left            =   2250
      TabIndex        =   9
      Top             =   2040
      Width           =   825
   End
   Begin VB.CommandButton bInvia 
      Caption         =   "&Invia"
      Height          =   375
      Left            =   1020
      TabIndex        =   8
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox tsMinuti 
      Height          =   285
      Left            =   2700
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "00"
      Top             =   1500
      Width           =   315
   End
   Begin VB.TextBox tsOra 
      Height          =   285
      Left            =   2130
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "23"
      Top             =   1500
      Width           =   315
   End
   Begin VB.TextBox taMinuti 
      Height          =   285
      Left            =   2700
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "00"
      Top             =   990
      Width           =   315
   End
   Begin VB.TextBox taOra 
      Height          =   285
      Left            =   2130
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "22"
      Top             =   990
      Width           =   315
   End
   Begin VB.Label lTitolo 
      Alignment       =   2  'Center
      Caption         =   "Programmazione orario di accensione del modem"
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
      Left            =   630
      TabIndex        =   10
      Top             =   60
      Width           =   2865
   End
   Begin VB.Label Label4 
      Caption         =   "Minuti"
      Height          =   255
      Left            =   2670
      TabIndex        =   7
      Top             =   690
      Width           =   555
   End
   Begin VB.Label Label3 
      Caption         =   "Ore"
      Height          =   225
      Left            =   2130
      TabIndex        =   6
      Top             =   690
      Width           =   315
   End
   Begin VB.Label Label2 
      Caption         =   "Spegnimento"
      Height          =   255
      Left            =   870
      TabIndex        =   1
      Top             =   1530
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Accensione"
      Height          =   255
      Left            =   870
      TabIndex        =   0
      Top             =   990
      Width           =   915
   End
End
Attribute VB_Name = "fOrarioModem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const Titolo As String = "Programmazione orario di accensione del modem"
Private Const Tempo As Integer = 10

Private Sub Form_Load()
    Dim Stringa As String
    
    'Disattivazione pulsanti
    bInvia.Enabled = False
    bEsci.Enabled = False
    DoEvents
    lTitolo.Caption = Titolo
    'Invio codice
    Me.MousePointer = vbHourglass
    lTitolo.Caption = "Attendere.."
    fMain.MSComm1.Output = CTRLC
    DoEvents
    Call Sleep(250)
    Stringa = InputComTimeOut(Tempo)
    DoEvents
    fMain.MSComm1.InBufferCount = 0
    DoEvents
    fMain.MSComm1.Output = ScaricaOrarioModem + vbCr
    DoEvents
    Call Sleep(250)
    DoEvents
repeat:
    Stringa = InputComTimeOut(Tempo)
    If Stringa = "TimeOut" Then
        Errore
        GoTo uscita
    End If
    If Stringa = "Simamet" Then
        GoTo repeat
    End If
    DoEvents
    taOra.Text = stripCrLf(Stringa)
    Stringa = InputComTimeOut(Tempo)
    If Stringa = "TimeOut" Then
        Errore
        GoTo uscita
    End If
    DoEvents
    taMinuti.Text = stripCrLf(Stringa)
    Stringa = InputComTimeOut(Tempo)
    If Stringa = "TimeOut" Then
        Errore
        GoTo uscita
    End If
    DoEvents
    tsOra.Text = stripCrLf(Stringa)
    Stringa = InputComTimeOut(Tempo)
    If Stringa = "TimeOut" Then
        Errore
        GoTo uscita
    End If
    DoEvents
    tsMinuti.Text = stripCrLf(Stringa)
    
uscita:
    lTitolo.Caption = Titolo
    Me.MousePointer = vbNormal
    'riattivazione pulsanti
    bInvia.Enabled = True
    bEsci.Enabled = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    fMain.Show
End Sub

Private Sub bEsci_Click()
    Unload Me
    fMain.Show
End Sub

Private Sub bInvia_Click()
    'Disattivazione pulsanti
    bInvia.Enabled = False
    bEsci.Enabled = False
    
    'Invio codice
    Me.MousePointer = vbHourglass
    lTitolo.Caption = "Attendere.."
   
    fMain.MSComm1.Output = CTRLC
    Call Sleep(250)
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = OrarioModem + vbCr
    Call Sleep(250)
    fMain.MSComm1.Output = taOra.Text + vbCr
    Call Sleep(250)
    fMain.MSComm1.Output = taMinuti.Text + vbCr
    Call Sleep(250)
    fMain.MSComm1.Output = tsOra.Text + vbCr
    Call Sleep(250)
    fMain.MSComm1.Output = tsMinuti.Text + vbCr
    Call Sleep(250)
    fMain.MSComm1.InBufferCount = 0
'    stringa = InputComTimeOut(40)
'    If stringa = "TimeOut" Then
'        Errore
'        GoTo uscita
'    End If

uscita:
    'riattivazione pulsanti
    bInvia.Enabled = True
    bEsci.Enabled = True
    lTitolo.Caption = Titolo
    Me.MousePointer = vbNormal
End Sub


Public Sub Errore()
    MsgBox ("Errore la centralina non risponde!")
End Sub

