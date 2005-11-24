VERSION 5.00
Begin VB.Form fAttuale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prova"
   ClientHeight    =   3975
   ClientLeft      =   2370
   ClientTop       =   2205
   ClientWidth     =   5340
   Icon            =   "fAttuale.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3975
   ScaleWidth      =   5340
   Begin VB.CommandButton bTrova 
      Caption         =   "&Trova"
      Height          =   375
      Left            =   840
      TabIndex        =   18
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox tQ 
      Height          =   285
      Left            =   1080
      TabIndex        =   17
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox tM 
      Height          =   285
      Left            =   1080
      TabIndex        =   16
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton bRetta 
      Caption         =   "&Retta"
      Height          =   375
      Left            =   840
      TabIndex        =   15
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Esci"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox tADC 
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Text            =   "2048"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox tValoff 
      Height          =   285
      Left            =   3480
      TabIndex        =   8
      Text            =   "0"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox tValMax 
      Height          =   285
      Left            =   3480
      TabIndex        =   7
      Text            =   "5"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox tValMin 
      Height          =   285
      Left            =   3480
      TabIndex        =   6
      Text            =   "0"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox tBitMax 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Text            =   "4095"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox tBitMin 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Calcola"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1920
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "ADC value"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Offset"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   1590
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Valore massimo"
      Height          =   225
      Left            =   1800
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Valore minimo"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   900
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Valore bit max (ADC)"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Valore bit min (ADC)"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   165
      Width           =   1455
   End
End
Attribute VB_Name = "fAttuale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub bRetta_Click()
    Dim m As Single
    Dim q As Single
    m = 0
    q = 0
    TrovaRetta Val(tBitMin), Val(tBitMax), _
    Val(tValMin), Val(tValMax), _
    m, q
    Text1.Text = Str(m) + "," + Str(q)
    tM.Text = m
    tQ.Text = q
End Sub

Private Sub bTrova_Click()
    Dim valMin As Single
    Dim valMax As Single
    Dim valOff As Single
    TrovaBitVal Val(tM), Val(tQ), _
    Int(Val(tBitMin)), Int(Val(tBitMax)), _
    valMin, valMax, valOff
    tValMin.Text = valMin
    tValMax.Text = valMax
    tValoff.Text = valOff
    
End Sub

Private Sub Command1_Click()
    Dim Valore As Single
    Valore = adc2value(Val(tADC), _
    Val(tBitMin), _
    Val(tBitMax), Val(tValMax), _
    Val(tValMin), Val(tValoff))
    Text1.Text = Str(Valore) + vbCrLf
    Valore = adc2value2(Val(tADC), _
    Val(tBitMin), _
    Val(tBitMax), Val(tValMax), _
    Val(tValMin), Val(tValoff))
    Text1.Text = Text1.Text + Str(Valore)

End Sub

Private Sub Command2_Click()
    Me.Hide
    fMain.Show
End Sub
