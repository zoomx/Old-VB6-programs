VERSION 5.00
Begin VB.Form fVislChnl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Canali"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "fVislChnl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   3840
   End
   Begin VB.CommandButton bFine 
      Caption         =   "Fine"
      Height          =   495
      Left            =   2280
      TabIndex        =   54
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   17
      Left            =   5235
      TabIndex        =   53
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   17
      Left            =   4515
      TabIndex        =   52
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "17"
      Height          =   255
      Index           =   17
      Left            =   3000
      TabIndex        =   51
      Top             =   3360
      Width           =   1470
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   16
      Left            =   5235
      TabIndex        =   50
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   16
      Left            =   4515
      TabIndex        =   49
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "16"
      Height          =   255
      Index           =   16
      Left            =   3000
      TabIndex        =   48
      Top             =   3000
      Width           =   1470
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   15
      Left            =   5235
      TabIndex        =   47
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   15
      Left            =   4515
      TabIndex        =   46
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "15"
      Height          =   255
      Index           =   15
      Left            =   3000
      TabIndex        =   45
      Top             =   2640
      Width           =   1470
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   14
      Left            =   5235
      TabIndex        =   44
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   14
      Left            =   4515
      TabIndex        =   43
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "14"
      Height          =   255
      Index           =   14
      Left            =   3000
      TabIndex        =   42
      Top             =   2280
      Width           =   1470
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   13
      Left            =   5235
      TabIndex        =   41
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   13
      Left            =   4515
      TabIndex        =   40
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "13"
      Height          =   255
      Index           =   13
      Left            =   3000
      TabIndex        =   39
      Top             =   1920
      Width           =   1470
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   12
      Left            =   5235
      TabIndex        =   38
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   12
      Left            =   4515
      TabIndex        =   37
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "12"
      Height          =   255
      Index           =   12
      Left            =   3000
      TabIndex        =   36
      Top             =   1560
      Width           =   1470
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   11
      Left            =   5235
      TabIndex        =   35
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   11
      Left            =   4515
      TabIndex        =   34
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "11"
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   33
      Top             =   1200
      Width           =   1470
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   10
      Left            =   5235
      TabIndex        =   32
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   10
      Left            =   4515
      TabIndex        =   31
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "10"
      Height          =   255
      Index           =   10
      Left            =   3000
      TabIndex        =   30
      Top             =   840
      Width           =   1470
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   9
      Left            =   5235
      TabIndex        =   29
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   9
      Left            =   4515
      TabIndex        =   28
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "9"
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   27
      Top             =   480
      Width           =   1470
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   8
      Left            =   2355
      TabIndex        =   26
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   8
      Left            =   1635
      TabIndex        =   25
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "8"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   24
      Top             =   3360
      Width           =   1470
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   7
      Left            =   2355
      TabIndex        =   23
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   7
      Left            =   1635
      TabIndex        =   22
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "7"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   3000
      Width           =   1470
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   6
      Left            =   2355
      TabIndex        =   20
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "6"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   1470
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   6
      Left            =   1635
      TabIndex        =   18
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   5
      Left            =   2355
      TabIndex        =   17
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   5
      Left            =   1635
      TabIndex        =   16
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   4
      Left            =   2355
      TabIndex        =   14
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   1635
      TabIndex        =   13
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   3
      Left            =   2355
      TabIndex        =   11
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   1635
      TabIndex        =   10
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1470
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   2
      Left            =   2355
      TabIndex        =   8
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   1635
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1470
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   1
      Left            =   2355
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   1635
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1470
   End
   Begin VB.Label lUnita 
      Caption         =   "unità"
      Height          =   255
      Index           =   0
      Left            =   2355
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lCanale 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   1635
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lNcanale 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1470
   End
End
Attribute VB_Name = "fVislChnl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Formato As String

Private Sub Form_Load()
    Dim nDecimali As Integer
    Dim i As Long
    Dim Stringa As String
    
    OpenCom
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = Chr$(3) 'CTRL+C
    Sleep (250)
    fMain.MSComm1.Output = TestSensori + vbCr

    'Azzera input buffer rs232
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.OutBufferCount = 0
    Sleep (250)
    For i = 0 To MaxCanali
        If Canale(i).Attivo = True Then
            lCanale(i).Caption = "0.0"
            lUnita(i).Caption = Canale(i).UnitaMisura
            lNcanale(i).Caption = Canale(i).Nome
            lNcanale(i).Enabled = True
            lCanale(i).Enabled = True
            lUnita(i).Enabled = True

        Else
            'lCanale(i).Visible = False
            lCanale(i).Caption = ""
            lUnita(i).Caption = ""
            lNcanale(i).Caption = ""
            lNcanale(i).Enabled = False
            lCanale(i).Enabled = False
            lUnita(i).Enabled = False
        End If
    Next
    If lDebug = False Then
    'rende visibili le caselle per i count
        For i = 9 To 14
            lCanale(i).Visible = False
        Next
               'Modifica il form per rendere invisibili le
       'caselle con i count
        With fVislChnl
            .Width = 4185 '6015
            .Height = 4245 ' 4740
        End With
        With bFine
            .Left = 1440 '2280
            .Top = 3735 '3840
        End With
    End If
    'Manda la programmazione dei canali
    For i = 0 To MaxCanali
        If Canale(i).Attivo = True Then
            fMain.MSComm1.Output = "1" + vbCr
        Else
            fMain.MSComm1.Output = "0" + vbCr
        End If
     Next
    
    'Aspetta l'OK
    Stringa = InputComTimeOut(6)
    
'    If Stringa <> "OK" Then
'        Stringa = "Errore! ricevuto " + Stringa + " invece di OK!"
'        MsgBox (Stringa)
'        bFine_Click
'        Exit Sub
'    End If
     
    'Aggiorna il formato dei dati
    Formato = "0"
    nDecimali = Val(frmOptions2.tDecimali.Text)
    If nDecimali < 0 Then nDecimali = 0
    If nDecimali > 7 Then nDecimali = 7
    If nDecimali <> 0 Then
        Formato = Formato + "."
        For i = 1 To nDecimali
            Formato = Formato + "0"
        Next
    End If
    
    'abilita il timer
    Timer1.Interval = 1000
    Timer1.Interval = Val(frmOptions2.tFreqAgg.Text) * 1000
    Timer1.Enabled = True
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        bFine_Click
        Exit Sub
    End If
End Sub

Private Sub bFine_Click()
    Timer1.Enabled = False
    Timer1.Interval = 0
    OpenCom
    fMain.MSComm1.Output = "0"
    Sleep (250)
    fMain.MSComm1.Output = Chr$(3) 'CTRLC
    Me.Hide
    Unload frmOptions2
    Unload Me
    fMain.Show
End Sub

Private Sub Timer1_Timer()
    bFine.Enabled = False
    Me.MousePointer = vbHourglass
    NuoviValori
End Sub

Public Sub NuoviValori()
    Dim Misura As Single
    Dim Lungo As Long
    Dim Stringa As String
    Dim i As Long
    
    fMain.MSComm1.InBufferCount = 0
    'Manda l'ordine di stampare i dati
    fMain.MSComm1.Output = "3" + vbCr
    'Acquisisce tutti i canali anche se inattivi
    For i = 0 To MaxCanali
        'Attende la risposta con timeout
        Stringa = InputComTimeOut(10)
        'Debug.Print Stringa
        If Stringa <> "TimeOut" Then
            'La converte in numeri
            Lungo = Val(Stringa)
            'Aggiorna la relativa finestra se il canale è attivo
            If Canale(i).Attivo = True Then
                'Converte il valore dell'ADC nella misura corrispondente
                Misura = adc2value(Lungo, Canale(i).Bitmin, Canale(i).Bitmax, CDbl(Canale(i).valMax), CDbl(Canale(i).valMin), CDbl(Canale(i).valOff))
                'lCanale(i).Caption = Format(Misura, "0.0##")
                lCanale(i).Caption = Format(Misura, Formato)
                'Debug.Print "Misura-->"; Misura
            End If
        Else
            MsgBox ("La centralina non risponde!")
            bFine_Click
            Exit Sub
        End If

    Next
    Me.Refresh
    Me.MousePointer = vbNormal
    bFine.Enabled = True
End Sub
