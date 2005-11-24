VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form fTara 
   Caption         =   "Taratura"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   Icon            =   "fTara.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bStop 
      Caption         =   "&Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6090
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton bFine 
      Caption         =   "&Esci"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6090
      TabIndex        =   5
      Top             =   3780
      Width           =   855
   End
   Begin VB.CommandButton bTarapH 
      Caption         =   "Tara &pH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1155
      TabIndex        =   4
      Top             =   3780
      Width           =   855
   End
   Begin VB.CommandButton bTaraT 
      Caption         =   "Tara &T1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   315
      TabIndex        =   3
      Top             =   3780
      Width           =   855
   End
   Begin VB.CommandButton bTaraCond 
      Caption         =   "Tara &Cd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1995
      TabIndex        =   2
      Top             =   3780
      Width           =   855
   End
   Begin VB.CommandButton bTaraBatt 
      Caption         =   "Tara&Batt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2835
      TabIndex        =   1
      Top             =   3780
      Width           =   975
   End
   Begin MSChartLib.MSChart MSChart1 
      Height          =   3795
      Left            =   105
      OleObjectBlob   =   "fTara.frx":0442
      TabIndex        =   0
      Top             =   105
      Width           =   6000
   End
   Begin VB.Label lCounts 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Counts"
      Height          =   330
      Left            =   6090
      TabIndex        =   7
      Top             =   945
      Width           =   750
   End
   Begin VB.Label lChannel 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Channel"
      Height          =   330
      Left            =   6090
      TabIndex        =   6
      Top             =   420
      Width           =   750
   End
End
Attribute VB_Name = "fTara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tarato As Boolean
Dim Formato As String
Dim continua As Boolean
Dim X(50) As Integer
Dim Y(50) As Integer

Private Sub Form_Load()
    Dim i As Integer
    'If Tarato = True Then SalvaSetup
    
    Me.Caption = "Taratura"
    lChannel.Caption = ""
    lCounts.Caption = ""
    'CaricaSetup
    Tarato = False
    continua = True
    bStop.Enabled = False
    'Disegno falso diagramma
    For i = 0 To 50
        X(i) = i
        Y(i) = 4095 * Cos(i / 180 * 3.14 - 3.14 / 2 + 0.7) - 1532
    Next
    MSChart1.chartType = VtChChartType2dLine
    MSChart1.ColumnCount = 1
    MSChart1.Column = 1
    MSChart1.Data = 1
    MSChart1.ChartData = Y
    MSChart1.Title = ""

    CaricaSetup
    Tarato = False
    'Accensione sensori
    fMain.MSComm1.Output = CTRLC
    Call Sleep(250)
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = ExternalOn + vbCr
    fMain.MSComm1.Output = INHOn + vbCr

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        bFine_Click
        Exit Sub
    End If
End Sub

Private Sub bFine_Click()
 
    If Tarato = True Then SalvaSetup
    'Spegnimento sensori
    fMain.MSComm1.Output = CTRLC
    Call Sleep(250)
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = ExternalOff + vbCr
    fMain.MSComm1.Output = INHOff + vbCr

    Me.Hide
    Unload Me
    fMain.Show
End Sub

Private Sub bStop_Click()
    continua = False
    bStop.Enabled = False
End Sub

Private Sub bTaraCond_Click()
    Dim cnd1 As Double
    Dim cnd2 As Double
    Dim c1 As Double
    Dim c2 As Double
    Dim m As Single
    Dim q As Single
    Dim Risposta As Integer
    Dim Stringa As String
    Dim valMin As Single
    Dim valMax As Single
    Dim valOff As Single
    Dim T As Double
    Dim Lungo As Long
    
    DisabilitaTasti
    
    'controllare che il termometro sia già stato
    'tarato ed emettere un messaggio appropriato
    'Text1.Text = ""
    c1 = 0
    cnd1 = 0
    Stringa = "Taratura conducibilità"
    Messaggio = "Metti il sensore di conducibilità in" + vbCrLf
    Messaggio = Messaggio + "liquido a conducibilità" + vbCrLf
    Messaggio = Messaggio + "nota." + vbCrLf
    Messaggio = Messaggio + "Quando sei pronto premi OK"
    Risposta = MsgBox(Messaggio, vbOKCancel, Stringa)
    If Risposta = vbCancel Then GoTo uscita
    
    Lungo = TaraCanale(1)
    
    If Lungo = 0 Then
        Errore
        GoTo uscita
    End If
    Debug.Print Stringa
    'Label2.Caption = ""
    c2 = Lungo
    'Text1.Text = "Misurati " + Str(c2) + " count"
    
    'prende anche una temperatura
'    fMain.MSComm1.Output = CTRLC
'    fMain.MSComm1.Output = CTRLC
'    Call Sleep(250)
'    fMain.MSComm1.InBufferCount = 0
'    fMain.MSComm1.Output = TaraTempE + vbCr
    'Call Sleep(250)
    'fMain.MSComm1.InBufferCount = 0
'    stringa = InputComTimeOut(6)
'    If stringa = "TimeOut" Then
'        Errore
'        GoTo uscita
'    End If
'    If Val(stringa) = 0 Then 'Gae 13mar2000
'        Errore
'        GoTo uscita
'    End If
    Me.MousePointer = vbNormal

'    'Calcolo temperatura
'    Intero = Val(stringa)
'    T = adc2value(CLng(Intero), Canale(0).Bitmin, _
'    Canale(0).Bitmax, Canale(0).valMax, _
'    Canale(0).valMin, Canale(0).valOff)

    Messaggio = "Immetti la conducibilità" + vbCrLf
    Messaggio = Messaggio + "in mS riferiti a 25°C"
    Stringa = InputBox(Messaggio, "Taratura sensore conducibilità", "12.88")
    If Stringa = "" Then
        'Text1.Text = ""
        GoTo uscita
    End If

    cnd2 = Val2(Stringa)
    'decorrezione per la temperatura
    
    'CondCorr(25 °C) = CondMis(T °C) / (1 + 0.0191 * (T - 25))
    
    'cnd2 = cnd2 * (1 + 0.0191 * (T - 25)) 'Cond. misurata alla temp. T
    
    'Text1.Text = Text1.Text + " a " + Str(cnd2) + "mS" + vbCrLf
    
    'Non linearità
    'cnd2 = 1 / cnd2 'Resistività a T °C
    
    'trova i parametri della retta
    TrovaRetta c1, c2, cnd1, cnd2, m, q
    'Trova i parametri per simapro
    TrovaBitVal m, q, 0, 4095, valMin, valMax, 0
    Canale(1).valMin = valMin
    Canale(1).valMax = valMax
    Canale(1).valOff = valOff
    Canale(1).Bitmin = 0
    Canale(1).Bitmax = 4095

    'Label2.Caption = "Taratura effettuata"
    Canale(1).UnitaMisura = "mS" 'Gae 13mar2000
    Tarato = True
uscita:
    AbilitaTasti


End Sub


Private Sub bTarapH_Click()
    Dim Lungo As Long
'    Lungo = TaraCanale(3)
'    MsgBox Str(Lungo)
'    Exit Sub
    'Disabilita gli altri tasti tranne esci
    Dim Intero As Integer
    Dim Risposta As Integer
    Dim Stringa As String
    Dim Volt7 As Double
    Dim T7 As Double
    Dim pH As Double
    Dim T As Double
    Dim m As Single
    Dim q As Single
    Dim c1 As Double
    Dim c2 As Double
    Dim pH1 As Double
    Dim pH2 As Double
    Dim valMin As Single
    Dim valMax As Single
    Dim valOff As Single

    DisabilitaTasti
    
    'Text1.Text = ""
    Messaggio = "Metti il sensore in una" + vbCrLf
    Messaggio = Messaggio + "soluzione tampone a pH 7 e premi OK"
    Stringa = "Taratura sensore pH"
    Risposta = MsgBox(Messaggio, vbOKCancel, Stringa)
    If Risposta = vbCancel Then GoTo uscita

    Lungo = TaraCanale(3)

    If Lungo = 0 Then
        Errore
        GoTo uscita
    End If

    c1 = Lungo
    pH1 = 7

    'prende anche una temperatura
    'fMain.MSComm1.Output = CTRLC
    'Call Sleep(250)
    'fMain.MSComm1.InBufferCount = 0
    'fMain.MSComm1.Output = TaraTempE + vbCr
    'Call Sleep(250)
    'stringa = InputComTimeOut(6)
    'If stringa = "TimeOut" Then
    '    Errore
    '    GoTo uscita
    'End If
    'Intero = Val(stringa)
    'T = adc2value(Intero, Canale(0).Bitmin, _
    Canale(0).Bitmax, Canale(0).valMax, _
    Canale(0).valMin, Canale(0).valOff)

'    Messaggio = "Se la soluzione non e' a Ph 7" + vbCrLf
'    Messaggio = Messaggio + " a causa della temperatura" + vbCrLf
'    Messaggio = Messaggio + "allora immetti qui la temperatura" + vbCrLf
'    Messaggio = Messaggio + "altrimenti lascia lo zero"
'    T7 = Val(InputBox(Messaggio, "Correzione temperatura", "0"))
'    Messaggio = ""


    'decorrrezione temperatura
    'pH1 = pH1 * (273 + 21) / (273 + T7)

    Volt7 = 5 / 4095 * Intero
    Messaggio = "Volt a pH 7=" + Format(Volt7, "0.0#")
    'Text1.Text = Messaggio
    Zero = Volt7
daccapo:
    Stringa = "Taratura sensore pH"
    Messaggio = "Adesso togli il sensore dalla" + vbCrLf
    Messaggio = Messaggio + "soluzione tampone a pH 7" + vbCrLf
    Messaggio = Messaggio + "sciacqualo con acqua distillata" + vbCrLf
    Messaggio = Messaggio + "e mettilo in una soluzione a pH noto." + vbCrLf
    Messaggio = Messaggio + "Premi OK quando sei pronto." + vbCrLf
    Risposta = MsgBox(Messaggio, vbOKCancel, Stringa)
    If Risposta = vbCancel Then GoTo uscita

    Lungo = TaraCanale(3)
    
    c2 = Lungo
    Volt7 = 5 / 4095 * Lungo
    If Volt7 = Zero Then
        Messaggio = "ERRORE! La misura sembra" + vbCrLf
        Messaggio = Messaggio + "identica alla precedente!"
        MsgBox (Messaggio)
        GoTo daccapo
    End If
    Stringa = InputBox("Immetti il pH della soluzione tampone", "Taratura sensore pH", "4")
    If Stringa = "" Then
        'Text1.Text = ""
        GoTo uscita
    End If
    If Val(Stringa) = 0 Then
        Errore
        GoTo uscita
    End If
    
    pH = Val2(Stringa)
    
    'decorrrezione temperatura
    'pH = pH * (273 + 21) / (273 + T)
    pH2 = pH
  
    Messaggio = "Volt a pH" + Str(pH) + "=" + Format(Volt7, "0.0#")

    'Text1.Text = Text1.Text + vbCrLf + Messaggio
    KpH = (pH - 7) / (Volt7 - Zero)
    'Text1.Text = Text1.Text + vbCrLf + "KpH=" + Str(KpH)
    
    
    'trova i parametri della retta
    TrovaRetta c1, c2, pH1, pH2, m, q
    'Trova i parametri per simapro
    TrovaBitVal m, q, 0, 4095, valMin, valMax, 0
    Canale(3).valMin = valMin
    Canale(3).valMax = valMax
    Canale(3).valOff = valOff
    Canale(3).Bitmin = 0
    Canale(3).Bitmax = 4095
    Canale(3).UnitaMisura = "pH"

    Tarato = True
uscita:
    AbilitaTasti


End Sub


Private Sub bTaraT_Click()
    Dim t1 As Double
    Dim t2 As Double
    Dim c1 As Double
    Dim c2 As Double
    Dim m As Single
    Dim q As Single
    Dim Risposta As Integer
    Dim Stringa As String
    Dim valMin As Single
    Dim valMax As Single
    Dim valOff As Single
    Dim T As Single
    Dim Lungo As Long
    
    DisabilitaTasti
    
    'Text1.Text = ""
    Stringa = "Taratura temperatura"
    Messaggio = "Metti il sensore di temperatura in" + vbCrLf
    Messaggio = Messaggio + "ambiente a temperatura" + vbCrLf
    Messaggio = Messaggio + "costante." + vbCrLf
    Messaggio = Messaggio + "Quando sei pronto premi OK"
    Risposta = MsgBox(Messaggio, vbOKCancel, Stringa)
    If Risposta = vbCancel Then GoTo uscita
    
    Lungo = TaraCanale(0)

    If Lungo = 0 Then
        Errore
        GoTo uscita
    End If
    
    'Label2.Caption = ""
    Me.MousePointer = vbNormal
    c1 = Lungo
    'Text1.Text = "Misurati " + Str(c1) + " count"
    Stringa = InputBox("Immetti la temperatura", "Taratura sensore Temperatura", "15")
    If Stringa = "" Then
        'Text1.Text = ""
        GoTo uscita
    End If

    t1 = Val2(Stringa)
    'Text1.Text = Text1.Text + " a " + Str(t1) + "°C" + vbCrLf
daccapo:
    Messaggio = "Adesso metti il sensore" + vbCrLf
    Messaggio = Messaggio + "in un ambiente a temperatura" + vbCrLf
    Messaggio = Messaggio + "diversa." + vbCrLf
    Messaggio = Messaggio + "Quando sei pronto premi OK"
    Risposta = MsgBox(Messaggio, vbOKCancel, Stringa)
    If Risposta = vbCancel Then GoTo uscita
    
    Lungo = TaraCanale(0)

    If Lungo = 0 Then
        Errore
        GoTo uscita
    End If
    
    Me.MousePointer = vbNormal
    c2 = Lungo
    If c2 = c1 Then
        Messaggio = "ERRORE! La misura sembra" + vbCrLf
        Messaggio = Messaggio + "identica alla precedente!"
        MsgBox (Messaggio)
        GoTo daccapo
    End If
    'Text1.Text = Text1.Text + "Misurati " + Str(c2) + " count"
    Stringa = InputBox("Immetti la temperatura", "Taratura sensore Temperatura", "25")
    If Stringa = "" Then
        'Text1.Text = ""
        GoTo uscita
    End If

    t2 = Val2(Stringa)
    'Text1.Text = Text1.Text + " a " + Str(t2) + "°C" + vbCrLf
    'trova i parametri della retta
    TrovaRetta c1, c2, t1, t2, m, q
    'Trova i parametri per simapro
    TrovaBitVal m, q, CSng(Canale(0).Bitmin), _
    CSng(Canale(0).Bitmax), valMin, valMax, _
    valOff
    Canale(0).valMin = valMin
    Canale(0).valMax = valMax
    Canale(0).valOff = valOff
    Canale(0).UnitaMisura = "C"
    'Canale(0).Bitmin = 0
    'Canale(0).Bitmax = 4095

    'Label2.Caption = "Taratura effettuata"
    Tarato = True
uscita:
'    bTarapH.Enabled = True
'    bTaraT.Enabled = True
'    bTaraT2.Enabled = True
'    bTaraCond.Enabled = True
'    bFine.Enabled = True
'    bTaraBatt.Enabled = True
    AbilitaTasti
    
End Sub
Private Sub bTaraBatt_Click()
    Dim Fattore1 As Double
    Dim VoltMisurati As Double
    Dim VoltEffettivi As Double
    Dim VoltConvertitore As Double
    Dim Fatt As Double
    Dim Stringa As String
    Dim Lungo As Long
    
    DisabilitaTasti

    Call Sleep(500)
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = LeggiBattFact + vbCr
    Stringa = InputComTimeOut(5)
    If Stringa <> "TimeOut" Then
        Fattore1 = Val2(Stringa)
    Else
        Errore
        GoTo uscita
    End If
    
    If Fattore1 = 0 Then
        fMain.MSComm1.Output = ScriviBattFact + vbCr
        fMain.MSComm1.Output = "1" + vbCr
        fMain.MSComm1.InBufferCount = 0
        Fattore1 = 1
    End If

    Lungo = TaraCanale(5) '18?
    
    VoltConvertitore = Lungo * 4095 / 5
    
    
'    Call Sleep(500)
'    fMain.MSComm1.InBufferCount = 0
'    fMain.MSComm1.Output = InfoAcq + vbCr
'    Stringa = InputComTimeOut(5)
'    Stringa = InputComTimeOut(5)
'
'    If Stringa <> "TimeOut" Then
'        VoltMisurati = Val2(Stringa)
'    Else
'        Errore
'        GoTo uscita
'    End If
    
    VoltConvertitore = VoltMisurati / Fattore1

    VoltEffettivi = Val2(InputBox("Immetti la tensione", "Taratura fattore batteria", VoltMisurati))
    
    Fattore1 = VoltEffettivi / VoltConvertitore
    
    fMain.MSComm1.Output = ScriviBattFact + vbCr
    fMain.MSComm1.Output = Trim(Str(Fattore1)) + vbCr
    fMain.MSComm1.InBufferCount = 0

uscita:
'    bTarapH.Enabled = True
'    bTaraT.Enabled = True
'    bTaraT2.Enabled = True
'    bTaraCond.Enabled = True
'    bFine.Enabled = True
'    bTaraBatt.Enabled = True
    AbilitaTasti


End Sub

Public Sub Shift(NewValue As Integer)
    Dim i As Integer
    For i = 0 To 49
        Y(i) = Y(i + 1)
    Next
    Y(50) = NewValue
End Sub

Public Sub DisabilitaTasti()
    bTarapH.Enabled = False
    bTaraT.Enabled = False
    'bTaraT2.Enabled = False
    bTaraCond.Enabled = False
    bFine.Enabled = False
    bTaraBatt.Enabled = False
End Sub

Public Sub AbilitaTasti()
    bTarapH.Enabled = True
    bTaraT.Enabled = True
    'bTaraT2.Enabled = True
    bTaraCond.Enabled = True
    bFine.Enabled = True
    bTaraBatt.Enabled = True
End Sub

Public Sub Errore()
    MsgBox ("Errore la centralina non risponde!")
    'Label2.Caption = ""
    Me.MousePointer = vbNormal

End Sub


Public Function TaraCanale(nCanale As Integer) As Integer
    Dim i As Long
    Dim Stringa As String
    Dim nDecimali As Integer
    Dim Lungo As Long
    Dim Misura As Single
    Dim UltimaTemp As Single
    Dim UltimoLungo As Long
    
    OpenCom
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = Chr$(3) 'CTRL+C
    Sleep (250)
    fMain.MSComm1.Output = TestSensori + vbCr

    'Azzera input buffer rs232
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.OutBufferCount = 0

'    For i = 0 To MaxCanali
'        If Canale(i).Attivo = True Then
            Me.Caption = "Taratura - Canale " + Canale(nCanale).Nome
'            Exit For
'        End If
'    Next
        'Manda la programmazione dei canali
    For i = 0 To MaxCanali
        If i = nCanale Then     'Se il canale è quello desiderato
            fMain.MSComm1.Output = "1" + vbCr
        Else
            fMain.MSComm1.Output = "0" + vbCr
        End If
     Next
    
    'Aspetta l'OK
    Stringa = InputComTimeOut(6)

    'Aggiorna il formato dei dati
    Formato = "0"
    'nDecimali = Val(frmOptions2.tDecimali.Text)
    nDecimali = 3
    If nDecimali < 0 Then nDecimali = 0
    If nDecimali > 7 Then nDecimali = 7
    If nDecimali <> 0 Then
        Formato = Formato + "."
        For i = 1 To nDecimali
            Formato = Formato + "0"
        Next
    End If
    bStop.Enabled = True
    
    continua = True
    Do
    
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
            'Aggiorna la relativa finestra se il canale è quello da monitorare
            If i = nCanale Then
                UltimoLungo = Lungo
                'Converte il valore dell'ADC nella misura corrispondente
                Misura = adc2value(Lungo, Canale(i).Bitmin, Canale(i).Bitmax, CDbl(Canale(i).valMax), CDbl(Canale(i).valMin), CDbl(Canale(i).valOff))
                'Misura = adc2value(Lungo, Canale(i).Bitmin, Canale(i).Bitmax, Canale(i).Valmax, Canale(i).Valmin, Canale(i).Valoff)
                'lCanale(i).Caption = Format(Misura, "0.0##")
                If i = 5 Then Misura = Misura * FattoreBatteriaInterna
                If i = 1 Then 'Si tratta ella conducibilità
                    If Misura <> 0 Then
                        'Non linearità
                        'Misura = 1 / Misura 'Gae 13mar2000
                    End If
                    'Misura = Misura / (1 + 0.0191 * (UltimaTemp - 25))
                End If
                If i = 0 Then UltimaTemp = Misura 'Conservo la temperatura
                lChannel.Caption = Format(Misura, Formato)
                'Modifica per visualizzare i count
                lCounts.Caption = Str(Lungo)
                Shift (Lungo)
                MSChart1.ChartData = Y

                'Debug.Print "Misura-->"; Misura
            End If
        Else
            Messaggio = "La centralina " + Versione + " non risponde!"
            MsgBox (Messaggio)
            'bFine_Click
            Exit Function
        End If
    Next

    
    Loop While continua = True
    Debug.Print continua
    Me.Caption = "Taratura"
    TaraCanale = UltimoLungo
End Function
