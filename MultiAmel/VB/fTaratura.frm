VERSION 5.00
Begin VB.Form fTaratura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tarature"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   ClipControls    =   0   'False
   Icon            =   "fTaratura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bTaraBatt 
      Caption         =   "Tara&Batt"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton bTaraCond 
      Caption         =   "Tara &Cd"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton bTaraT2 
      Caption         =   "Tara T&2"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton bTaraT 
      Caption         =   "Tara &T1"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton bFine 
      Caption         =   "&Esci"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton bTarapH 
      Caption         =   "Tara &pH"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "fTaratura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tarato As Boolean

Private Sub Form_Load()
    'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Text1.Text = ""
    Label2.Caption = ""
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
        'CloseCom
'        Me.Hide
'        Unload Me
'        fMain.Show
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

Private Sub bTarapH_Click()
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
'    bTarapH.Enabled = False
'    bTaraCond.Enabled = False
'    bTaraT.Enabled = False
'    bTaraT2.Enabled = False
'    bFine.Enabled = False
'     bTaraBatt.Enabled = False
     
    Text1.Text = ""
    Messaggio = "Metti il sensore in una" + vbCrLf
    Messaggio = Messaggio + "soluzione tampone a pH 7 e premi OK"
    Stringa = "Taratura sensore pH"
    Risposta = MsgBox(Messaggio, vbOKCancel, Stringa)
    If Risposta = vbCancel Then GoTo uscita
    'fMain.MSComm1.Output = CTRLC
    
'    fMain.MSComm1.Output = CTRLC
    Call Sleep(250)
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = TarapH + vbCr
    Call Sleep(250)
    Label2.Caption = "Attendere.."
    Me.MousePointer = vbHourglass
    'Questa linea non dovrebbe esserci ma siccome la risposta arriva
    'dopo almeno 20 secondi...
    'fMain.MSComm1.InBufferCount = 0
    Stringa = InputComTimeOut(40)
    If Stringa = "TimeOut" Then
        Errore
        GoTo uscita
    End If
    If Val(Stringa) = 0 Then 'Gae 13mar2000
        Errore
        GoTo uscita
    End If
    
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
    
    Debug.Print Stringa
    Label2.Caption = ""
    Me.MousePointer = vbNormal
    Intero = Val(Stringa)
    c1 = Intero
    pH1 = 7
    'decorrrezione temperatura
    'pH1 = pH1 * (273 + 21) / (273 + T7)
    
    Volt7 = 5 / 4095 * Intero
    Messaggio = "Volt a pH 7=" + Format(Volt7, "0.0#")
    Text1.Text = Messaggio
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
    'OpenCom
    'fMain.MSComm1.Output = CTRLC
    'fMain.MSComm1.Output = CTRLC
'    fMain.MSComm1.Output = CTRLC
    Call Sleep(250)
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = TarapH + vbCr
    Call Sleep(250)
    Label2.Caption = "Attendere.."
    Me.MousePointer = vbHourglass
    'fMain.MSComm1.InBufferCount = 0
    Stringa = InputComTimeOut(40)
    If Stringa = "TimeOut" Then
        Errore
        GoTo uscita
    End If
    
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
    'Debug.Print stringa
    Label2.Caption = ""
    Me.MousePointer = vbNormal
    Intero = Val(Stringa)
    c2 = Intero
    Volt7 = 5 / 4095 * Intero
    If Volt7 = Zero Then
        Messaggio = "ERRORE! La misura sembra" + vbCrLf
        Messaggio = Messaggio + "identica alla precedente!"
        MsgBox (Messaggio)
        GoTo daccapo
    End If
    Stringa = InputBox("Immetti il pH della soluzione tampone", "Taratura sensore pH", "4")
    If Stringa = "" Then
        Text1.Text = ""
        GoTo uscita
    End If
    If Val(Stringa) = 0 Then 'Gae 13mar2000
        Errore
        GoTo uscita
    End If
    
    pH = Val2(Stringa)
    'decorrrezione temperatura
    'pH = pH * (273 + 21) / (273 + T)
    pH2 = pH
  
    Messaggio = "Volt a pH" + Str(pH) + "=" + Format(Volt7, "0.0#")

    Text1.Text = Text1.Text + vbCrLf + Messaggio
    KpH = (pH - 7) / (Volt7 - Zero)
    Text1.Text = Text1.Text + vbCrLf + "KpH=" + Str(KpH)
    
    
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

    'salva i parametri!
    Intero = FreeFile
    Stringa = sGetAppPath + Versione + "Parametri.txt"
    Open Stringa For Output As #Intero
    Print #Intero, Zero
    Print #Intero, KpH
    Close Intero
    Label2.Caption = "Taratura effettuata"
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
    
    DisabilitaTasti
'    bTarapH.Enabled = False
'    bTaraT.Enabled = False
'    bTaraT2.Enabled = False
'    bTaraCond.Enabled = False
'    bFine.Enabled = False
'    bTaraBatt.Enabled = False
    
    Text1.Text = ""
    Stringa = "Taratura temperatura"
    Messaggio = "Metti il sensore di temperatura in" + vbCrLf
    Messaggio = Messaggio + "ambiente a temperatura" + vbCrLf
    Messaggio = Messaggio + "costante e aspetta una decina di minuti." + vbCrLf
    Messaggio = Messaggio + "Quando sei pronto premi OK"
    Risposta = MsgBox(Messaggio, vbOKCancel, Stringa)
    If Risposta = vbCancel Then GoTo uscita
    OpenCom
'    fMain.MSComm1.Output = CTRLC
    Call Sleep(250)
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = TaraTempE + vbCr
    Call Sleep(250)
    Label2.Caption = "Attendere.."
    Me.MousePointer = vbHourglass
    'fMain.MSComm1.InBufferCount = 0
    Stringa = InputComTimeOut(6)
    If Stringa = "TimeOut" Then
        Errore
        GoTo uscita
    End If
    If (Stringa) = "" Then 'Gae 13mar2000
        Errore
        GoTo uscita
    End If
    Debug.Print Stringa
    Label2.Caption = ""
    Me.MousePointer = vbNormal
    c1 = Val(Stringa)
    Text1.Text = "Misurati " + Str(c1) + " count"
    'Text1.Text = c1
    Stringa = InputBox("Immetti la temperatura", "Taratura sensore Temperatura", "15")
    If Stringa = "" Then
        Text1.Text = ""
        GoTo uscita
    End If

    t1 = Val2(Stringa)
    Text1.Text = Text1.Text + " a " + Str(t1) + "°C" + vbCrLf
daccapo:
    Messaggio = "Adesso metti il sensore" + vbCrLf
    Messaggio = Messaggio + "in un ambiente a temperatura" + vbCrLf
    Messaggio = Messaggio + "diversa e aspetta una decina di minuti." + vbCrLf
    Messaggio = Messaggio + "Quando sei pronto premi OK"
    Risposta = MsgBox(Messaggio, vbOKCancel, Stringa)
    If Risposta = vbCancel Then GoTo uscita
'    fMain.MSComm1.Output = CTRLC
    Call Sleep(250)
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = TaraTempE + vbCr
    Call Sleep(250)
    Label2.Caption = "Attendere.."
    Me.MousePointer = vbHourglass
    'fMain.MSComm1.InBufferCount = 0
    Stringa = InputComTimeOut(6)
    If Stringa = "TimeOut" Then
        Errore
        GoTo uscita
    End If
    
    Debug.Print Stringa
    Label2.Caption = ""
    Me.MousePointer = vbNormal
    c2 = Val(Stringa)
    If c2 = c1 Then
        Messaggio = "ERRORE! La misura sembra" + vbCrLf
        Messaggio = Messaggio + "identica alla precedente!"
        MsgBox (Messaggio)
        GoTo daccapo
    End If
    Text1.Text = Text1.Text + "Misurati " + Str(c2) + " count"
    Stringa = InputBox("Immetti la temperatura", "Taratura sensore Temperatura", "25")
    If Stringa = "" Then
        Text1.Text = ""
        GoTo uscita
    End If

    t2 = Val2(Stringa)
    Text1.Text = Text1.Text + " a " + Str(t2) + "°C" + vbCrLf
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

    Label2.Caption = "Taratura effettuata"
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
    
    DisabilitaTasti
'    bTarapH.Enabled = False
'    bTaraT.Enabled = False
'    bTaraT2.Enabled = False
'    bTaraCond.Enabled = False
'    bFine.Enabled = False
'    bTaraBatt.Enabled = False
    
    'controllare che il termometro sia già stato
    'tarato ed emettere un messaggio appropriato
    Text1.Text = ""
    c1 = 0
    cnd1 = 0
    Stringa = "Taratura conducibilità"
    Messaggio = "Metti il sensore di conducibilità in" + vbCrLf
    Messaggio = Messaggio + "liquido a conducibilità" + vbCrLf
    Messaggio = Messaggio + "nota." + vbCrLf
    Messaggio = Messaggio + "Quando sei pronto premi OK"
    Risposta = MsgBox(Messaggio, vbOKCancel, Stringa)
    If Risposta = vbCancel Then GoTo uscita
    'OpenCom
'    fMain.MSComm1.Output = CTRLC
'    fMain.MSComm1.Output = CTRLC
    Call Sleep(500)
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = TaraCond + vbCr
    'Call Sleep(250)
    Label2.Caption = "Attendere.."
    Me.MousePointer = vbHourglass
    'fMain.MSComm1.InBufferCount = 0
    Stringa = InputComTimeOut(30)
    If Stringa = "TimeOut" Then
        Errore
        GoTo uscita
    End If
    If Val(Stringa) = 0 Then 'Gae 13mar2000
        Errore
        GoTo uscita
    End If
    Debug.Print Stringa
    Label2.Caption = ""
    c2 = Val(Stringa)
    Text1.Text = "Misurati " + Str(c2) + " count"
    
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
        Text1.Text = ""
        GoTo uscita
    End If

    cnd2 = Val2(Stringa)
    'decorrezione per la temperatura
    
    'CondCorr(25 °C) = CondMis(T °C) / (1 + 0.0191 * (T - 25))
    
    'cnd2 = cnd2 * (1 + 0.0191 * (T - 25)) 'Cond. misurata alla temp. T
    
    Text1.Text = Text1.Text + " a " + Str(cnd2) + "mS" + vbCrLf
    
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

    Label2.Caption = "Taratura effettuata"
    Canale(1).UnitaMisura = "mS" 'Gae 13mar2000
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
    
    DisabilitaTasti
'    bTarapH.Enabled = False
'    bTaraT.Enabled = False
'    bTaraT2.Enabled = False
'    bTaraCond.Enabled = False
'    bFine.Enabled = False
'    bTaraBatt.Enabled = False

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

    Call Sleep(500)
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = InfoAcq + vbCr
    Stringa = InputComTimeOut(5)
    Stringa = InputComTimeOut(5)

    If Stringa <> "TimeOut" Then
        VoltMisurati = Val2(Stringa)
    Else
        Errore
        GoTo uscita
    End If
    
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

Public Sub Errore()
    MsgBox ("Errore la centralina non risponde!")
    Label2.Caption = ""
    Me.MousePointer = vbNormal

End Sub

Public Sub DisabilitaTasti()
    bTarapH.Enabled = False
    bTaraT.Enabled = False
    bTaraT2.Enabled = False
    bTaraCond.Enabled = False
    bFine.Enabled = False
    bTaraBatt.Enabled = False
End Sub

Public Sub AbilitaTasti()
    bTarapH.Enabled = True
    bTaraT.Enabled = True
    bTaraT2.Enabled = True
    bTaraCond.Enabled = True
    bFine.Enabled = True
    bTaraBatt.Enabled = True
End Sub
