VERSION 5.00
Begin VB.Form fIntervallo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intervallo di acquisizione"
   ClientHeight    =   2235
   ClientLeft      =   1740
   ClientTop       =   2130
   ClientWidth     =   5310
   Icon            =   "fIntervallo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2235
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar sGiorni 
      Height          =   375
      Left            =   4320
      Max             =   60
      Min             =   -1
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.VScrollBar sOre 
      Height          =   375
      Left            =   4560
      Max             =   60
      Min             =   -1
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.VScrollBar sMinuti 
      Height          =   375
      Left            =   4800
      Max             =   60
      Min             =   -1
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.VScrollBar sSecondi 
      Height          =   375
      Left            =   5040
      Max             =   60
      Min             =   -1
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton bSecondip 
      Caption         =   "+"
      Height          =   255
      Left            =   3750
      TabIndex        =   13
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton bSecondim 
      Caption         =   "-"
      Height          =   255
      Left            =   3750
      TabIndex        =   14
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox tSecondi 
      Height          =   285
      Left            =   3390
      TabIndex        =   3
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton bContinua 
      Caption         =   "&Avvio >"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton bIndietro 
      Caption         =   "< &Indietro"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton bFine 
      Caption         =   "A&nnulla"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton PGiornop 
      Caption         =   "+"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton pGiornom 
      Caption         =   "-"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton bOrap 
      Caption         =   "+"
      Height          =   255
      Left            =   2190
      TabIndex        =   9
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton bOram 
      Caption         =   "-"
      Height          =   255
      Left            =   2190
      TabIndex        =   10
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton bMinp 
      Caption         =   "+"
      Height          =   255
      Left            =   2910
      TabIndex        =   11
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton bMinm 
      Caption         =   "-"
      Height          =   255
      Left            =   2910
      TabIndex        =   12
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox tGiorno 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox tOra 
      Height          =   285
      Left            =   1830
      TabIndex        =   1
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox tMin 
      Height          =   285
      Left            =   2550
      TabIndex        =   2
      Text            =   "1"
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lSecondi 
      Alignment       =   2  'Center
      Caption         =   "Secondi"
      Height          =   255
      Index           =   1
      Left            =   3270
      TabIndex        =   19
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Intervallo di acquisizione"
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
      Left            =   1200
      TabIndex        =   18
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lGiorni 
      Alignment       =   2  'Center
      Caption         =   "Giorni"
      Height          =   255
      Index           =   2
      Left            =   990
      TabIndex        =   17
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lOre 
      Alignment       =   2  'Center
      Caption         =   "Ore"
      Height          =   255
      Index           =   3
      Left            =   1710
      TabIndex        =   16
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lMinuti 
      Alignment       =   2  'Center
      Caption         =   "Minuti"
      Height          =   255
      Index           =   4
      Left            =   2550
      TabIndex        =   15
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "fIntervallo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim Secondi As Integer
    Dim Minuti As Integer
    Dim Ore As Integer
    Dim Giorni As Integer
    Dim Resto As Long
    
'    sGiorni.Visible = False
'    sOre.Visible = False
'    sMinuti.Visible = False
'    sSecondi.Visible = False


    If Intervallo <> 0 Then
    'Se l'intervallo è <>0 è stato scaricato dalla centralina
    'per cui lo copio nei controlli testo corrispondenti
        Giorni = Intervallo / 86400
        Resto = Intervallo Mod 86400
        Ore = Resto / 3600
        Resto = Resto Mod 3600
        Minuti = Resto / 60
        Secondi = Resto Mod 60
        tGiorno.Text = Trim(Str(Giorni))
        tOra.Text = Trim(Str(Ore))
        tMin.Text = Trim(Str(Minuti))
        tSecondi.Text = Trim(Str(Secondi))
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Me.Hide
        Unload Me
        fMain.Show
    End If
End Sub

Private Sub bContinua_Click()
    Dim Dummy As String
    Dim i As Integer
    Dim Risposta As String
    Dim Dummy2 As Boolean
    Dim Linea As String
    Dim Linea2 As String
    Dim Lungo As Long
    Dim TimeStop As Long
    
    Intervallo = Val(tGiorno.Text) * 86400 + Val(tOra.Text) * 3600 + Val(tMin.Text) * 60 + Val(tSecondi.Text)
    If Intervallo < MinimoIntervalloAcquisizione Then
        Messaggio = "Intervallo <" + Str(MinimoIntervalloAcquisizione) + " secondi, troppo piccolo"
        MsgBox (Messaggio)
        tSecondi.Text = Str(MinimoIntervalloAcquisizione)
        Exit Sub
    End If
    
    'Disattivazione pulsanti
    bContinua.Enabled = False
    bFine.Enabled = False
    bIndietro.Enabled = False
    bMinp.Enabled = False
    bMinm.Enabled = False
    bOrap.Enabled = False
    bOram.Enabled = False
    bSecondim.Enabled = False
    bSecondip.Enabled = False
    PGiornop.Enabled = False
    pGiornom.Enabled = False
    
    'Apre la comunicazione con la RS232
    OpenCom
    fMain.MSComm1.InBufferCount = 0
    'Imposta la lettura dal buffer ad un carattere alla volta
    fMain.MSComm1.InputLen = 1
    fMain.MSComm1.Output = Chr$(3)
    'Attende la risposta con timeout
    TimeStop = Timer + TmOut ' Imposta l'ora di fine
    'fMain.MSComm1.InBufferCount = 0
    Linea = ""
    Dummy = ""
    Do
        DoEvents
    Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
    If fMain.MSComm1.InBufferCount >= 1 Then
        ' Legge il dato di risposta  sulla porta
        ' seriale.
        TimeStop = Timer + TmOut ' Imposta l'ora di fine

        Do Until Dummy = vbLf Or (Timer > TimeStop)
            DoEvents
            Dummy = fMain.MSComm1.Input
            Linea = Linea + Dummy
        Loop

        'Controlla la presenza della stringa Simamet
        i = InStr(Linea, "Simamet")
        If i = 0 Then
            MsgBox ("La centralina Simamet non risponde al CTRL C --" + Linea)
            uscita
            Exit Sub
        End If
    Else
        MsgBox ("La centralina Simamet non risponde al CTRL C --")
        uscita
        Exit Sub
    End If

    
    'Manda Acquisizione (=1 vedi dichiarazione costanti)
     fMain.MSComm1.Output = Acquisizione + vbCr

    'Azzera input buffer rs232
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.OutBufferCount = 0
    
    'Manda il nome della stazione
    Linea = frmOptions.tStazione.Text
    If Len(Linea) > 20 Then Linea = Left$(Linea, 20)
    fMain.MSComm1.Output = Linea + vbCr
    'Open "log.txt" For Output As #12
    'Print #12, Linea
    'Close 12
    
    fMain.MSComm1.InBufferCount = 0
    
    
    'Manda l'ora del PC
    'Manda 1 (uso la costante Acquisizione perchè=1)
    fMain.MSComm1.Output = Acquisizione + vbCr
    
    'Manda data e ora attuale
    fMain.MSComm1.Output = Trim(Str$(Year(Now))) + vbCr
    fMain.MSComm1.Output = Trim(Str$(Month(Now))) + vbCr
    fMain.MSComm1.Output = Trim(Str$(Day(Now))) + vbCr
    fMain.MSComm1.Output = Trim(Str$(Hour(Now))) + vbCr
    fMain.MSComm1.Output = Trim(Str$(Minute(Now))) + vbCr
    fMain.MSComm1.Output = Trim(Str$(Second(Now))) + vbCr
    
    'Lieve ritardo
    Call Sleep(250)
   
    fMain.MSComm1.InBufferCount = 0
    'l'orario e' adesso oppure prefissato?
    If Orario = "NOW" Then
        fMain.MSComm1.Output = "2" + vbCr

    Else
        'Manda 1
        fMain.MSComm1.Output = "1" + vbCr
              
        'Manda data e ora di partenza programmata
        'NOTA Uso delle variabili globali e non direttamente i controlli perche' appena si esce dal
        'form orario i valori dei controlli text diventano quelli assegnati durante l'evento load
        'Ho scoperto il perchè: scaricavo il form!
        fMain.MSComm1.Output = PAnno + vbCr
        fMain.MSComm1.Output = PMese + vbCr
        fMain.MSComm1.Output = PGiorno + vbCr
        fMain.MSComm1.Output = POra + vbCr
        fMain.MSComm1.Output = PMinuti + vbCr
        fMain.MSComm1.Output = "0" + vbCr
        
        'Lieve ritardo
        Call Sleep(250)

        fMain.MSComm1.InBufferCount = 0
    End If
    
    'manda l'intervallo di campionamento in secondi
    Intervallo = Val(tGiorno.Text) * 86400 + Val(tOra.Text) * 3600 + Val(tMin.Text) * 60 + Val(tSecondi.Text)
    fMain.MSComm1.Output = Trim(Str$(Intervallo)) + vbCr
    fMain.MSComm1.InBufferCount = 0
    'Manda la programmazione dei canali
    For i = 0 To MaxCanali
        If Canale(i).Attivo = True Then
        'se il canale non è attivo
        'al posto del nome manda uno spazio
            If Canale(i).Nome = "" Then
                fMain.MSComm1.Output = "  " + vbCr
            Else
                fMain.MSComm1.Output = Canale(i).Nome + vbCr
            End If
        Else
                fMain.MSComm1.Output = "  " + vbCr
        End If
        'Se il canale è inattivo manda 0
        'se è attivo un altro numero
        If Canale(i).Attivo = True Then
            fMain.MSComm1.Output = "1" + vbCr
        Else
            fMain.MSComm1.Output = "0" + vbCr
        End If
        If Canale(i).Attivo = True Then
        'Se il canale non è attivo al posto dell'unità
        'di misura manda spazi
            If Canale(i).UnitaMisura = "" Then
                fMain.MSComm1.Output = "  " + vbCr
            Else
                fMain.MSComm1.Output = Canale(i).UnitaMisura + vbCr
            End If
        Else
                fMain.MSComm1.Output = "  " + vbCr
        End If
        fMain.MSComm1.Output = Trim(Str$(Canale(i).Bitmin)) + vbCr
        fMain.MSComm1.Output = Trim(Str$(Canale(i).Bitmax)) + vbCr
        fMain.MSComm1.Output = Trim(Str$(Canale(i).valMin)) + vbCr
        fMain.MSComm1.Output = Trim(Str$(Canale(i).valMax)) + vbCr
        fMain.MSComm1.Output = Trim(Str$(Canale(i).valOff)) + vbCr
    Next
    
    fMain.MSComm1.Output = Trim(Str$(mmxcount)) + vbCr
    fMain.MSComm1.Output = Trim(Str$(msxcount)) + vbCr
    'fMain.MSComm1.InBufferSize = 1024
    'fMain.MSComm1.InBufferCount = 0
    'Attende il PARTITO!
    'Linea = InputComTimeOut(5)
    TimeStop = Timer + TmOut ' Imposta l'ora di fine
    Linea = ""
    Dummy = ""
    Do
        DoEvents
    Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
    If fMain.MSComm1.InBufferCount >= 1 Then
        ' Legge il dato di risposta  sulla porta
        ' seriale.
        TimeStop = Timer + TmOut ' Imposta l'ora di fine

        Do Until Dummy = vbLf Or (Timer > TimeStop)
            DoEvents
            Dummy = fMain.MSComm1.Input
            Linea = Linea + Dummy
        Loop

        'MsgBox (Linea)
        i = InStr(Linea, "PARTITO!")
        If i = 0 Then
            Messaggio = "La centralina Simamet non risponde alla programmazione" + vbCrLf + Linea
            MsgBox (Messaggio)
            uscita
            Exit Sub
        End If
    Else
        MsgBox ("La centralina Simamet non risponde alla programmazione")
        uscita
        Exit Sub
    End If

    'Continua
    CloseCom
    'Manda un messaggio su acquisizione partita o programmazione effettuata
    MsgBox ("Centralina Simamet programmata!")
    
    
    'Esce!
    UnloadAllForms (Me.Name)
    Unload Me
    End

    
    
    'Riattivazione pulsanti
    bContinua.Enabled = True
    bFine.Enabled = True
    bIndietro.Enabled = True
    bMinp.Enabled = True
    bMinm.Enabled = True
    bOrap.Enabled = True
    bOram.Enabled = True
    bSecondim.Enabled = True
    bSecondip.Enabled = True
    PGiornop.Enabled = True
    pGiornom.Enabled = True
    
    
    
    Programmato = True
    
    'Disattivazione pulsanti in form principale
    fMain.bScarica.Enabled = False
    fMain.bProgramma.Enabled = False
    fMain.bTestSensori.Enabled = False
    fMain.bOrarioModem.Enabled = False
    fMain.bTaraBatt.Enabled = False
    'fMain.StatusBar1.Panels(2).Text = "Programmata"
    Unload Me
    fMain.Show
End Sub

Private Sub bFine_Click()
    Unload Me
    fMain.Show
End Sub

Private Sub bIndietro_Click()
    Me.Hide
    fPartenza.Show
End Sub

Private Sub bOram_Click()
    Dim Ora As Integer
    Ora = Val(tOra) - 1
    If Ora < 0 Then Ora = 23
    tOra = Ora
End Sub

Private Sub bOrap_Click()
    Dim Ora As Integer
    Ora = Val(tOra) + 1
    If Ora > 23 Then Ora = 0
    tOra = Ora
End Sub

Private Sub bSecondim_Click()
    Dim Secondi As Integer
    Secondi = Val(tSecondi) - 1
    If Secondi < 0 Then Secondi = 59
    tSecondi = Secondi
End Sub

Private Sub bSecondip_Click()
    Dim Secondi As Integer
    Secondi = Val(tSecondi) + 1
    If Secondi > 59 Then Secondi = 0
    tSecondi = Secondi
End Sub

Private Sub bMinp_Click()
    Dim Minuti As Integer
    Minuti = Val(tMin) + 1
    If Minuti > 59 Then Minuti = 0
    tMin = Minuti
End Sub

Private Sub bMinm_Click()
    Dim Minuti As Integer
    Minuti = Val(tMin) - 1
    If Minuti < 0 Then Minuti = 59
    tMin = Minuti
End Sub

Private Sub pGiornom_Click()
    Dim Giorno As Integer
    Giorno = Val(tGiorno) - 1
    If Giorno < 0 Then Giorno = 0
    tGiorno = Giorno
End Sub

Private Sub PGiornop_Click()
    Dim Giorno As Integer
    Giorno = Val(tGiorno) + 1
    tGiorno = Giorno
End Sub

Private Sub sGiorni_Change()
    Dim Giorni As Integer
    Giorni = sGiorni.value
    If Giorni < 0 Then
        Giorni = 0
        sGiorni.value = 0
    End If
    tGiorno = Giorni

End Sub

Private Sub sOre_Change()
    Dim Ore As Integer
    Ore = sOre.value
    If Ore < 0 Then
        Ore = 23
        sOre.value = 23
    End If
    If Ore = 24 Then
        Ore = 0
        sOre.value = 0
    End If
    Ore = 23 - Ore
    tOra = Ore

End Sub

Private Sub sMinuti_Change()
    Dim Minuti As Integer
    Minuti = sMinuti.value
    If Minuti < 0 Then
        Minuti = 59
        sMinuti.value = 59
    End If
    If Minuti = 60 Then
        Minuti = 0
        sMinuti.value = 0
    End If
    Minuti = 59 - Minuti
    tMin = Minuti

End Sub

Private Sub sSecondi_Change()
    Dim Secondi As Integer
    Secondi = sSecondi.value
    If Secondi < 0 Then
        Secondi = 59
        sSecondi.value = 59
    End If
    If Secondi = 60 Then
        Secondi = 0
        sSecondi.value = 0
    End If
    Secondi = 59 - Secondi
    tSecondi = Secondi
End Sub

Private Sub uscita()
    'Continua
    
    'Riattivazione pulsanti
    bContinua.Enabled = True
    bFine.Enabled = True
    bIndietro.Enabled = True
    Me.MousePointer = vbNormal
End Sub
