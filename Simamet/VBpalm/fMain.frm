VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form fMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simamet"
   ClientHeight    =   3405
   ClientLeft      =   2790
   ClientTop       =   3120
   ClientWidth     =   5685
   ForeColor       =   &H00FFFFC0&
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bPalm 
      Height          =   570
      Left            =   2940
      Picture         =   "fMain.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cambia orario accensione e spegnimento del modem"
      Top             =   2520
      Width           =   570
   End
   Begin VB.CommandButton bTaraBatt 
      Caption         =   "Tara Batteria"
      Height          =   465
      Left            =   225
      TabIndex        =   7
      Top             =   1710
      Width           =   1095
   End
   Begin VB.CommandButton bOrarioModem 
      Height          =   570
      Left            =   4080
      Picture         =   "fMain.frx":1004
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cambia orario accensione e spegnimento del modem"
      Top             =   2520
      Width           =   570
   End
   Begin VB.CommandButton bRemota 
      Height          =   570
      Left            =   3510
      Picture         =   "fMain.frx":130E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Chiama con il modem"
      Top             =   2520
      Width           =   570
   End
   Begin VB.CommandButton bTestSensori 
      Height          =   570
      Left            =   1755
      Picture         =   "fMain.frx":1618
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Test sensori"
      Top             =   2520
      Width           =   570
   End
   Begin VB.CommandButton bProva 
      Caption         =   "Prova"
      Height          =   375
      Left            =   225
      TabIndex        =   8
      Top             =   135
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton bConnetti 
      Height          =   570
      Left            =   0
      Picture         =   "fMain.frx":1922
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Connetti"
      Top             =   2520
      Width           =   570
   End
   Begin VB.CommandButton bScarica 
      Height          =   570
      Left            =   585
      Picture         =   "fMain.frx":1C2C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Scarica i dati"
      Top             =   2520
      Width           =   570
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3150
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            Text            =   "Centralina N."
            TextSave        =   "Centralina N."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton bProgramma 
      Height          =   570
      Left            =   1170
      Picture         =   "fMain.frx":1F36
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Programma canali"
      Top             =   2520
      Width           =   570
   End
   Begin VB.CommandButton bFine 
      Height          =   570
      Left            =   4950
      Picture         =   "fMain.frx":2240
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Esci"
      Top             =   2520
      Width           =   570
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2565
      Top             =   1665
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComDlg.CommonDialog CmDialog1 
      Left            =   1800
      Top             =   1665
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Simamet"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   1575
      Left            =   270
      TabIndex        =   9
      Top             =   405
      Width           =   4950
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bProva_Click()

    fMain.MSComm1.CommPort = 2
    fMain.MSComm1.Settings = "19200,n,8,1"
    fMain.MSComm1.InBufferSize = 2048
    Collegato = True
    Me.MousePointer = vbNormal
    AbilitaTasti
    bFine.Enabled = True
    bTestSensori.Enabled = True
    OpenCom
    'Me.StatusBar1.Panels(1).Text = "Connesso"
End Sub

Private Sub Form_Load()
    Dim Ide As Boolean
    
    bScarica.Enabled = False
    bProgramma.Enabled = False
    bTestSensori.Enabled = False
    bOrarioModem.Enabled = False
    bTaraBatt.Enabled = False
    bTaraBatt.Visible = False
    bProva.Visible = False
    bRemota.Visible = False
    bOrarioModem.Visible = False
    
    
    Dim SaveTitle As String
    'Evita che venga lanciata un'ulteriore copia dell'applicazione
    If App.PrevInstance Then
        SaveTitle = App.Title
        App.Title = "... duplicate instance."      'Pretty, eh?
        fMain.Caption = "... duplicate instance."
        AppActivate SaveTitle
        SendKeys "% ~", True
        End
    End If
    
    'Disattivazione pulsanti per versione utente e non laboratorio
    If lDebug = False Then
        bProva.Visible = False
        'bProva2.Visible = False
        bTaraBatt.Visible = False
        bRemota.Visible = False
        bOrarioModem.Visible = False

    Else
        bProva.Visible = True
        'bProva2.Visible = True
        bTaraBatt.Visible = True
        bRemota.Visible = True
        bOrarioModem.Visible = True
        fMain.Caption = fMain.Caption + " Versione laboratorio"
    End If
    
    If SetInIDE() Then
        bProva.Visible = True
        'bProva2.Visible = True
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        UnloadAllForms (Me.Name)
        Unload Me
        End
    End If
End Sub

Private Sub Form_Paint()
    Gradient Me, 0, 255, 255, 0
End Sub

Private Sub Form_DblClick()
    frmAbout.Show 1
End Sub

Private Sub bInvia_Click()
    Dim i As Integer
    Dim bytel As Byte
    Dim byteh As Byte
    Dim nfile As Long

    If Collegato = True Then
 
    Else
        MsgBox ("Prima collegarsi ad una centralina Simamet")
    End If
End Sub

Private Sub bConnetti_Click()
Dim TimeStop As Long
Dim Linea As String
Dim Dummy As String
Dim Stringa As String
Dim Risposta As Long
Dim i As Long

ScegliCom:
    Me.Hide
    fCom.Show 1
    If ComPort = 0 Then Exit Sub
    OpenCom
    If ComOk = False Then GoTo ScegliCom
    DisabTasti
    bFine.Enabled = False
    bRemota.Enabled = False
    Me.MousePointer = vbHourglass
    DoEvents

    OpenCom

Riprova:
    Me.MSComm1.InBufferCount = 0
    Me.MSComm1.Output = Chr$(3)
    DoEvents
    Call Sleep(250)
    Me.MSComm1.Output = Chr$(3)
    DoEvents
    Call Sleep(250)
    Me.MSComm1.InBufferCount = 0
    Me.MSComm1.Output = Chr$(3)
    Call Sleep(500)
        
    'Attende la risposta con timeout
    TimeStop = Timer + TmOut ' Imposta l'ora di fine
    'I caratteri dalla RS232 vengono presi uno alla volta.
    Me.MSComm1.InputLen = 1
    Do
        DoEvents
    Loop Until (Me.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
    If Me.MSComm1.InBufferCount >= 1 Then
        ' Legge il dato di risposta  sulla porta
        ' seriale.
        TimeStop = Timer + TmOut
        Linea = ""
        Dummy = ""
        Do Until Dummy = vbLf Or (Timer > TimeStop)
            DoEvents
            Dummy = Me.MSComm1.Input
            Linea = Linea + Dummy
        Loop
            
        'controlla se nella risposta c'e' Simamet
        i = InStr(Linea, "Simamet")
        If i = 0 Then


           'Non c'e' ma il programma sul datalogger
           'potrebbe essere fermo. Controlla se c'e'
           'il prompt #
'           i = InStr(Linea, "#")
           'Potrebbe non esserci il prompt ma l'eco di vbCr+vbLF
'           If Linea = vbCr + vbLf Then i = 1
'           If i = 0 Then
'               'Non c'e', comunicazione errata.
'                GoTo Failed
               Timeout1
'               AbilitaTasti
               bFine.Enabled = True
               bConnetti.Enabled = True
               bRemota.Enabled = True
               Exit Sub
'           Else
               'C'e', facciamo ripartire il programma
'               Me.MSComm1.Output = Chr$(18)
'               Call Sleep(2000)
               'E controlla che il lancio sia avvenuto
'               GoTo Riprova
'           End If
        Else
            Collegato = True
        End If
             
    Else
        GoTo Failed
        Timeout1
        bFine.Enabled = True
        bConnetti.Enabled = True
        Exit Sub
    End If
                   
    Call Sleep(250)
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = InfoAcq + vbCr
    Stringa = ""
    Stringa = InputComTimeOut(5)
    If Stringa <> "TimeOut" Then
        Me.StatusBar1.Panels(2).Text = Left(Stringa, Len(Stringa) - 2) & " Bytes"
    Else
        GoTo Failed
    End If
    
    Stringa = InputComTimeOut(5)
    If Stringa <> "TimeOut" Then
        Me.StatusBar1.Panels(1).Text = Left(Stringa, Len(Stringa) - 2) & " volt"
        TensioneBatteria = Val2(Stringa)
    Else
        GoTo Failed
    End If

    Stringa = InputComTimeOut(5)
    If Stringa <> "TimeOut" Then
        Me.StatusBar1.Panels(3).Text = "Centralina N. " + Left(Stringa, Len(Stringa) - 2)
    Else
        GoTo Failed
    End If

     'Legge il Fattore Batteria
    Call Sleep(250)
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = LeggiBattFact + vbCr
    Stringa = ""
    Stringa = InputComTimeOut(5)
    If Stringa <> "TimeOut" Then
        FattoreBatteriaInterna = Val2(Stringa)
    Else
        GoTo Failed
    End If


    'Risposta = ScaricaProgrammazione
    If Risposta <> 0 Then ProgrammazioneCaricata = True
 
    Me.MSComm1.InBufferCount = 0
    Me.MousePointer = vbNormal
    AbilitaTasti
    bFine.Enabled = True
    bRemota.Enabled = False
    'Me.StatusBar1.Panels(1).Text = "Connesso"
    Me.MSComm1.InBufferCount = 0
    Exit Sub

Failed:
        Timeout1
        bFine.Enabled = True
        bConnetti.Enabled = True
        bRemota.Enabled = True
        CloseCom
    
End Sub

Private Sub bRemota_Click()
    Me.Hide
    fModem.Show
End Sub

Private Sub bFine_Click()
    Dim Stile As Long
    Dim Risposta As Long
    Dim Titolo As String

    If Programmato = False And Collegato = True Then
        ' Definisce messaggio.
        Messaggio = "Il datalogger non è stato programmato." + vbCrLf
        Messaggio = Messaggio + "Si vuole uscire ugualmente ?" + vbCrLf
        Stile = vbYesNo + vbCritical + vbDefaultButton2 ' Definisce pulsanti.
        Titolo = "ATTENZIONE!"  ' Definisce titolo.
        Risposta = MsgBox(Messaggio, Stile, Titolo)
        If Risposta = vbNo Then
            Exit Sub
        Else
            OpenCom
            MSComm1.Output = CTRLC
            Sleep (25)
            MSComm1.Output = Dormi + vbCr
        End If
    End If

    UnloadAllForms (Me.Name)
    Unload Me
    End
End Sub

Private Sub bProgramma_Click()
    Dim Stile As Long
    Dim Risposta As Long
    Dim Titolo As String

    DisabTasti

    If Scaricato = False Then
        'msgbox "non hai scaricato!" Continuo?
        Messaggio = "I dati eventualmente raccolti verranno cancellati!"
        Messaggio = Messaggio + vbCr + vbCr + " Continuo?"
        
        Stile = vbYesNo + vbCritical + vbDefaultButton2 ' Definisce pulsanti.
        Titolo = "ATTENZIONE!"  ' Definisce titolo.
        Risposta = MsgBox(Messaggio, Stile, Titolo)
        'Adesso si controlla la risposta
        If Risposta = vbYes Then   ' L'utente sceglie il
                                   ' pulsante Sì.
                CancellaFlash
                AbilitaTasti
                Me.Hide
                frmOptions.Show
                Exit Sub
        Else    ' L'utente sceglie il
                ' pulsante No o annulla.
            AbilitaTasti
            Exit Sub
        End If
    Else
        'msgbox "non hai scaricato!" Continuo?
        Messaggio = "Cancello i dati raccolti dalla centralina Simamet?" ' Definisce messaggio.
        Stile = vbYesNo + vbCritical + vbDefaultButton2 ' Definisce pulsanti.
        Titolo = "ATTENZIONE!"  ' Definisce titolo.
        Risposta = MsgBox(Messaggio, Stile, Titolo)
        'Adesso si controlla la risposta
        If Risposta = vbYes Then   ' L'utente sceglie il
                                   ' pulsante Sì.
            CancellaFlash
            AbilitaTasti
            Me.Hide
            frmOptions.Show
            Exit Sub
        Else    ' L'utente sceglie il
                ' pulsante No o annulla.
            AbilitaTasti
            Exit Sub
        End If
    End If
    
End Sub

Private Sub bScarica_Click()
    Dim Linea As String     'Variabile dove registro ogni linea di dati ricevuta
    Dim MioFile As String
    Dim Dummy As String
    Dim Blocco() As Byte
    Dim Buffer As Variant
    
    'Controlla che non sia stato gia' programmato
    'Lanciato ****************
        
    Palm = False
        
    'impostazioni iniziali di CmDialog1
    NewPath sGetAppPath

    On Error GoTo Annulla
    CmDialog1.CancelError = True
    'Controlla se si vuole sostituire il file,
    'che la directory eventualmente immessa esista,
    'non prende in considerazione files e directory a sola lettura
    'non mostra la casella sola lettura
    CmDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNHideReadOnly
    'Filtri di dialogo
    CmDialog1.Filter = "File Ascii (*.dat)|*.dat|File Sima (*.sim)|*.sim|Tutti i file (*.*)|*.*"
    Dummy = sGetAppPath()
    Dummy = Dummy + Trim(Stazione)
    Dummy = Dummy + Format(Year(Now), "0000")
    Dummy = Dummy + Format(Month(Now), "00")
    Dummy = Dummy + Format(Day(Now), "00")
    Dummy = Dummy + Format(Hour(Now), "00")
    Dummy = Dummy + Format(Minute(Now), "00")
    Dummy = Dummy + Format(Second(Now), "00")
    Dummy = Dummy + ".dat"
    fMain.CmDialog1.FileName = Dummy
'    fMain.CmDialog1.filename = ""
    CmDialog1.ShowSave
    FileOut = CmDialog1.FileName
    Dummy = LCase(Right(FileOut, 4))
    Select Case CmDialog1.FilterIndex
'        Case Xls
'            If Dummy <> ".xls" Then FileOut = FileOut + ".xls"
'            TipoFile = "EXCEL2"
        Case 2
            If Dummy <> ".sim" Then FileOut = FileOut + ".sim"
            TipoFile = "SIMAPRO2"
        Case 1
            If Dummy <> ".dat" Then FileOut = FileOut + ".dat"
            TipoFile = "ASCII-DAT"
        Case Else
            FileOut = FileOut + ".dat"
            TipoFile = "ASCII-DAT"
    End Select
    
    DoEvents
    
    Me.MousePointer = vbHourglass
    DisabTasti
    FileName = FileOut
    fCounter.Show
    fCounter.Scarica
    Exit Sub
Annulla:
    Me.MousePointer = vbDefault
    AbilitaTasti
    'Imposta la lettura del buffer a tutto il buffer alla volta
    Me.MSComm1.InputLen = 0
    DoEvents
End Sub

Private Sub bPalm_Click()
    Dim Dummy As String

ScegliCom:
    'Fa comparire il dialogo scelta porta COM
    Me.Hide
    fCom.Show 1
    If ComPort = 0 Then Exit Sub
    OpenCom
    If ComOk = False Then GoTo ScegliCom
    DisabTasti
    bFine.Enabled = False
    bRemota.Enabled = False
    Me.MousePointer = vbHourglass
    DoEvents
    'Apre la porta COM
    OpenCom

    'Fa apparire la Common Dialog per il nome del file
    'impostazioni iniziali di CmDialog1
    NewPath sGetAppPath

    On Error GoTo Annulla
    CmDialog1.CancelError = True
    'Controlla se si vuole sostituire il file,
    'che la directory eventualmente immessa esista,
    'non prende in considerazione files e directory a sola lettura
    'non mostra la casella sola lettura
    CmDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNHideReadOnly
    'Filtri di dialogo
    CmDialog1.Filter = "File Ascii (*.dat)|*.dat|File Sima (*.sim)|*.sim|Tutti i file (*.*)|*.*"
    Dummy = sGetAppPath()
    Dummy = Dummy + Trim(Stazione)
    Dummy = Dummy + Format(Year(Now), "0000")
    Dummy = Dummy + Format(Month(Now), "00")
    Dummy = Dummy + Format(Day(Now), "00")
    Dummy = Dummy + Format(Hour(Now), "00")
    Dummy = Dummy + Format(Minute(Now), "00")
    Dummy = Dummy + Format(Second(Now), "00")
    'Dummy = Dummy + ".dat"
    fMain.CmDialog1.FileName = Dummy
'    fMain.CmDialog1.filename = ""
    CmDialog1.ShowSave
    FileOut = CmDialog1.FileName
    Dummy = LCase(Right(FileOut, 4))
    Select Case CmDialog1.FilterIndex
'        Case Xls
'            If Dummy <> ".xls" Then FileOut = FileOut + ".xls"
'            TipoFile = "EXCEL2"
        Case 2
            If Dummy <> ".sim" Then FileOut = FileOut + ".sim"
            TipoFile = "SIMAPRO2"
        Case 1
            If Dummy <> ".dat" Then FileOut = FileOut + ".dat"
            TipoFile = "ASCII-DAT"
        Case Else
            FileOut = FileOut + ".dat"
            TipoFile = "ASCII-DAT"
    End Select
    
    DoEvents

    bFine.Enabled = True
    bRemota.Enabled = True
    Palm = True
    
        DisabTasti
    FileName = FileOut
    fCounter.Show
    fCounter.Scarica
    Exit Sub
Annulla:
    Me.MousePointer = vbDefault
    AbilitaTasti
    'Imposta la lettura del buffer a tutto il buffer alla volta
    Me.MSComm1.InputLen = 0
    DoEvents

End Sub


Private Sub bTestSensori_Click()
    Me.Hide
    frmOptions2.Show
End Sub

Private Sub bOrarioModem_Click()
    Me.Hide
    fOrarioModem.Show
    Exit Sub
End Sub
Private Sub bTaraBatt_Click()
    Dim Fattore1 As Single
    Dim VoltMisurati As Single
    Dim VoltEffettivi As Single
    Dim VoltConvertitore As Single
    Dim Fatt As Single
    Dim Stringa As String
    
    'DisabilitaTasti
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
    'AbilitaTasti

End Sub


Private Sub Timeout1()
    'Prova a far ripartire il programma
    Dim Mes As String
    Me.MSComm1.Output = Chr$(18)
    Mes = "         Errore nella comunicazione" + vbCr + "     la stazione Simamet non risponde!" + vbCr + "Controllare che sia in modo Comandi" + vbCr + "  Controllare il cavo di collegamento"
    MsgBox (Mes)
    UnloadAllForms (Me.Name)
    Me.MousePointer = vbNormal
    Me.Show
    'Me.StatusBar1.Panels(3).Text = "Errore nella comunicazione"
End Sub

Public Sub AbilitaTasti()
    'Abilita i tasti del form principale
    bScarica.Enabled = True
    bProgramma.Enabled = True
    bConnetti.Enabled = True
    bTestSensori.Enabled = True
    bOrarioModem.Enabled = True
    bTaraBatt.Enabled = True
End Sub

Public Sub DisabTasti()
    'Disabilita i tasti del form principale
    bScarica.Enabled = False
    bProgramma.Enabled = False
    bConnetti.Enabled = False
    bTestSensori.Enabled = False
    bOrarioModem.Enabled = False
    bTaraBatt.Enabled = False
End Sub

Public Sub CancellaFlash()
        'Cancella la memoria Flash
        OpenCom
        Me.MSComm1.Output = Chr$(3)
        Call Sleep(250)
        Me.MSComm1.Output = StopPrg + vbCr
        Call Sleep(250)
        Me.MSComm1.Output = Chr$(5)
        Call Sleep(20)
        Me.MSComm1.Output = Chr$(250)
        Call Sleep(500)
        Me.MSComm1.Output = Chr$(18)  'CTRL+R
        Call Sleep(500)
        'Azzera input buffer rs232
        Me.MSComm1.InBufferCount = 0
End Sub

Public Sub Errore()
    MsgBox ("Errore la centralina non risponde!")
    'Label2.Caption = ""
    Me.MousePointer = vbNormal

End Sub
