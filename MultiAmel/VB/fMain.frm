VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form fMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programmazione centralina Versione"
   ClientHeight    =   2730
   ClientLeft      =   2790
   ClientTop       =   3120
   ClientWidth     =   5310
   ForeColor       =   &H00FFFFC0&
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bAscii 
      Caption         =   "Send ASCII"
      Height          =   375
      Left            =   4080
      TabIndex        =   27
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton bzero 
      Caption         =   "0"
      Height          =   495
      Left            =   4080
      TabIndex        =   26
      ToolTipText     =   "Send ascii 0"
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton bTerminal 
      Height          =   495
      Left            =   4080
      Picture         =   "fMain.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Terminale"
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton bCTRLC 
      Caption         =   "CTRL + C"
      Height          =   495
      Left            =   4560
      TabIndex        =   24
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton bExpert 
      Height          =   495
      Left            =   0
      Picture         =   "fMain.frx":053A
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Expert mode"
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton bCTRLR 
      Caption         =   "CTRL + R"
      Height          =   495
      Left            =   4560
      TabIndex        =   22
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton bLastFile 
      Height          =   495
      Left            =   495
      Picture         =   "fMain.frx":0844
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Visualizza l'ultimo file"
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton bOrarioModem 
      Height          =   495
      Left            =   3810
      Picture         =   "fMain.frx":0B4E
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Orario Modem"
      Top             =   1950
      Width           =   495
   End
   Begin VB.CommandButton bRemota 
      Height          =   495
      Left            =   3315
      Picture         =   "fMain.frx":0E58
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Modem"
      Top             =   1950
      Width           =   495
   End
   Begin VB.CommandButton bTara2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1980
      Picture         =   "fMain.frx":1162
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Taratura 2"
      Top             =   1950
      Width           =   495
   End
   Begin VB.CommandButton bTaratura 
      Height          =   495
      Left            =   4080
      Picture         =   "fMain.frx":146C
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Taratura"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton bDormi 
      Height          =   495
      Left            =   4305
      Picture         =   "fMain.frx":18AE
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Dormi"
      Top             =   1950
      Width           =   495
   End
   Begin VB.CommandButton bTestSensori 
      Height          =   495
      Left            =   1485
      Picture         =   "fMain.frx":1BB8
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Test Sensori"
      Top             =   1950
      Width           =   495
   End
   Begin VB.CommandButton bProva 
      Caption         =   "Abilita tasti"
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton bConnetti 
      Height          =   495
      Left            =   0
      Picture         =   "fMain.frx":1EC2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Connetti"
      Top             =   1950
      Width           =   495
   End
   Begin VB.CommandButton bScarica 
      Height          =   495
      Left            =   495
      Picture         =   "fMain.frx":21CC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Scarico dati"
      Top             =   1950
      Width           =   495
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   2445
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2117
            MinWidth        =   2117
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton bProgramma 
      Height          =   495
      Left            =   990
      Picture         =   "fMain.frx":24D6
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Programma canali"
      Top             =   1950
      Width           =   495
   End
   Begin VB.CommandButton bFine 
      Height          =   495
      Left            =   4800
      Picture         =   "fMain.frx":27E0
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Esci"
      Top             =   1950
      Width           =   495
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   19200
   End
   Begin MSComDlg.CommonDialog CmDialog1 
      Left            =   600
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameCentralina 
      Caption         =   "Centralina N."
      Height          =   1695
      Left            =   2160
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   2775
      Begin VB.Label Label5 
         Caption         =   "Bytes"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Volts"
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lBytes 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lVolts 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Memoria occupata"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Tensione batteria"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "fMain.frx":2C2A
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lAttendere 
      BackStyle       =   0  'Transparent
      Caption         =   "Attendere...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2040
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MultiAmel"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4665
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub bProva_Click()
'Ok = GetDiskFreeSpace("c:", SectorsPerCluster, _
BytesPerSector, NumberOfFreeClusters, _
TtoalNumberOfClusters)
'Bytes = NumberOfFreeClusters * SectorsPerCluster * BytesPerSector
'Clusters = NumberOfFreeClusters * SectorsPerCluster

    'fMain.MSComm1.CommPort = 3
    'fMain.MSComm1.Settings = "19200,n,8,1"
    'fMain.MSComm1.InBufferSize = 2048
    Collegato = True
    Me.MousePointer = vbNormal
    AbilitaTasti
    bFine.Enabled = True
    bTestSensori.Enabled = True
    bDormi.Enabled = True
    OpenCom
    'Me.StatusBar1.Panels(1).Text = "Connesso"

End Sub



Private Sub Form_Load()
    'bProva.Visible = False
    'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    bScarica.Enabled = False
    bProgramma.Enabled = False
    bTestSensori.Enabled = False
    bDormi.Enabled = False
    bTaratura.Enabled = False
    bOrarioModem.Enabled = False
    bLastFile.Enabled = False
    bCTRLR.Enabled = False
    
    isTerminal = False
    
    
    
    
'    If Versione = "MultiAmel" Then
'        On Error Resume Next
'        Messaggio = sGetAppPath + "connessionecavo.ico"
'        bConnetti.Picture = LoadPicture(Messaggio)
'        Messaggio = sGetAppPath + "programma.ico"
'        bProgramma.Picture = LoadPicture(Messaggio)
'        Messaggio = sGetAppPath + "Scarica32x32.ico"
'        bScarica.Picture = LoadPicture(Messaggio)
''        Messaggio = sGetAppPath + "Scarica32x32.ico"
''        bTestSensori.Picture = LoadPicture(Messaggio)
'        Messaggio = sGetAppPath + "modem.ico"
'        bRemota.Picture = LoadPicture(Messaggio)
'        Messaggio = sGetAppPath + "OrarioModem32x32.ico"
'        bOrarioModem.Picture = LoadPicture(Messaggio)
'        Messaggio = sGetAppPath + "dormi.ico"
'        bDormi.Picture = LoadPicture(Messaggio)
'        Messaggio = sGetAppPath + "esci3.ico"
'        bFine.Picture = LoadPicture(Messaggio)
'        On Error GoTo 0
'    End If
    
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
    
    
    'Label1.Caption = Versione
    fMain.Caption = "Programmazione centralina Multipar" ' + Versione
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

Private Sub bCTRLC_Click()
    If Me.MSComm1.PortOpen = True Then
        Me.MSComm1.Output = CTRLC
    End If
End Sub

Private Sub bCTRLR_Click()
    If Me.MSComm1.PortOpen = True Then
        Me.MSComm1.Output = Chr$(18)
    End If
End Sub

Private Sub bzero_Click()
    If Me.MSComm1.PortOpen = True Then
        Me.MSComm1.Output = Chr$(0)
    End If

End Sub

Private Sub bAscii_Click()
    Dim i As Integer
    Dim prompt As String
    prompt = "Immettere il codice del carattere ASCII da inviare"
    prompt = prompt + vbCrLf + "Esempi"
    prompt = prompt + vbCrLf + "CTRL+C = 003"
    prompt = prompt + vbCrLf + "CTRL+E = 005"
    prompt = prompt + vbCrLf + "CTRL+R = 018"
    prompt = prompt + vbCrLf + "Can250 = 250"
    prompt = prompt + vbCrLf + "cancella memoria 5 + 250"
    i = Val(InputBox(prompt, ""))
    If Me.MSComm1.PortOpen = True Then
        Me.MSComm1.Output = Chr$(i)
    End If
End Sub

Private Sub bDormi_Click()
    OpenCom
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = Dormi + vbCr
    bProgramma.Enabled = False
    bScarica.Enabled = False
    bTestSensori.Enabled = False
    bTaratura.Enabled = False
    bOrarioModem.Enabled = False
    bCTRLR.Enabled = False

End Sub

Private Sub bExpert_Click()
    Static eEnabled As Boolean
    
    If eEnabled = False Then
    
        bProva.Visible = True
        bProva.Enabled = True
        bCTRLR.Visible = True
        bCTRLR.Enabled = True
        bCTRLC.Visible = True
        bCTRLC.Enabled = True
        bTaratura.Visible = True
        bTerminal.Visible = True
        bzero.Visible = True
        bAscii.Visible = True
        eEnabled = True
    Else
        bProva.Visible = False
        bProva.Enabled = False
        bCTRLR.Visible = False
        bCTRLR.Enabled = False
        bCTRLC.Visible = False
        bCTRLC.Enabled = False
        bTaratura.Visible = False
        bTerminal.Visible = False
        bzero.Visible = False
        bAscii.Visible = False
        eEnabled = False
    End If
End Sub

Private Sub bTara2_Click()
'    lDebug = Not lDebug
'    Me.Hide
'    fOrarioModem.Show
'    Exit Sub
    'DisabTasti
    Load fTara
    Me.Hide
    fTara.Show
    Exit Sub
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
    'Call Sleep(500)
    
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
            
        'controlla se nella risposta c'e' MultiAmel o versione
        i = InStr(Linea, Versione)
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
                   
    i = WriteINI("Cavo", "UltimaCom", ComPort, FileIni)
                   
                   
    Call Sleep(250)
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = InfoAcq + vbCr
    Stringa = ""
    Stringa = InputComTimeOut(5)
    If Stringa <> "TimeOut" Then
        lBytes.Caption = Left(Stringa, Len(Stringa) - 2)
        Me.StatusBar1.Panels(2).Text = Left(Stringa, Len(Stringa) - 2) & " Bytes"
    Else
        GoTo Failed
    End If
    
    Stringa = InputComTimeOut(5)
    If Stringa <> "TimeOut" Then
        lVolts.Caption = Left(Stringa, Len(Stringa) - 2)
        Me.StatusBar1.Panels(1).Text = Left(Stringa, Len(Stringa) - 2) & " volt"
        TensioneBatteria = Val2(Stringa)
    Else
        GoTo Failed
    End If

    Stringa = InputComTimeOut(5)
    If Stringa <> "TimeOut" Then
        frameCentralina.Caption = "Centralina N. " + Left(Stringa, Len(Stringa) - 2)
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
    'Controllo che la tensione batteria sia a posto
    If TensioneBatteria < MinimaTensioneBatteria Then
    
        Messaggio = "La tensione della batteria è troppo bassa."
        Messaggio = Messaggio + vbCr + "Il campionamento potrà essere pocco accurato o"
        Messaggio = Messaggio + vbCr + "la centralina potrà bloccarsi e andare in basso consumo"
        Messaggio = Messaggio + vbCr + "in attesa che venga cambiata la batteria."
        Messaggio = Messaggio + vbCr + vbCr + " Continuo?"
        Stile = vbYesNo + vbCritical + vbDefaultButton2 ' Definisce pulsanti.
        Titolo = "ATTENZIONE!"  ' Definisce titolo.
        Risposta = MsgBox(Messaggio, Stile, Titolo)
        If Risposta = vbNo Then
            AbilitaTasti
            Exit Sub

        End If
    End If
    
    DoEvents
    'Set frmOptions.tbsOptions.SelectedItem = frmOptions.tbsOptions.Tabs(1)

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
                lAttendere.Visible = True
                DoEvents
                CancellaFlash
                lAttendere.Visible = False
                DoEvents
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
        Messaggio = "Cancello i dati raccolti dalla centralina " + Versione + "?" ' Definisce messaggio.
        Stile = vbYesNo + vbCritical + vbDefaultButton2 ' Definisce pulsanti.
        Titolo = "ATTENZIONE!"  ' Definisce titolo.
        Risposta = MsgBox(Messaggio, Stile, Titolo)
        'Adesso si controlla la risposta
        If Risposta = vbYes Then   ' L'utente sceglie il
                                   ' pulsante Sì.
            lAttendere.Visible = True
            DoEvents
            CancellaFlash
            lAttendere.Visible = False
            DoEvents
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
        
        
    'impostazioni iniziali di CmDialog1
    NewPath sGetAppPath

    CmDialog1.CancelError = True
    'Controlla se si vuole sostituire il file,
    'che la directory eventualmente immessa esista,
    'non prende in considerazione files e directory a sola lettura
    'non mostra la casella sola lettura
    CmDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNHideReadOnly
    'Filtri di dialogo
    CmDialog1.Filter = "File Ascii (*.dat)|*.dat|File Sima (*.sim)|*.sim|Tutti i file (*.*)|*.*"
    Dummy = sGetAppPath()
    OpenCom
    MSComm1.InBufferCount = 0
    MSComm1.Output = GetStationName + vbCrLf
    Stazione = InputComTimeOut(5)
    fMain.MSComm1.Output = CTRLC
    Stazione = Left$(Stazione, 20)
    Stazione = Trim(Stazione)
    If Stazione = "" Or Stazione = "TimeOut" Or Stazione = vbCrLf Then Stazione = "MultiAmel"
    Dummy = Dummy + Stazione
    Dummy = Dummy + Format(Year(Now), "0000")
    Dummy = Dummy + Format(Month(Now), "00")
    Dummy = Dummy + Format(Day(Now), "00")
    Dummy = Dummy + Format(Hour(Now), "00")
    Dummy = Dummy + Format(Minute(Now), "00")
    Dummy = Dummy + Format(Second(Now), "00")
    Dummy = Dummy + ".dat"
    fMain.CmDialog1.FileName = Dummy
    'CmDialog1.InitDir = InitDirData
    On Error GoTo Annulla
    CmDialog1.ShowSave
    
    FileOut = CmDialog1.FileName
    Dummy = LCase(Right(FileOut, 4))
    Select Case CmDialog1.FilterIndex
        Case 2
            If Dummy <> ".sim" Then FileOut = FileOut + ".sim"
        Case 1
            If Dummy <> ".dat" Then FileOut = FileOut + ".dat"
    End Select
    DoEvents
    
    Me.MousePointer = vbHourglass
    DisabTasti
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

Private Sub bTaratura_Click()
    Me.Hide
    fTaratura.Show
End Sub

Private Sub bOrarioModem_Click()
    Me.Hide
    fOrarioModem.Show
    Exit Sub
End Sub

Private Sub bTerminal_Click()
    Me.Hide
    frmTerminal.Show
End Sub

Private Sub bLastFile_Click()
    Dim Stringa As String
    'LastFileSaved = "C:\SCANLOG.TXT"
    If LastFileSaved = "" Then Exit Sub
    If Dir$(LastFileSaved) = "" Then Exit Sub
    Stringa = "notepad " + LastFileSaved
    Shell Stringa, vbNormalFocus
End Sub

Private Sub Timeout1()
    'Prova a far ripartire il programma
    Dim Mes As String
    Me.MSComm1.Output = Chr$(18)
    Mes = "         Errore nella comunicazione" + vbCr + "     la stazione " + Versione + " non risponde!" + vbCr + "Controllare che sia in modo Comandi" + vbCr + "  Controllare il cavo di collegamento"
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
    bDormi.Enabled = True
    bTaratura.Enabled = True
    bOrarioModem.Enabled = True
    bRemota.Enabled = True
    bLastFile.Enabled = True
    bCTRLR.Enabled = True
    bTerminal.Enabled = True
    bTara2.Enabled = True
    DoEvents
End Sub

Public Sub DisabTasti()
    'Disabilita i tasti del form principale
    bScarica.Enabled = False
    bProgramma.Enabled = False
    bConnetti.Enabled = False
    bTestSensori.Enabled = False
    bDormi.Enabled = False
    bTaratura.Enabled = False
    bRemota.Enabled = False
    bOrarioModem.Enabled = False
    bLastFile.Enabled = False
    bCTRLR.Enabled = False
    bTerminal.Enabled = False
    bTara2.Enabled = False
    DoEvents
End Sub

Public Function CancellaFlash() As Boolean
        'Cancella la memoria Flash
        Dim Risposta As String
        Dim Stringa As String
        Dim Intero As Integer
        
        CancellaFlash = False
        OpenCom
        Me.MSComm1.InBufferCount = 0
        Me.MSComm1.Output = Chr$(3)
        Risposta = InputComTimeOut(5)
        Debug.Print Risposta
        If Left(Risposta, Len(Risposta) - 2) <> Versione Then
            Stringa = "ERRORE! La centralina non risponde! (CancellaFlash01) -->" + Risposta
            MsgBox Stringa, vbOKOnly
            ScriviErroreSuLog Stringa
            Exit Function
        End If

        'ScriviErroreSuLog "CHR$3-->" + Risposta

        Me.MSComm1.Output = StopPrg + vbCr
        Risposta = InputComTimeOut(5)   'vbCrLf
        If Risposta <> vbCrLf Then
            Stringa = "ERRORE! La centralina non risponde! (CancellaFlash02) -->" + Char2ascii(Risposta)
            MsgBox Stringa, vbOKOnly
            ScriviErroreSuLog Stringa
            Exit Function
        End If

        'ScriviErroreSuLog "StopProg-->" + Risposta

        Risposta = InputComTimeOutBin(5, 1)   '# senza vbCrLf
        If Risposta <> "#" Then
            Stringa = "ERRORE! La centralina non risponde! (CancellaFlash03) -->" + Risposta
            MsgBox Stringa, vbOKOnly
            ScriviErroreSuLog Stringa
            Exit Function
        End If
        'ScriviErroreSuLog Risposta
        'Sleep 1000
        Me.MSComm1.Output = Chr$(5)
        Risposta = InputComTimeOutBin(5, 1)
        If Risposta <> Chr$(5) Then
            Stringa = "ERRORE! La centralina non risponde! (CancellaFlash04) -->" + Char2ascii(Risposta)
            MsgBox Stringa, vbOKOnly
            Exit Function
        End If
        Me.MSComm1.Output = Chr$(250)
        Risposta = InputComTimeOut(5)   '1+vbCrLf=ok 0+vbCrLf=failed
        If Risposta <> "0" + vbCrLf Then
            If Risposta = "1" + vbCrLf Then
                Stringa = "ERRORE! La centralina non riesce a cancellare la memoria! (CancellaFlash05)"
            Else
                Stringa = "ERRORE! La centralina non risponde! (CancellaFlash06) -->" + Char2ascii(Risposta)
            End If
            Intero = MsgBox(Stringa, vbAbort, vbIgnore)
            ScriviErroreSuLog Stringa
            Select Case Intero
                'Case vbRetry
                '    fMain.MSComm1.Output = Chr$(18)
                '    Sleep (5000)
                '    GoTo retry1
                Case vbAbort
                    Exit Function
                Case vbIgnore
                    Me.MSComm1.Output = Chr$(18)  'CTRL+R
                    Risposta = InputComTimeOut(10)
                    If Left(Risposta, Len(Risposta) - 2) <> Versione Then
                        Stringa = "ERRORE! La centralina non risponde! (CancellaFlash10) -->" + Risposta
                        MsgBox Stringa, vbOKOnly
                        Exit Function
                    End If
                Me.MSComm1.Output = LeggiDFPNT + vbCr
                Risposta = InputComTimeOut(5)
                Select Case Risposta
                    Case "0"
                        Stringa = "OK la memoria è stata cancellata!"
                        MsgBox Stringa, vbOKOnly
                    Case Else
                        Stringa = "Ci sono problemi a cancellare la menoria!"
                        MsgBox Stringa, vbOKOnly
                        Exit Function
                End Select
            End Select

            
        End If
        'ScriviErroreSuLog Risposta
        'Me.MSComm1.Output = Chr$(250)
        Risposta = InputComTimeOutBin(5, 1)
        If Risposta <> "#" Then
            Stringa = "ERRORE! La centralina non risponde! (CancellaFlash07) -->" + Risposta
            MsgBox Stringa, vbOKOnly
            Exit Function
        End If
        
        'ScriviErroreSuLog Risposta
        
        Me.MSComm1.Output = Chr$(18)  'CTRL+R
        Risposta = InputComTimeOut(5)
        If Risposta = "TimeOut" Then
            Me.MSComm1.Output = CTRLC
            Risposta = InputComTimeOut(5)
        End If
        If Left(Risposta, Len(Risposta) - 2) <> Versione Then
            Stringa = "ERRORE! La centralina non risponde! (CancellaFlash08) -->" + Risposta
            MsgBox Stringa, vbOKOnly
            Exit Function
        End If
        
        'ScriviErroreSuLog Risposta
        
        CancellaFlash = True
        Me.MSComm1.InBufferCount = 0
End Function


' L'evento OnComm viene utilizzato per l'intercettazione
' di eventi ed errori di comunicazione.
Private Static Sub MSComm1_OnComm()
    Dim EVMsg$
    Dim ERMsg$
    
    If isTerminal = False Then Exit Sub
    
    ' Sceglie i diversi casi a seconda del valore della proprietà CommEvent.
    Select Case MSComm1.CommEvent
        ' Messaggi degli eventi.
        Case comEvReceive
            Dim Buffer As Variant
            Buffer = MSComm1.Input
            Debug.Print "Ricezione - " & StrConv(Buffer, vbUnicode)
            ShowData frmTerminal.txtTerm, (StrConv(Buffer, vbUnicode))
        Case comEvSend
        Case comEvCTS
            EVMsg$ = "Rilevata modifica in CTS"
        Case comEvDSR
            EVMsg$ = "Rilevata modifica in DSR"
        Case comEvCD
            EVMsg$ = "Rilevata modifica in CD"
        Case comEvRing
            EVMsg$ = "Il telefono sta squillando"
        Case comEvEOF
            EVMsg$ = "Raggiunta la fine del file"

        ' Messaggi di errore.
        Case comBreak
            ERMsg$ = "Ricevuta interruzione"
        Case comCDTO
            ERMsg$ = "Timeout Carrier Detect"
        Case comCTSTO
            ERMsg$ = "Timeout CTS"
        Case comDCB
            ERMsg$ = "Errore durante il recupero di DCB"
        Case comDSRTO
            ERMsg$ = "Timeout DSR"
        Case comFrame
            ERMsg$ = "Errore di frame"
        Case comOverrun
            ERMsg$ = "Errore di overrun"
        Case comRxOver
            ERMsg$ = "Overflow del buffer di ricezione"
        Case comRxParity
            ERMsg$ = "Errore di parità"
        Case comTxFull
            ERMsg$ = "Buffer di trasmissione pieno"
        Case Else
            ERMsg$ = "Errore o evento sconosciuto"
    End Select
    
    If Len(EVMsg$) Then
        ' Visualizza i messaggi degli eventi sulla barra di stato.
        frmTerminal.sbrStatus.Panels("Status").Text = "Stato: " & EVMsg$
                
        ' Attiva il timer in modo che il messaggio sulla barra
        ' di stato venga cancellato dopo 2 secondi.
        frmTerminal.Timer2.Enabled = True
        
    ElseIf Len(ERMsg$) Then
        ' Visualizza i messaggi di errore sulla barra di stato.
        frmTerminal.sbrStatus.Panels("Status").Text = "Stato: " & ERMsg$
        
        ' Visualizza i messaggi di errore in una finestra di messaggio.
        Beep
        Ret = MsgBox(ERMsg$, 1, "Fare clic su Annulla per uscire o su OK per ignorare il messaggio.")
        
        ' Se l'utente fa clic su Annulla (2)...
        If Ret = 2 Then
            fMain.MSComm1.PortOpen = False    ' Chiude la porta ed esce.
        End If
        
        ' Attiva il timer in modo che il messaggio sulla barra
        ' di stato venga cancellato dopo 2 secondi.
        frmTerminal.Timer2.Enabled = True
    End If
End Sub
