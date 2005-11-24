Attribute VB_Name = "Modulo1"
Option Explicit
Private Sub Main()
  'show the splash screen
   Versione = "MultiAmel"
   frmSplash.Show
   'Execute Init instructions
   Init
   DoEvents
   'Call Sleep(2000)
  'show the main application
   fMain.Show
   DoEvents
  'perform any other startup functions as required by your program
  '{code}
  'unload the splash screen and free its memory
   Unload frmSplash
   Set frmSplash = Nothing
End Sub

Public Sub Init()
    Dim nfile As Integer
    Dim rint As Integer
    Dim Path As String
    Dim i As Long
    
    Path = sGetAppPath()

    

    If Versione = "MultiAmel" Then
        TestataPrg = "Poseidon Sensors Setup File"
        FileIni = sGetAppPath + "MultiAmel.ini"

    Else
        TestataPrg = "Multipar Sensors Setup File"
        FileIni = sGetAppPath + "Multipar.ini"
    End If
    
    
    SE = ";"    'Il separatore di elenco è la virgola
    frmSplash.lblWarning = ""

    'Legge dal file i valori di Zero e KpH
    Messaggio = sReadINI("Tarature", "KpH", FileIni)
    KpH = Val(Messaggio)
    If KpH = 0 Then KpH = 2.3
    Messaggio = sReadINI("Tarature", "Zero", FileIni)
    Zero = Val(Messaggio)
    If Zero = 0 Then Zero = 2.449
    Messaggio = sReadINI("Tarature", "Kc", FileIni)
    Kc = Val(Messaggio)
    If Kc = 0 Then Kc = 16.4
    

'    Messaggio = Path + "MH4p.ini"
'    nfile = OpenFile(Messaggio, "I")
'    On Error Resume Next
'    If nfile <> 0 Then
'        Input #nfile, Zero
'        Input #nfile, KpH
'        Close nfile
'        If Zero <= 0 Or Zero >= 14 Then
'            Zero = 2.449
'            frmSplash.lblWarning = "AVVISO: File configurazione errato!"
'        End If
'        If KpH <= 0 Or KpH >= 10 Then
'            KpH = 2.3
'            frmSplash.lblWarning = "AVVISO: File configurazione errato!"
'        End If
'    Else
'        frmSplash.Hide
'        DoEvents
'        frmSplash.lblWarning = "AVVISO: File configurazione mancante!"
'        DoEvents
'        frmSplash.Show
'        Zero = 2.449
'        KpH = 2.3
'    End If
'
'    Kc = 16.4
    'Attivazione di tutti i canali
    For i = 0 To 3
        Canale(i).Attivo = True
    Next
    
    Stazione = Versione

    'definizione parametri per i vari canali
    'Canale 0 A36 Temperatura
    Canale(0).Attivo = True
    Canale(0).Nome = "Temperatura"
    Canale(0).UnitaMisura = "C"
    Canale(0).Bitmin = 0
    Canale(0).Bitmax = 4095
    Canale(0).valMin = 0
    Canale(0).valMax = 5
    Canale(0).valOff = 0
    'Canale 1 A37 Conducibilità
    Canale(1).Attivo = True
    Canale(1).Nome = "Conducibilita'"
    Canale(1).UnitaMisura = "mS"
    Canale(1).Bitmin = 0
    Canale(1).Bitmax = 4095
    Canale(1).valMin = 0
    Canale(1).valMax = 5
    Canale(1).valOff = 0
    'Canale 2 A38 Livello
    Canale(2).Attivo = True
    Canale(2).Nome = "Livello"
    Canale(2).UnitaMisura = "m"
    Canale(2).Bitmin = 813
    Canale(2).Bitmax = 4063
    Canale(2).valMin = 0
    Canale(2).valMax = 20
    Canale(2).valOff = 0
    'Canale 3 A39 pH
    Canale(3).Attivo = True
    Canale(3).Nome = "pH"
    Canale(3).UnitaMisura = " "
    Canale(3).Bitmin = 0
    Canale(3).Bitmax = 4095
    Canale(3).valMin = 0
    Canale(3).valMax = 5
    Canale(3).valOff = 0
    'Canale 4 Temperatura interna
    Canale(4).Attivo = True
    Canale(4).Nome = "Temp. Interna"
    Canale(4).UnitaMisura = "C"
    Canale(4).Bitmin = 0
    Canale(4).Bitmax = 4095
    Canale(4).valMin = 0
    Canale(4).valMax = 5
    Canale(4).valOff = 0
    'Canale 5 Tensione batteria
    Canale(5).Attivo = True
    Canale(5).Nome = "Tens. Batteria"
    Canale(5).UnitaMisura = "volt"
    Canale(5).Bitmin = 0
    Canale(5).Bitmax = 4095
    Canale(5).valMin = 0
    Canale(5).valMax = 5
    Canale(5).valOff = 0

    FattoreBatteriaInterna = 2.8
    fDebug = False
    lDebug = False
    i = InStr(Command$, "/lab")
    If i <> 0 Then lDebug = True
    i = InStr(Command$, "/debug")
    If i <> 0 Then fDebug = True
    'If Command$ = "/lab" Then lDebug = True
    'fDebug = False
    CTRLC = Chr(3)
    fdn = 0
    'Apre il file di log
    If fDebug Then
        FileName = sGetAppPath + "log.txt"
        fdn = FreeFile
        Open FileName For Append As #fdn
        Print #fdn,
        Print #fdn, "-----------------------------------------------------"
        Print #fdn, Versione
        Print #fdn, Date, Time

    End If

End Sub

Public Sub CaricaSetup()
    Dim Filnb As Integer
    Dim i As Integer
    Dim Stringa As String
    
    
    Stringa = sReadINI("Configurazione", "UltimaProgrammazione", FileIni)
    If Stringa <> "" Then
        fMain.CmDialog1.InitDir = Stringa
    End If

    On Error GoTo Annulla
    fMain.CmDialog1.CancelError = True
    'Controlla se si vuole sostituire il file,
    'che la directory eventualmente immessa esista,
    'non prende in considerazione files e directory a sola lettura
    'non mostra la casella sola lettura
    fMain.CmDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNHideReadOnly
    'Filtri di dialogo
    fMain.CmDialog1.Filter = "File Programmazione (*.prg)|*.prg|Tutti i file (*.*)|*.*"
    NewPath (sGetAppPath())
    fMain.CmDialog1.FileName = Stringa
    fMain.CmDialog1.ShowOpen
    On Error GoTo 0
    FileOut = fMain.CmDialog1.FileName
    DoEvents
        
    'Me.MousePointer = vbHourglass
    'Salva i dati
    Filnb = FreeFile
    Open FileOut For Input As #Filnb
    Input #Filnb, Stringa
    If Stringa <> TestataPrg Then
        Messaggio = "ERRORE! " + FileOut + " non è un file di configurazione!"
        MsgBox (Messaggio)
        'Me.MousePointer = vbNormal
        Exit Sub
    End If
        Input #Filnb, Stazione
    For i = 0 To MaxCanali
        Input #Filnb, Canale(i).Nome
        Input #Filnb, FileOut
        Canale(i).Attivo = CBool(FileOut)
        Input #Filnb, Canale(i).UnitaMisura
        Input #Filnb, Canale(i).Bitmin
        Input #Filnb, Canale(i).Bitmax
        Input #Filnb, Canale(i).valMin
        Input #Filnb, Canale(i).valMax
        Input #Filnb, Canale(i).valOff
    Next
    
    

    'Me.MousePointer = vbDefault
    'AggiornaTbs (tbsOptions.SelectedItem.Index)
    Close #Filnb
    i = WriteINI("Configurazione", "UltimaProgrammazione", fMain.CmDialog1.FileName, FileIni)
    FileOut = ""
    Exit Sub
Annulla:
    'Me.MousePointer = vbDefault
    DoEvents
    'CloseCom
End Sub

Public Sub SalvaSetup()
    Dim Filnb As Integer
    Dim i As Long
    Dim Stringa As String
    
    'i = tbsOptions.SelectedItem.Index - 1
    'Applica (i)

    Stringa = sReadINI("Configurazione", "UltimaProgrammazione", FileIni)
    If Stringa <> "" Then
        fMain.CmDialog1.InitDir = Stringa
    End If
    
    On Error GoTo Annulla
    fMain.CmDialog1.CancelError = True
    'Controlla se si vuole sostituire il file,
    'che la directory eventualmente immessa esista,
    'non prende in considerazione files e directory a sola lettura
    'non mostra la casella sola lettura
    fMain.CmDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNHideReadOnly
    'Filtri di dialogo
    fMain.CmDialog1.Filter = "File Programmazione (*.prg)|*.prg|Tutti i file (*.*)|*.*"
    NewPath (sGetAppPath())
    fMain.CmDialog1.FileName = Stringa
    fMain.CmDialog1.ShowSave
    On Error GoTo 0
    FileOut = fMain.CmDialog1.FileName
    DoEvents
        
    'Me.MousePointer = vbHourglass
    'Salva i dati
    Filnb = FreeFile
    Open FileOut For Output As #Filnb
    Print #Filnb, TestataPrg
    Print #Filnb, Stazione
    For i = 0 To MaxCanali
        Print #Filnb, Canale(i).Nome
        If Canale(i).Attivo = True Then
            Print #Filnb, "True"
        Else
            Print #Filnb, "False"
        End If
        Print #Filnb, Canale(i).UnitaMisura
        Print #Filnb, Canale(i).Bitmin
        Print #Filnb, Canale(i).Bitmax
        Print #Filnb, Str(Canale(i).valMin)
        Print #Filnb, Str(Canale(i).valMax)
        Print #Filnb, Str(Canale(i).valOff)
    Next
    
    
    'Me.MousePointer = vbDefault
    
    Close #Filnb
    i = WriteINI("Configurazione", "UltimaProgrammazione", fMain.CmDialog1.FileName, FileIni)
    FileOut = ""

    Exit Sub
Annulla:
    'Me.MousePointer = vbDefault
    DoEvents
    'CloseCom

End Sub
