Attribute VB_Name = "Modulo1"
Option Explicit

Private Sub Main()
  'show the splash screen
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
    FileIni = sGetAppPath + "Simamet.ini"
    SE = ","
    Stazione = "Simamet"
    frmSplash.lblWarning = ""


    'Stabilisce quali sono le impostazioni internazionali
    Decimale = QueryValue(HKEY_USERS, ".Default\Control Panel\International", "sDecimal")
    If Decimale <> "" Then
        Decimale = Left(Decimale, Len(Decimale) - 1)
    Else
        Decimale = Mid(Format(0.5, "0.0"), 2, 1)
    End If
    
    Migliaia = QueryValue(HKEY_USERS, ".Default\Control Panel\International", "sThousand")
    If Migliaia <> "" Then
        Migliaia = Left(Migliaia, Len(Migliaia) - 1)
    Else
        Migliaia = Mid(Format(1000, "#,###"), 2, 1)
    End If
    ConnessioneRemota = False
    CTRLC = Chr(3)
    ProgrammazioneCaricata = False
    Stazione = "Simamet"
    FattoreBatteriaInterna = 2.8
    fDebug = False
    lDebug = False
    i = InStr(Command$, "/lab")
    If i <> 0 Then lDebug = True
    i = InStr(Command$, "/debug")
    If i <> 0 Then fDebug = True
    fdn = 0
    'Apre il file di log
    If fDebug Then
        FileName = sGetAppPath + "log.txt"
        fdn = FreeFile
        Open FileName For Append As #fdn
        Print #fdn,
        Print #fdn, "-----------------------------------------------------"
        Print #fdn, "Multipar"
        Print #fdn, Date, Time

    End If
    
    If SetInIDE() Then
        lDebug = True
    End If
End Sub

Public Sub CaricaSetup()
    Dim Filnb As Integer
    Dim i As Integer
    Dim Stringa As String
    
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
    fMain.CmDialog1.FileName = ""
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
    
    
    FileOut = ""
    'Me.MousePointer = vbDefault
    'AggiornaTbs (tbsOptions.SelectedItem.Index)
    Close #Filnb
    Exit Sub
Annulla:
    'Me.MousePointer = vbDefault
    DoEvents
    'CloseCom
End Sub

Public Sub SalvaSetup()
    Dim Filnb As Integer
    Dim i As Integer
    
    'i = tbsOptions.SelectedItem.Index - 1
    'Applica (i)

    
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
    fMain.CmDialog1.FileName = ""
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
    
    
    FileOut = ""
    'Me.MousePointer = vbDefault
    
    Close #Filnb
    Exit Sub
Annulla:
    'Me.MousePointer = vbDefault
    DoEvents
    'CloseCom

End Sub



