Attribute VB_Name = "Funzioni"
Option Explicit
Public Sub OpenCom()
    'Apre la porta com
    'Se e' andata bene ComOk e' True altrimenti e' False
    Dim Msg As String

    On Error GoTo ErroreCom
    ComOk = False
    'Apre la porta seriale se non è già aperta
    If fMain.MSComm1.PortOpen = False Then fMain.MSComm1.PortOpen = True
    ComOk = True
    Exit Sub
ErroreCom:
    Select Case Err.Number
        Case 8005  'La Com è già aperta
            Msg = "Errore la porta Com" + Str$(ComPort) + " è già in uso"
            MsgBox Msg, , "Errore"
            Err.Clear   ' Cancella i campi dell'oggetto
            ComOk = False
            Exit Sub
        Case 8002
            Msg = "Errore la porta Com" + Str$(ComPort) + " non esiste!"
            MsgBox Msg, , "Errore"
            Err.Clear   ' Cancella i campi dell'oggetto
            ComOk = False
            Exit Sub
        Case Else
            ErrHandler
            Exit Sub
    End Select

End Sub

Public Sub CloseCom()
    'Chiude la porta seriale se non è già chiusa
    fMain.MSComm1.InBufferCount = 0
    If fMain.MSComm1.PortOpen = True Then fMain.MSComm1.PortOpen = False
End Sub

Public Sub WaitCom()
'Aspetta che sulla COM ci siano dei caratteri.
'Senza TIMEOUT!
    Do
        DoEvents
    Loop Until fMain.MSComm1.InBufferCount >= 1
End Sub

Public Function InputComTimeOut(TimeOut As Integer) As String
'Attende un input il cui terminatore e' vbLF
'Con TIMEOUT
    Dim TimeStop As Long
    Dim Linea As String
    Dim Dummy As String
    
        TimeStop = Timer + TimeOut
        fMain.MSComm1.InputLen = 1
        Do
            DoEvents

        Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
        If fMain.MSComm1.InBufferCount >= 1 Then
            Linea = ""
            Dummy = ""
            TimeStop = Timer + TimeOut ' Imposta l'ora di fine
            Do Until Dummy = vbLf Or (Timer > TimeStop)
                DoEvents
                Dummy = fMain.MSComm1.Input
                Linea = Linea + Dummy
            Loop
        Else
            Linea = "TimeOut"
        End If
        InputComTimeOut = Linea

End Function

Public Function InputComTimeOutTerm(TimeOut As Integer, Terminator As Byte) As String
'Attende un input il cui terminatore e' Terminator
'Con TIMEOUT

        Dim TimeStop As Long
        Dim Linea As String
        Dim Dummy As String

        TimeStop = Timer + TimeOut
        fMain.MSComm1.InputLen = 1
        Do
            DoEvents
        Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
        If fMain.MSComm1.InBufferCount >= 1 Then
            Linea = ""
            Dummy = ""
            TimeStop = Timer + TimeOut ' Imposta l'ora di fine
            Do Until Dummy = Chr(Terminator) Or (Timer > TimeStop)
                DoEvents
                Dummy = fMain.MSComm1.Input
                Linea = Linea + Dummy
            Loop
        Else
            Linea = "TimeOut"
        End If
        InputComTimeOutTerm = Linea

End Function
Public Function InputComTimeOutBin(TimeOut As Integer, NumByte As Integer) As String
'Attende un input binario senza terminatore
'Con TIMEOUT
    Dim TimeStop As Long
    Dim Linea As String
    'Dim dummy As String
    
        'fMain.MSComm1.InputMode = comInputModeBinary

        TimeStop = Timer + TimeOut
        fMain.MSComm1.InputLen = NumByte
        Do
            DoEvents
        Loop Until (fMain.MSComm1.InBufferCount >= NumByte) Or (Timer > TimeStop)
'        If fMain.MSComm1.InBufferCount >= NumByte Then
'            linea = ""
'        Else
'            linea = "TimeOut"
'        End If
        Linea = fMain.MSComm1.Input
        If Linea <> "" Then
            Debug.Print "inputCom--->"; Char2ascii(Linea)
        End If
        InputComTimeOutBin = Linea

End Function

Public Function InputComTimeOutBin2(TimeOut As Integer, NumByte As Integer) As String
'Attende un input binario senza terminatore
'Con TIMEOUT sul singolo carattere!!!
    Dim TimeStop As Long
    Dim Linea As String
    Dim TimerOut As Boolean
    Dim BloccoDati(32768) As Byte
    Dim Blocco() As Byte
    Dim iBloccoDati As Long
    Dim TimeOuts As Integer
    Dim i As Long
    

    iBloccoDati = 0
    'ReDim BloccoDati(1000)
    Do
        DoEvents
        TimeStop = Timer + 2
        Do
            DoEvents
        Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
        If fMain.MSComm1.InBufferCount = 0 Then
            TimeOuts = TimeOuts + 1
        Else
            'dati = dati + fMain.MSComm1.InBufferCount
            Blocco = fMain.MSComm1.Input
            For i = LBound(Blocco) To UBound(Blocco)
                BloccoDati(iBloccoDati) = Blocco(i)
                iBloccoDati = iBloccoDati + 1
            Next i
            TimeOuts = 0
        End If
        If TimeOuts > 3 Then Exit Do
        
        
        DoEvents
        
    Loop Until TimeOuts > 5
    
    For i = 0 To iBloccoDati - 1
        InputComTimeOutBin2 = InputComTimeOutBin2 + Chr$(BloccoDati(i))
    Next i

End Function


Public Function ScaricaProgrammazione() As Boolean
'Scarica la programmazione dal TFX
'Se ci sono canali attivi restituisce true altrimenti false
Dim Blocco() As Byte        'Blocco dati temporaneo in bytes
Dim Buffer As Byte       'buffer temporaneo per i dati
Dim BloccoDati() As Byte    'Blocco dati
Dim iBloccoDati As Long
Dim Bytes As Long           'Numero bytes scaricati
Dim TimeOuts As Long        'Contatore dei Time Out
Dim iDumm As Long
Dim Dummy As String
Dim Float As Single
Dim Intero As Integer
Dim Lungo As Long
Dim Stringa As String
Dim lStringa As Long
Dim CanaliAttivi As Boolean
Dim dati As Long
Dim iBlocco As Long
Dim TimeStop As Integer
Dim i As Long

CanaliAttivi = False
fMain.MSComm1.InBufferCount = 0
fMain.MSComm1.InputLen = 0
fMain.MSComm1.RThreshold = 0
fMain.MSComm1.InBufferCount = 0
fMain.MSComm1.InputMode = comInputModeBinary

OpenCom

fMain.MSComm1.InBufferCount = 0
fMain.MSComm1.Output = InfoProg + vbCr

TimeOuts = 0
Bytes = 0
Intero = 0
dati = 0
iBloccoDati = 0
ReDim BloccoDati(1000)
Do
    DoEvents
    TimeStop = Timer + 2
    Do
        DoEvents
    Loop Until (fMain.MSComm1.InBufferCount >= 1) Or (Timer > TimeStop)
    If fMain.MSComm1.InBufferCount = 0 Then
        TimeOuts = TimeOuts + 1
    Else
        dati = dati + fMain.MSComm1.InBufferCount
        Blocco = fMain.MSComm1.Input
        For i = LBound(Blocco) To UBound(Blocco)
            BloccoDati(iBloccoDati) = Blocco(i)
            iBloccoDati = iBloccoDati + 1
        Next i
        TimeOuts = 0
    End If
    If TimeOuts > 3 Then Exit Do
    
    
    DoEvents
    
Loop Until Bytes >= 910


If iBloccoDati < 910 Then
    Messaggio = "Errore! Ricevuti" + Str(iBloccoDati) + " dati invece di 910"
    MsgBox (Messaggio)
'    For i = 0 To iBloccoDati - 1
'        Debug.Print Asc(Blocco(i))
'    Next
    ScaricaProgrammazione = False
    Exit Function
    'Esci
End If

'Leggo i dati di tutti i canali
iBlocco = 40
For i = 0 To MaxCanali
    'legge la lunghezza del nome del canale
    lStringa = BloccoDati(iBlocco)
    iBlocco = iBlocco + 1
    'If lStringa <> 16 Then Debug.Print "errore!"
   
    'Print #Filnb2, lStringa
    'leggo il nome del canale
    Canale(i).Nome = bMID(BloccoDati, iBlocco, lStringa - 1)
    'Print #Filnb2, Canale(i).Nome
    iBlocco = iBlocco + lStringa

    'leggo se è attivo o meno. lStringa è una variabile riciclata
    lStringa = BloccoDati(iBlocco)
    iBlocco = iBlocco + 1
    'Print #Filnb2, lStringa
    If lStringa = 0 Then
        Canale(i).Attivo = False
    Else
        Canale(i).Attivo = True
        CanaliAttivi = True
    End If

    'Leggo lunghezza stringa unità di misura
    lStringa = BloccoDati(iBlocco)
    iBlocco = iBlocco + 1

    'leggo l'unità di misura
    Canale(i).UnitaMisura = bMID(BloccoDati, iBlocco, lStringa - 1)
    iBlocco = iBlocco + lStringa
    'Print #Filnb2, Canale(i).UnitaMisura

    'Leggo Bitmin
    Dummy = bMID(BloccoDati, iBlocco, 4)
    iBlocco = iBlocco + 4
    Canale(i).Bitmin = String2long(Dummy)
    'Print #Filnb2, Canale(i).Bitmin

    'leggo Bitmax
    Dummy = bMID(BloccoDati, iBlocco, 4)
    iBlocco = iBlocco + 4
    Canale(i).Bitmax = String2long(Dummy)
    'Print #Filnb2, Canale(i).Bitmax

    'leggo Valmin
    Canale(i).sValmin = SwapString(bMID(BloccoDati, iBlocco, 4))
    Canale(i).valMin = String2single(Canale(i).sValmin)
    'Print #Filnb2, Canale(i).sValmin
    iBlocco = iBlocco + 4

    'leggo Valmax
    Canale(i).sValmax = SwapString(bMID(BloccoDati, iBlocco, 4))
    Canale(i).valMax = String2single(Canale(i).sValmax)
    'Print #Filnb2, Canale(i).sValmax
    iBlocco = iBlocco + 4
 
   'leggo Valoff
    Canale(i).sValoff = SwapString(bMID(BloccoDati, iBlocco, 4))
    Canale(i).valOff = String2single(Canale(i).sValoff)
    'Print #Filnb2, Canale(i).sValoff
    iBlocco = iBlocco + 4
    
    'Qui si legge la soglia
    'Lungo = String2long(SwapString(bMID(BloccoDati, iBlocco, 4)))
    Lungo = String2long(bMID(BloccoDati, iBlocco, 4))
    'Canale(i).sSoglia = SwapString(Mid(BloccoDati, iBlocco, 4))
    Canale(i).Soglia = Count2value(CByte(i), Lungo)
    'Print #Filnb2, Canale(i).sSoglia
    iBlocco = iBlocco + 4
    
'    'Qui si legge il valore dell'allarme
'    Lungo = String2long(SwapString(bMID(BloccoDati, iBlocco, 4)))
'    Canale(i).vAllarme = Count2value(i, Lungo)
'    If Canale(i).vAllarme <> 0 Then
'        Canale(i).Allarme = True
'    End If
'    iBlocco = iBlocco + 4
    
Next

ScaricaProgrammazione = CanaliAttivi

End Function

Public Function SalvaProgrammazione(FileOut As String) As Boolean
    Dim Filnb As Long
    Dim i As Long

    SalvaProgrammazione = False
    On Error GoTo Annulla
    If FileOut = "" Then
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
        'fMain.CmDialog1.InitDir = InitDirPrg
        fMain.CmDialog1.ShowSave
        On Error GoTo 0
        FileOut = fMain.CmDialog1.FileName
        DoEvents
    End If
    'Me.MousePointer = vbHourglass
    'Salva i dati
    Filnb = FreeFile
    Open FileOut For Output As #Filnb
    Print #Filnb, TestataPrg
    Print #Filnb, Stazione
    For i = 0 To MaxCanali
        Print #Filnb, Canale(i).Nome
        Print #Filnb, Canale(i).Attivo
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
    SalvaProgrammazione = True
    Exit Function
Annulla:
    'Me.MousePointer = vbDefault
    DoEvents
    'CloseCom
End Function


Public Function Data2sec70(d As Date) As Long
    'Trasforma una data in numero di secondi a partire
    'dall 1/1/1970 alle 0:0:0
    Dim PartenzaSerial As Double
    Dim ArrivoSerial As Double
    
    'trasformazione data di partenza in numero seriale
    PartenzaSerial = DateSerial(1970, 1, 1)
    'Trasformazione data di arrivo in numero seriale
    ArrivoSerial = Dat2Ser(d)
    'MsgBox (Str(ArrivoSerial))
    'Calcolo differenza
    ArrivoSerial = ArrivoSerial - PartenzaSerial
    'Trasformazione da giorni in secondi
    ArrivoSerial = (ArrivoSerial * 86400) + CorrezioneTempo
    Data2sec70 = ArrivoSerial
End Function

Public Function Sec80toDate(seconds As Long) As Date
    'Trasforma il numero di secondi a partire
    'dall 1/1/1980 alle 0:0:0 in una data
    'Serve per trasformare la data del TFX11 contenuta
    'nella variabile ? in una data
    Dim PartenzaSerial As Double
    Dim ArrivoSerial As Double
    Dim Data As Date
    
    'trasformazione data di partenza in numero seriale
    PartenzaSerial = DateSerial(1980, 1, 1)
    'trasformazione in secondi
    PartenzaSerial = PartenzaSerial * 86400
    'Aggiunta ai secondi seconds
    ArrivoSerial = PartenzaSerial + seconds
    'MsgBox (Str(ArrivoSerial))
    'Trasformazione da secondi a giorni
    ArrivoSerial = ArrivoSerial / 86400
    'trasformazione in data
    Data = CDate(ArrivoSerial)
    'MsgBox (Format(Data, "yyyy/mm/dd hh:nn:ss"))
    Sec80toDate = Data
End Function


Public Function String2sn(Stringa As String) As String
    'Trasforma una stringa in un numero seriale
    Dim lStringa As Integer
    Dim iSerNumb As Long
    Dim Dummy As String
    Dim i As Integer
    
    lStringa = Len(Stringa)
    If lStringa <> 4 Then
        Messaggio = "La lunghezza del numero di serie è errata ->" + Str(lStringa) + " invece di 4"
        MsgBox (Messaggio)
    End If
    
    'Capovolge il numero poichè è codificato in big endian (prima il byte più significativo)
    Stringa = SwapString(Stringa)
    
    'Infine la trasformazione in numero
    iSerNumb = bytes2long(Stringa)
    String2sn = Format(iSerNumb)
End Function

Public Function CalcolaBattFact() As Single
'Calcola il valore di BattFact in funzione dei valori
'immessi nella programmazione
    CalcolaBattFact = Canale(5).valMax / 5
    If CalcolaBattFact = 0 Then CalcolaBattFact = 2.8
End Function

Public Function ScriviErroreSuLog(Errore As String) As Boolean
    Dim FileLogName As String
    Dim nf As Long
    
    FileLogName = sGetAppPath + App.EXEName + ".log"
    nf = FreeFile
    Open FileLogName For Append As #nf
    Print #nf, "---------------------------------------------"
    Print #nf, Now, App.EXEName
    Print #nf, "Err number -->"; Err.Number
    Print #nf, Err.Description
    Print #nf, Errore
    Close nf
    ScriviErroreSuLog = True
    
End Function
