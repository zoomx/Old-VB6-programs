VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form fCounter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scarico dati"
   ClientHeight    =   2790
   ClientLeft      =   3570
   ClientTop       =   3435
   ClientWidth     =   6750
   Icon            =   "fCounter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2790
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1920
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "Bytes"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Scarico dati in corso"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "fCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Filnb As Long     'Numero file output libero
Public bAscii As Boolean

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Close Filnb
        Close
        Esci
    End If
End Sub

Public Sub Esci()
    OpenCom
    fMain.MSComm1.Output = CTRLC
    Close Filnb
    'Close
    Unload Me
    fMain.MousePointer = vbDefault
    fMain.MSComm1.InputMode = comInputModeText
    fMain.AbilitaTasti
    fMain.Show
End Sub

Public Sub Scarica()

'ATTENZIONE!
'Nel trasferire i dati numerici da TFX a PC bisogna
'invertire l'ordine dei bytes. Ci� viene eseguito
'dalla routine Strin2long o SwapString



Text1.Visible = False

Dim Blocco() As Byte        'Blocco dati temporaneo in bytes
Dim Buffer As Byte       'buffer temporaneo per i dati
Dim BloccoDati() As Byte    'Blocco dati
Dim iBloccoDati As Long     'Indice all'interno di BloccoDati()
Dim DFPNT As Long           'Numero di bytes da scaricare
Dim DFPNT2 As Long          'Numero bytes da scaricare dal Palm
Dim Bytes As Long           'Numero bytes scaricati
Dim LungCounter As Long
Dim Barra As Double
Dim IncBarra As Double 'Incremento barra contatore per ogni riga
Dim TimeOuts As Long        'Contatore dei Time Out
Dim iDumm As Long
Dim Dummy As String
Dim Float As Single

'Dim Stazione As String      'Stringa che contiene il nome della stazione
Dim lStazione As Integer
Dim PAnno As Integer        'Anno di partenza
Dim PMese As Integer
Dim PGiorno As Integer
Dim POra As Integer
Dim PMinuti As Integer
Dim PSecondi As Integer
Dim PData As String         'Data di partenza in stringa
Dim lpData As Long          'Data in numero seriale
Dim SerNumb As String
Dim nSerie As Long           'Numero di serie del Datalogger
'Dim Intervallo As Long   'Intevallo di campionamento in secondi
Dim Intero As Integer
Dim Lungo As Long
Dim Stringa As String
Dim lStringa As Long
Dim CanaliAttivi As Integer
Dim sCanAttivi As String

Dim RS As String

Dim iBlocco As Long
Dim i As Long
Dim j As Long
Dim Tempog As Double  'tempo in giorni
Dim Tempo As Long
Dim dTempo As String
Dim nCanale As Byte
Dim Misura As Integer

Dim TimeStop As Long
Dim NomeFile As String
Dim Linea As String
Dim dati As Long
    
Dim Vbatteria As Single
Dim MyStr As String
      
Dim iResponse As Integer
Dim Emergenza As Boolean
Dim QuantiDati As Long
Dim Tutti255 As Boolean

Tutti255 = False
Emergenza = False

'Stabilisce se il file � binario o meno
bAscii = True
If Right(FileName, 3) = "sim" Then bAscii = False

fMain.Hide

'Azzeramento del buffer di input per evitare che le routines
'per il Palm falliscano
fMain.MSComm1.InBufferCount = 0

If Palm Then
    MsgBox ("Collega il Palm")
Else
    Label1.Caption = "Collegamento in corso..."
End If


DoEvents

Barra = 0
Scaricato = False
ProgressBar1.value = 0

DoEvents

If Palm Then
    Label1.Caption = "In attesa di ricevere i dati dal Palm"
    'Stringa = InputComTimeOutTerm(20, vbCr)
    'In VB5 funziona mentre in VB6 da errore!!!
    Stringa = InputComTimeOutTerm(20, 13)
    If Stringa = "TimeOut" Then
        MsgBox ("Non ricevo dati dal Palm")
        Esci
        Exit Sub
    End If
    DFPNT = Val(Stringa) - 1
    DFPNT2 = DFPNT
    Label1.Caption = "Collegamento in corso..."

Else
 

'Manda il comando di InfoAcq
OpenCom
fMain.MSComm1.InBufferCount = 0
fMain.MSComm1.Output = CTRLC

If fDebug Then Print #fdn, "Scarico Dati"

'Prende la risposta con TimeOut
RS = InputComTimeOut(5)
If fDebug Then Print #fdn, "CTRLC"; RS
If RS = "TimeOut" Then
    MsgBox ("ERRORE -> La centralina non risponde (TIMEOUT)!")
    Esci
    Exit Sub
End If

fMain.MSComm1.InBufferCount = 0
fMain.MSComm1.Output = InfoAcq + vbCr

'Prende la risposta con TimeOut
RS = InputComTimeOut(5)
If fDebug Then Print #fdn, "InfoAcq"; RS
fMain.MSComm1.InBufferCount = 0
If RS = "TimeOut" Then
    MsgBox ("ERRORE -> La centralina non risponde al comando InfoAcq!")
    Esci
    Exit Sub
End If

DFPNT = Val(RS) - 1
'DFPNT dovrebbe essere sempre positivo ma non si sa mai!
If DFPNT < 0 Then DFPNT = 0
DFPNT2 = DFPNT
'If DFPNT = 0 Then
'    'Il puntatore � =0 proviamo a prendere
'    'la copia che sul TFX � nella variabile Spunt
'    fMain.MSComm1.InBufferCount = 0
'    fMain.MSComm1.Output = Spunt + vbCr
'    'Prende la risposta con TimeOut
'    RS = InputComTimeOut(5)
'    fMain.MSComm1.InBufferCount = 0
'    If RS = "TimeOut" Then
'        MsgBox ("ERRORE -> La centralina non risponde al comando Spunt!")
'        Esci
'        Exit Sub
'    End If
'    DFPNT = Val(RS) - 1
'    If DFPNT < 0 Then DFPNT = 0
'End If
End If

If DFPNT = 0 Then
    'Qui eventualmente si puo' mettere una
    'routine di scarico dati di emergenza
    
    'MsgBox ("Non ci sono dati in memoria!")
    'Esci
    'Exit Sub
    
    If Palm = False Then
        iResponse = MsgBox("Sembra che nella centralina non vi sono dati. Avvio uno scarico della memoria?", vbYesNoCancel + vbQuestion + vbDefault, "Simamet")
    If iResponse = 6 Then
        'QuantiDati = InputBox("Quanti dati scarico?", "Simamet", 10000)
        QuantiDati = 450000
        DFPNT = QuantiDati
        Emergenza = True
    Else
        Esci
        Exit Sub
    
    End If
    Else
        MsgBox ("Il Palm sembra non contenere alcun dato!")
        Esci
        Exit Sub
    End If

End If


ProgressBar1.Max = DFPNT + 30
IncBarra = 1
ProgressBar1.value = 0


RS = ""

fMain.MSComm1.InBufferCount = 0
fMain.MSComm1.InputLen = 1
Label2.Caption = ""

Label1.Caption = "Scarico dati in corso"


fMain.MSComm1.InputLen = 0
fMain.MSComm1.RThreshold = 0
Sleep (200)
fMain.MSComm1.InBufferCount = 0
If Emergenza = False Then
    fMain.MSComm1.Output = ScaricoDati + vbCr
Else
    fMain.MSComm1.Output = Scarico_emergenza + vbCr
    Sleep (250)
    fMain.MSComm1.InBufferCount = 0
    fMain.MSComm1.Output = Trim(Str(QuantiDati)) + vbCr
End If

fMain.MSComm1.InputMode = comInputModeBinary

BloccoDati = ""
TimeOuts = 0
Bytes = 0
Intero = 0
dati = 0
iBloccoDati = 0
ReDim BloccoDati(DFPNT + 100)
Do
    DoEvents
    TimeStop = Timer + 1
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
            'Controlla se i dati ricevuti non sono tutti zero
            'Il controllo viene eseguito solo se ci sono pi� di 100 dati
            If iBloccoDati > 100 Then
                TimeStop = 0
                For j = iBloccoDati - 40 To iBloccoDati
                    TimeStop = TimeStop + BloccoDati(j)
                Next
                'Debug.Print TimeStop
                If TimeStop = (255 * 40) Then
                'If TimeStop = 0 Then
                    Tutti255 = True
                    DFPNT = iBloccoDati
                    ProgressBar1.value = ProgressBar1.Max
                    Exit Do
                End If
            End If

        Next i
        ProgressBar1.value = iBloccoDati
        Label2 = Format(iBloccoDati)
        TimeOuts = 0
    End If
    If TimeOuts > 3 Then Exit Do
    
    
    DoEvents
    
Loop Until Bytes = DFPNT
Label2 = Format(iBloccoDati)

If DFPNT > DFPNT2 Then
    DFPNT = DFPNT2
End If


'Salva una copia binaria dei files
Label1.Caption = "Attendere..."
Dummy = sGetAppPath()
Dummy = Dummy + Format(Year(Now), "0000")
Dummy = Dummy + Format(Month(Now), "00")
Dummy = Dummy + Format(Day(Now), "00")
Dummy = Dummy + Format(Hour(Now), "00")
Dummy = Dummy + Format(Minute(Now), "00")
Dummy = Dummy + Format(Second(Now), "00")
Dummy = Dummy + ".bin"
Filnb = FreeFile
Open Dummy For Binary As #Filnb
'For i = 0 To UBound(BloccoDati)
'    Put #Filnb, , BloccoDati(i)
'    'DoEvents
'    Next
Put #Filnb, , BloccoDati()
Close Filnb


If fDebug Then Print #fdn, "Byte Scaricati"; iBloccoDati

If iBloccoDati < DFPNT And Tutti255 = False Then
    Messaggio = "ERRORE! Ricevuti" + Str(iBloccoDati) + " dati invece di" + Str(DFPNT)
    MsgBox (Messaggio)
    Esci
    Exit Sub
End If
    

ProgressBar1.value = ProgressBar1.Max

Label2.Caption = ""
Label1.Caption = "Processamento dei dati"
DoEvents
ProgressBar1.Max = DFPNT + 1

Filnb = FreeFile

If bAscii = True Then
    Open FileOut For Output As #Filnb
Else
    'Se il file esiste gi� lo elimina
    Stringa = Dir(FileOut)
    If Stringa <> "" Then
        Kill (FileOut)
    End If
    'Apre il file
    Open FileOut For Binary As #Filnb
End If

'Controlla che i dati non siano tutti zero
iBlocco = 0
For i = 110 To 210
    iBlocco = iBlocco + BloccoDati(i)
Next
'((200 - 40) * 255)=40800
'If iBlocco = 40800 Then
If iBlocco = 0 Then
    MsgBox "La memoria � vuota, contiene solo zeri!", 48, "Attenzione!!!"
    Esci
    Exit Sub
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

    'leggo se � attivo o meno. lStringa � una variabile riciclata
    lStringa = BloccoDati(iBlocco)
    iBlocco = iBlocco + 1
    'Print #Filnb2, lStringa
    If lStringa = 0 Then
        Canale(i).Attivo = False
    Else
        Canale(i).Attivo = True
    End If

    'Leggo lunghezza stringa unit� di misura
    lStringa = BloccoDati(iBlocco)
    iBlocco = iBlocco + 1

    'leggo l'unit� di misura
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
Next
'leggo mmxcount
Stringa = SwapString(bMID(BloccoDati, iBlocco, 4))
mmxcount = String2single(Stringa)
iBlocco = iBlocco + 4
'leggo msxcount
Stringa = SwapString(bMID(BloccoDati, iBlocco, 4))
msxcount = String2single(Stringa)
iBlocco = iBlocco + 4

'Controllo che il resto dei dati non sia zero!

lStazione = BloccoDati(iBlocco)
iBlocco = iBlocco + 1
Lungo = lStazione

Stazione = bMID(BloccoDati, iBlocco, Lungo - 1)
iBlocco = iBlocco + lStazione

SerNumb = bMID(BloccoDati, iBlocco, 4)
iBlocco = iBlocco + 4
SerNumb = String2sn(SerNumb)

ProgressBar1.value = iBlocco
Label2 = Format(iBlocco)

'preleviamo la data a partire dall'anno
Dummy = bMID(BloccoDati, iBlocco, 4)
iBlocco = iBlocco + 4
PAnno = String2long(Dummy)

Dummy = bMID(BloccoDati, iBlocco, 4)
iBlocco = iBlocco + 4
PMese = String2long(Dummy)

Dummy = bMID(BloccoDati, iBlocco, 4)
iBlocco = iBlocco + 4
PGiorno = String2long(Dummy)

Dummy = bMID(BloccoDati, iBlocco, 4)
iBlocco = iBlocco + 4
POra = String2long(Dummy)

Dummy = bMID(BloccoDati, iBlocco, 4)
iBlocco = iBlocco + 4
PMinuti = String2long(Dummy)

Dummy = bMID(BloccoDati, iBlocco, 4)
iBlocco = iBlocco + 4
PSecondi = String2long(Dummy)

PData = Format(PGiorno, "0#") + "/" + Format(PMese, "0#") + "/" + Format(PAnno, "0#") + " "
PData = PData + Format(POra, "0#") + ":" + Format(PMinuti, "0#") + ":" + Format(PSecondi, "0#")

'L'intervallo di campionamento in secondi
Dummy = bMID(BloccoDati, iBlocco, 4)
iBlocco = iBlocco + 4
Intervallo = String2long(Dummy)
ProgressBar1.value = iBlocco

'Copia dei dati dei canali attivi
CanaliAttivi = 0
For i = 0 To MaxCanali
'Il canale � attivo?
If Canale(i).Attivo = True Then
            sCanale(CanaliAttivi).Nome = Canale(i).Nome
            sCanale(CanaliAttivi).UnitaMisura = Canale(i).UnitaMisura
            sCanale(CanaliAttivi).Bitmin = Canale(i).Bitmin
            sCanale(CanaliAttivi).Bitmax = Canale(i).Bitmax
            sCanale(CanaliAttivi).valMin = Canale(i).valMin
            sCanale(CanaliAttivi).valMax = Canale(i).valMax
            sCanale(CanaliAttivi).valOff = Canale(i).valOff
            CanaliAttivi = CanaliAttivi + 1
    End If
Next


'Stampiamo
If bAscii = True Then
    Print #Filnb, App.Title
    Print #Filnb,
    Print #Filnb, "Stazione "; Stazione
    Print #Filnb, "Datalogger numero "; SerNumb
    Print #Filnb, "Acquisizione partita il "; PGiorno; "/"; PMese; "/"; PAnno;
    Print #Filnb, " alle "; POra; ":"; PMinuti; ":"; PSecondi
    Print #Filnb, "Intervallo di campionamento "; Intervallo; " secondi"
    If Palm = True Then
        Print #Filnb,
    Else
        Print #Filnb, "Tensione batteria "; fMain.StatusBar1.Panels(1).Text; '" volt"
    End If
    Print #Filnb,
    Print #Filnb, "Canali"
    Print #Filnb,
    Print #Filnb, "N.Canale; NomeCanale; Unit� di misura; Bitmin; Bitmax; Valmin; Valmax; Valoff"
    Print #Filnb,
    For i = 0 To MaxCanali
        'Il canale � attivo?
        If Canale(i).Attivo = True Then
            Print #Filnb, i + 1; ";";       'Il numero del canale � aumentato di uno
            Print #Filnb, Canale(i).Nome; ";";
            Print #Filnb, Canale(i).UnitaMisura; ";";
            Intero = UnsInt(Canale(i).Bitmin)
            Print #Filnb, Intero; ";";
            Intero = UnsInt(Canale(i).Bitmax)
            Print #Filnb, Intero; ";";
            Print #Filnb, Str(Canale(i).valMin); ";";
            Print #Filnb, Str(Canale(i).valMax); ";";
            Print #Filnb, Str(Canale(i).valOff)
        End If
    Next
    Print #Filnb,
    Print #Filnb, "N.Canale; NomeCanale; Unit� di misura; Fattore"
    Print #Filnb,
    'Trova l'ultimo canale attivo
    iDumm = 0
    For i = 0 To MaxCanali
        'Il canale � attivo?
        If Canale(i).Attivo = True Then
            iDumm = i
        End If
    Next
    
    Print #Filnb, Str(iDumm + 1); ";Pluviometro;mm;"; mmxcount
    Print #Filnb, Str(iDumm + 2); ";Anemometro;m/s;"; msxcount
    
    
    Print #Filnb,
    'Qui si dovrebbe stampare l'intestazione delle colonne
    Print #Filnb, "Data";
    For i = 0 To MaxCanali
        If Canale(i).Attivo = True Then
            Print #Filnb, ";count;";
            Print #Filnb, Trim(Canale(i).Nome);
        End If
    Next
    Print #Filnb, ";count;Pioggia;count;Vento;count;Batteria"
Else
    Stazione = stringC(Stazione, 20)
    Put #Filnb, , Stazione
    lpData = Data2sec70(CDate(PData))
    Put #Filnb, , lpData
    Stringa = Now
    Lungo = Data2sec70((Stringa))
    Put #Filnb, , Lungo
    'Calcolo numero di canali monitorati
    CanaliAttivi = 0
    sCanAttivi = ""
    For i = 0 To MaxCanali
        'Il canale � attivo?
        If Canale(i).Attivo = True Then
            sCanAttivi = sCanAttivi + Format(i, "0#")
            CanaliAttivi = CanaliAttivi + 1
        End If
    Next
    Put #Filnb, , CanaliAttivi
    
    
    'Configurazione dei canali
    For i = 0 To MaxCanali
        'Il canale � attivo?
        If Canale(i).Attivo = True Then
            Stringa = stringC(Canale(i).Nome, 15)
            Put #Filnb, , Stringa
            Stringa = stringC(Canale(i).UnitaMisura, 4)
            Put #Filnb, , Stringa
            Intero = UnsInt(Canale(i).Bitmin)
            Put #Filnb, , Intero
            Intero = UnsInt(Canale(i).Bitmax)
            Put #Filnb, , Intero
            Put #Filnb, , Canale(i).sValmin
            Put #Filnb, , Canale(i).sValmax
            Put #Filnb, , Canale(i).sValoff
            Float = 0
            Put #Filnb, , Float
            Put #Filnb, , Intervallo
            Put #Filnb, , Float
        End If
    Next
'Le seguenti linee sono presenti in MH4 ma non c'erano in
'Simamet. Le aggiungo commentate perch� non so se
'scombinino il simapro.
    'Adesso i dati del canale della batteria
'    Stringa = stringC("Batteria", 16)
'    Put #Filnb, , Stringa
'    Stringa = stringC("Volt", 5)
'    Put #Filnb, , Stringa
'    Intero = UnsInt(0)
'    Put #Filnb, , Intero
'    Intero = UnsInt(4095)
'    Put #Filnb, , Intero
'    Float = 0
'    Put #Filnb, , Float
'    Float = 5 * FattoreBatteriaInterna
'    Put #Filnb, , Float
'    Float = 0
'    Put #Filnb, , Float
'    Put #Filnb, , Float
'    Put #Filnb, , Intervallo
'    Put #Filnb, , Float

End If


'E ora i dati
    ProgressBar1.value = i
    
If bAscii Then
    Tempog = Dat2Ser(CDate(PData))
    'per ogni gruppo di misure prese allo stesso tempo
    'Debug.Print "Iblocco="; iBlocco
    For i = iBlocco To DFPNT Step (2 * (CanaliAttivi + 3))
        ProgressBar1.value = i
        'trasformazione data corrente
        Stringa = CDate(Tempog)
        dTempo = Stringa
        Print #Filnb, dTempo;
        'Print #Filnb, Format(dTempo, "dd/mm/yyyy hh:mm"); ";";
        'per ogni canale
        For j = 1 To CanaliAttivi
            
            'Stabilisco il canale da leggere
            nCanale = j - 1
            'leggo e converto la misura
            Stringa = bMID(BloccoDati, i + (j - 1) * 2, 2)
            'Stringa = SwapString(Stringa)
            Lungo = String2long(Stringa)
            Print #Filnb, ";"; Lungo; ";";
            Stringa = Count2value(nCanale, Lungo)
            'Convertire la misura?
            Print #Filnb, Stringa;
        Next
        j = i + (j - 1) * 2 '(forse+2)
        'leggiamo la pioggia
        Stringa = bMID(BloccoDati, j, 2)
        Lungo = String2long(Stringa)
        Print #Filnb, ";"; Lungo; ";";
        Stringa = Format((Lungo * mmxcount), "#0.00")
        Print #Filnb, Stringa;
        'leggiamo il vento
        Stringa = bMID(BloccoDati, j + 2, 2)
        Lungo = String2long(Stringa)
        Print #Filnb, ";"; Lungo; ";";
        Stringa = Format((Lungo * msxcount), "##0.00")
        Print #Filnb, Stringa;
        'leggiamo la tensione batteria
         Stringa = bMID(BloccoDati, j + 4, 2)
        Lungo = String2long(Stringa)
        Print #Filnb, ";"; Lungo; ";";
        Stringa = Format((CSng(Lungo) / 65535 * 5 * FattoreBatteriaInterna), "#0.00")
        Print #Filnb, Stringa
       
'        j = CanaliAttivi
'        'Stabilisco il canale da leggere
'        nCanale = j - 1
'        'Print #Filnb, Str(nCanale); ",";
'        'leggo e converto la misura
'        stringa = bMID(BloccoDati, i + (j - 1) * 2, 2)
'        'Stringa = SwapString(Stringa)
'        Lungo = String2long(stringa)
'        Print #Filnb, Lungo; ",";
'        Float = sCount2value(nCanale, Lungo)
'        'Convertire la misura?
'        Print #Filnb, Float
'        Print #Filnb,
        Tempog = Tempog + (Intervallo / 86400)
    Next
Else
    Tempog = Dat2Ser(CDate(PData))
    'per ogni gruppo di misure prese allo stesso tempo
    For i = iBlocco To DFPNT Step (2 * CanaliAttivi)
        ProgressBar1.value = i
        'trasformazione data corrente
        Stringa = CDate(Tempog)
        Tempo = Data2sec70((Stringa))
        'per ogni canale
        For j = 1 To CanaliAttivi '* 2 Step 2
            'Stabilisco il canale da leggere
            nCanale = j - 1
            'leggo e converto la misura
            Stringa = bMID(BloccoDati, i + (j - 1) * 2, 2)
            Stringa = SwapString(Stringa)
            Put #Filnb, , Tempo
            Put #Filnb, , nCanale
            Put #Filnb, , Stringa
        Next
'Queste righe sono aggiunte commentate per i motivi esposti
'sopra. Anche queste in MH4 sono presenti
'        'Adesso la batteria
'        j = CanaliAttivi + 1
'        'leggo e converto la misura
'        Stringa = bMID(BloccoDati, i + (j - 1) * 2, 2)
'        Stringa = SwapString(Stringa)
'        nCanale = j - 1
'        Put #Filnb, , Tempo
'        Put #Filnb, , nCanale
'        Put #Filnb, , Stringa
        

        Tempog = Tempog + (Intervallo / 86400)
    Next
End If
    ProgressBar1.value = ProgressBar1.Max
    Label1.Caption = "Processamento terminato"
    Close Filnb
    Call Sleep(250)
    Scaricato = True
    Esci
End Sub

Public Function f4to1(FourthByte As Byte, ThirdByte As Byte, SecondByte As Byte, FirstByte As Byte) As Long
    'Trasforma 4 byte in un long senza segno
    Dim Dummy As Long
    f4to1 = Int(ThirdByte)
    f4to1 = f4to1 * 65536
    Dummy = Int(SecondByte)
    f4to1 = f4to1 + 256 * Dummy
    Dummy = Int(FirstByte)
    f4to1 = f4to1 + Dummy
End Function

Public Function f2to1(SecondByte As Byte, FirstByte As Byte) As Long
    'Trasforma 2 byte in un long senza segno
    f2to1 = Int(SecondByte)
    f2to1 = f2to1 * 256
    f2to1 = f2to1 + FirstByte
End Function

Public Function FormattaData(Data As Variant) As String
    'Trasforma una data Variant in una data formattata in aaaa/mm/gg oo:mm:ss
    FormattaData = Format$(Data, "yyyy/mm/dd")
    FormattaData = FormattaData + " " + Format$(Data, "hh:mm:ss")
End Function

Public Function GetRs(i As Integer) As String
    'Aspetta finch� non arriva un carattere
    Dim Risposta As String
    Dim TimeStop As Long
    TimeStop = Timer + 5 ' Imposta l'ora di fine
    Do
        DoEvents
    Loop Until (fMain.MSComm1.InBufferCount >= i) Or (Timer > TimeStop)
    If fMain.MSComm1.InBufferCount >= i Then
        ' Legge il dato di risposta  sulla porta seriale.
        Risposta = fMain.MSComm1.Input
    Else
        Risposta = "ERRORE"
    End If
    GetRs = Risposta
End Function

Private Function Count2value(i As Byte, Valore As Long) As String
    'Trasforma un valore un count in una misura.
    'Il numero del canale � quello del TFX11 e non di SimaPro
    Dim valore2 As Single
    valore2 = (Valore - Canale(i).Bitmin) / _
    (Canale(i).Bitmax - Canale(i).Bitmin) * _
    (Canale(i).valMax - Canale(i).valMin) + Canale(i).valMin + Canale(i).valOff
    Count2value = valore2
End Function

Private Function sCount2value(i As Byte, Valore As Long) As String
    'Trasforma un valore un count in una misura.
    'Il numero del canale � quello di SimaPro e non del TFX11
    Dim valore2 As Single
    valore2 = (Valore - sCanale(i).Bitmin) / _
    (sCanale(i).Bitmax - sCanale(i).Bitmin) * _
    (sCanale(i).valMax - sCanale(i).valMin) + sCanale(i).valMin + sCanale(i).valOff
    sCount2value = valore2
End Function



