Attribute VB_Name = "VariabiliGlobali"
Option Explicit
'Sezione TYPE
'ogni tipo di sensore
Type DatiSensore
    Nome As String * 16    'Nome sensore
    Guadagno As Integer
    UnitaMisura As String
    Volt2Mis As Double
    Bitmin As Long
    Bitmax As Long
    valMin As Single
    valMax As Single
    valOff As Single
End Type

Type DatiCanale
    Nome As String * 16     'Nome sensore
    Attivo As Boolean
    UnitaMisura As String * 5
    Soglia As Single
    Volt2Mis As Double
    Bitmin As Long
    Bitmax As Long
    valMin As Single
    valMax As Single
    valOff As Single
    sValmin As String * 4   'Servono ad evitare di tradurre un
    sValmax As String * 4   'valore single espresso in 4 byte
    sValoff As String * 4   'contenuti in una stringa nel valore single
End Type



'Sezione CONST

'Comandi per il TFX11
Public Const Acquisizione As String = "1" 'Partenza acquisizione
Public Const ScaricoDati As String = "2"  'Scarico dati
Public Const Dormi As String = "3"       'Manda il TFX11 a dormire
Public Const Prova As String = "4"       'Acqisizione di prova
Public Const InfoAcq As String = "8"
Public Const TestSensori As String = "10"   'Test sensori
Public Const InfoProg As String = "11"  'Scarica la programmazione
Public Const OrarioModem As String = "16"   'Invia l'orario accensione e spegnimento modem
Public Const ScaricaOrarioModem As String = "17" 'Scarica l'orario di accensione e spegnimento del modem
Public Const LeggiBattFact As String = "22" 'Legge il fattore conversione tensione batteria
Public Const ScriviBattFact As String = "23" 'Scrive il fattore conversione tensione batteria
Public Const LeggiOrarioTFX As String = "24" 'Legge l'ora dal TFX
Public Const LeggiDFMAX As String = "25"    'Legge DFMAX quantità massima di memoria dati
Public Const continua As String = "50"
Public Const Spunt As String = "90"
Public Const Batteria As String = "94"  'Mostra il livello di tensione della batteria
Public Const Xmit As String = "95"      'manda un Xmit+
Public Const Copyright As String = "96"  'Mostra il copyright
Public Const ScarErr As String = "97"    'Scarico errori
Public Const Scarico_emergenza As String = "98" 'Eventuale scarico di emergenza non utile se si usa l'offload
Public Const StopPrg As String = "99"    'Ferma il programma


'Variabili globali per il programma
Public Const TmOut As Integer = 10 'Timeout comunicazioni
Public Const None As Integer = 0
Public ComPort As Integer
Public PAnno As String
Public PMese As String
Public PGiorno As String
Public POra As String
Public PMinuti As String
Public Programmato As Boolean
Public Scaricato As Boolean    'Serve per sapere se si e' scaricato o meno. Deve essere globale
Public Messaggio As String
Public ComOk As Boolean
Public Collegato As Boolean
Public FileOut As String
Public PathOut As String
Public DriveOut As String
Public comando As String
Public Esci As Boolean
Public DataProgrammazione As Date
'Array di configurazione di tutti i sensori
Public Sensore() As DatiSensore
Public Canale(19) As DatiCanale
Public aCanaliAttivi(19, 2) As Integer  'Serve per costruire la tabella dei canali attivi
                                        'secondo le indicazioni di SimaPro in cui il
                                        'canale 0 non è il canale 0 della centralina ma
                                        'semplicemente il primo canale monitorato.
Public sCanale(19) As DatiCanale        'Copia del vettore canale ma con diverso
                                        'ordine. Serve per calcolare il valore reale
                                        'della misura effettuata visto che non c'e'
                                        'corrispondenza fra il numero di canale sul TFX
                                        'e quello sul SimaPro
Public Const MaxCanali As Integer = 17  'Numero di canali di acquisizione
                                        'C'è anche il canale 0
Public FileName As String
Public Const Vero As Boolean = True
Public Const Falso As Boolean = False
Public Const TestataPrg As String = "Simamet Sensors Setup File"
Public Stazione As String
Public Intervallo As Long  'Intevallo di campionamento in secondi
Public Const MinimoIntervalloAcquisizione As Long = 30
Public Const CorrezioneTempo As Double = -3600 * 1  'Correzione tempo fra VB e SimaPro
Public CTRLC As String
Public Decimale As String
Public Migliaia As String
Public ModemString As String
Public ConnessioneRemota As Boolean
Public ChiamaFlag As Boolean    'Serve per sapere
                                'se il bottone chiama
                                'è stato premuto
Public ProgrammazioneCaricata As Boolean
Public FattoreBatteriaInterna As Single 'Fattore di correzione della batteria interna
                                        'per passare da 0-5 V ai 12v e oltre
                                        'ATTENZIONE! fare in modo che sia identico
                                        'a quello che si trova in CheckBat nel firmware
Public Const MinimaTensioneBatteria As Single = 10.8    'Si riferisce alla batteria della centralina
Public TensioneBatteria As Single   'Tensione batteria della centralina
Public FileIni As String    'Nome file di inizializzazione
Public SE As String         'Separatore di elenco
Public fDebug As Boolean    'Se e' vero stampa sul file di log
Public fdn As Integer       'E' il numero di file del file di log
Public lDebug As Boolean    'Se è vero fa comparire piu' pulsanti e menu speciali
                            'visibili in precedenti versioni e adesso nascosti

Public mmxcount As Single   'mm per count del pluviometro
Public msxcount As Single   'm/s per count dell'anemometro

