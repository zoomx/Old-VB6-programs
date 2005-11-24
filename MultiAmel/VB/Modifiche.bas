Attribute VB_Name = "Modifiche"
'Elenco modifiche e aggiunte

'2000 02 11
'Effettuati vari aggiornamenti presi da MH4fix e MH4ver2
'Le varie funzioni presenti in modulo1.bas
'sono state divise fra vari moduli

'2000 02 14
'Aggiunti i due canali di temperatura interna e
'Tensione batteria.
'In fVislChnl la tensione della batteria è corretta
'in un If Then moltiplicando per un fattore che sta
'tra le variabili globali
'E' stata aggiunta la taratura con un suo form
'La stampa nel form fCounter è stata modificata

'2000 02 15
'Aggiunta la taratura del termometro
'Aggiunte le funzioni che calconano
'm e q di una retta a partire da una coppia
'di punti e quelle che ricavano i
'parametri per simapro a partire da
'm e q
'E' stato aggiunto il form fAttuale che
'adesso serve puramente per prove

'2000 02 16
'correzioni varie  fra cui la funzione str() nei
'print del setup e l'implementazione
'della taratura anche a livello di setup
'Nel form fIntervallo sono stati nascosti
'i controlli relativi ai secondi ormai inutili
'Aggiunta la taratura della conducibilita'
'Ovviamente sono state modificate le formule
'in fCounter poiche' adesso si tiene conto
'dei parametri tipo simapro

'2000 02 18
'Modifiche per il modem
'Aggiunto il form modem
'Allungati i timeout


'2000 02 27
'Gaetano ha risolto il baco del blocco dello scarico dati
'quando la conducibilità è =0 (perchè poi c'è una divisione)
'Tale modifica è stata apportata in fCounter
'   Case 1  'Conducibilita'
'   'Float = Kc / (Float - 0.02)
'   If Float = 0 Then
'       Conducibilita = 0
'   Else
'       Conducibilita = 1 / Float
'   End If

'2000 02 28
'Eliminato baco di incompatibilità fra la versione sotto
'ambiente operativo e quella eseguibile. Quella sotto ambiente
'operativo scrive e legge Vero e Falso mentre l'eseguibile scrive
'e legge True e False
'In frmOptions e frmOptions2 nella routine bSalva_Click è stata
'apportata la seguente modifica
'    For i = 0 To MaxCanali
'        Print #Filnb, Canale(i).Nome
'        Inizio modifica
'        If Canale(i).Attivo = True Then
'            Print #Filnb, "True"
'        Else
'            Print #Filnb, "False"
'        End If
'        Fine modifica al posto della linea seguente
'        'Print #Filnb, Canale(i).Attivo
'In lettura il problema è stato risolto precedentemente
'con l'uso della funzione cbool()

'2000 03 09
'
'In fTaratura tolta la linea segnata
'    Volt7 = 5 / 4095 * Intero
'    If Volt7 = Zero Then
'        Messaggio = "ERRORE! La misura sembra" + vbCrLf
'        Messaggio = Messaggio + "identica alla precedente!"
'        MsgBox (Messaggio)
'        GoTo daccapo
'    End If
'    Text1.Text = Text1.Text + Messaggio
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'    stringa = InputBox("Immetti il pH della soluzione tampone", "Taratura sensore pH", "4")
'    If stringa = "" Then
'        Text1.Text = ""
'        GoTo uscita
'    End If

'Cambiato il valore di FattoreBatteriaInterna da 2.8 a 2.65

'2000 03 13
'Gaetano ha fatto le seguenti modifiche
'in fVislChnl NuoviValori()
'If i = 5 Then Misura = Misura * FattoreBatteriaInterna
'If i = 1 Then Misura = 1 / Misura 'Gae 13mar2000
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'lCanale(i).Caption = Format(Misura, Formato)
'in fTaratura bTarapH
'    If stringa = "TimeOut" Then
'        Errore
'        GoTo uscita
'    End If
'>    If Val(stringa) = 0 Then 'Gae 13mar2000
'>        Errore
'>        GoTo uscita
'>    End If

'    stringa = InputBox("Immetti il pH della soluzione tampone", "Taratura sensore pH", "4")
'    If stringa = "" Then
'        Text1.Text = ""
'        GoTo uscita
'    End If
'>    If Val(stringa) = 0 Then 'Gae 13mar2000
'>        Errore
'>        GoTo uscita
'>    End If
'in fTaratura bTaraT
'    If stringa = "TimeOut" Then
'        Errore
'        GoTo uscita
'    End If
'>    If (stringa) = "" Then 'Gae 13mar2000
'>        Errore
'>        GoTo uscita
'>    End If
'In fTaratura bTaraCond
'    If stringa = "TimeOut" Then
'        Errore
'        GoTo uscita
'    End If
'>    If Val(stringa) = 0 Then 'Gae 13mar2000
'>        Errore
'>        GoTo uscita
'>    End If
'    Label2.Caption = "Taratura effettuata"
'>    Canale(1).UnitaMisura = "mS" 'Gae 13mar2000





'2000 03 21
'Adesso se viene scaricata la programmazione automaticamente
'viene usato come Intervallo di acquisizione l'intervallo della
'precedente programmazione
'Per far questo adesso Intervallo e' una variabile globale
'ed e' stata aggiunta una routine nell'evento Form_load di fIntervallo

'2000 04 12
'Commentati alcuni CloseCom che potevano causare
'la disconnessione del modem
'Adesso le variabili di taratura, l'ultimo numero chiamato
'e i settaggi della porta com con il modem
'vengono memorizzate in un file INI
 
'2000 04 13
'In fVislChnl viene conservata la temperatura misurata
'e viene poi utilizzata per correggere la
'conducibilità
'In fTaratura bTaraCond modificato il messaggio di immissione
'della conducibilità: viene specificato che deve essere
'riferita a 25°C
'
'In frmOptions e frmOptions2 Form_load aggiunta la seguente linea
'    Item = tbsOptions.SelectedItem.Index - 1
'prima di
'    AggiornaTbs (tbsOptions.SelectedItem.Index)
'senza la quale sono possibili scambi di impostazioni
'fra schede diverse. Basta premere su test, cambiare scheda,
'andare avanti, uscire dal test, andare daccapo su test.
'La scheda precedentemente selezionata adesso
'conterrà il setup della prima scheda.
'
'La rubrica è stata corretta. Prima, usando una tabella
'di tipo testo non permetteva di modificare i record inseriti.
'La tabella è stata cambiata in tipo access, è stato aggiunta la
'Microsoft DAO 2.5/3.5 compatibility library e la lettura
'dell'ultimo numero in fModem è stata spostata dall'evento
'form_activate all'evento form_load.

'2000 04 14
'Aggiunto lo scarico della memoria quando DFPNT=0
'Riarrangiamento e aggiornamento dei moduli .bas

'2000 05 11
'In frmOptions e frmOptions2 modificata la numerazione delle
'schede. Adesso invece di partire da 0 partono da 1
'Aggiunto il caricamento ed il salvataggio del setup
'rispettivamente nell'evento load e quit du fTaratura

'2000 05 14
'in fOrarioModem_load aggiunta la linea
'    stringa = InputComTimeOut(Tempo)
'al posto dell'attesa e cancellazione del buffer.
'Ciò per evitare che un'attesa troppo corta faccia apparire
'la scritta Multipar nell'orario
'Aggiunta la visualizzazione dei count in fVislChnl

'2000 05 29
'Aumentati i timeout in fModem_Connetti da 2 a 5
'Aumentato TmOut a 10

'2000 07 05
'Il pulsante Cambia Orario Modem è stato spostato da
'fMAin a fModem
'In fOrarioModem Form_Load aggiunto un controllo sul
'primo campo dell'orario ricevuto per evitare che sia la
'stringa "Multipar"

'2000 07 06
'Aggiunta la funzione CalcolaBattFact() in Funzioni.bas
'Aggiunta la possibilità di immettere la temperatura della
'soluzione tampone a Ph7

'2000 07 19
'In fTaratura commentati i CTRLC perchè quando il
'TFX va al menù principale spegne tutti i sensori.

'2000 07 20
'Eliminate le correzioni in temperatura in fTaratura e fVslChnl
'Aggiunta la lettura del valore FattoreBatteriaInterna
'In fTaratura aggiunto Val2 a tutte le funzioni InputBox che altrimenti
'non tengono conto del punto decimale
'Aggiunta la taratura fattore batteria
'Scambiati i valori di INHOn e INHOff

'2000 07 27
'In fModem corretto il valore di default per le porte COM
'da 57600,8,n,1 aq 57600,n,8,1 e aggiunta la gestione errore
'per settaggi errati
'Il pulsante bOrarioModem è tornato in fMain perchè in
'fModem non poteva funzionare mai

'2000 08 21
'Eliminate alcune dimenticanze:
'- pulsante OrarioModem in fModem
'- ritorno a fMain invece che a fModem in FOrarioModem
'- icona in fOrarioModem
'- option explicit in tutti i moduli e quindi dichiarazione di
'   tutte le variabili
'In FunzioniComuniNumeriche String2single corretto un secondobyite in SecondoByte!!

'2000 08 25
'Modificato frmOptions2
'In Form_Load viene aggiornata la prima tabella prima
'della sua selezione. Se si seleziona prima di aggiornare
'succede che il contenuto della tabella, che all'inizio è
'nullo, viene copiato nei dati del canale 0, cancellandoli.
'L'aggiornamento a video è stato spostato nell'evento
'Form_Paint.
'É stato aggiunto un ulteriore form fPhCond tra frmOptions2 e
'fVislChnl dove si chiede se si vuole misurare il pH
'o la Conducibilità. Poichè così facendo si disattiva uno dei due
'canali all'uscita dal test entrambi vengono riattivati.
'C'è però il remoto pericolo che uno dei due sia stato
'disattivato da programmazione e che quindi venga riattivato.

'2000 09 01
'In fStazione bContinua_Click() aggiunta la linea Stazione = tStazione.Text
'senza la quale il nome stazione rimaneva Multipar.
'In fCounter Esci() eliminata la dichiarazione della variabile
'CRTLC in quanto esistente già come globale
'Il separatore di dati ASCII è adesso una variabile
'stringa globale

'2000 09 06
'Aggiunte delle linee per una eventuale registrazione
'di un log del collegamento

'2000 09 13
'Adesso l'avvio del programma è con parametri
'L'unico parametro è per adesso /lab che se assente
'non vengono visualizzate tutte le opzioni di setup in
'frmOptions e frmOptions2
'In fMain bProgramma aggiunto un DoEvents subito dopo il
'primo DisabTasti perchè altrimenti il successivo comando,
'che provoca il lento caricamento di frmOptions, non permette
'la disabilitazione rapida dei tasti.
'Aggiunta la label lAttendere in fMain che per adesso non
'viene utilizzata.

'2000 09 14
'Aggiunta la routine che controlla se i dati scaricati sono tutti zero.
'Per adesso però non è attiva perchè non ho una centralina per provare.

'2000 09 15
'I parametri non vengono più salvati in MH4p.ini ma in
'MultiparParametri.txt
'Attivata la routine di controllo dei dati scaricati in emergenza.
'La conducibilità adesso è lineare. Sono state disattivate le
'linee con 1/conducibilità in fCounter, fVslChnl e fTaratura. Per
'trovale le linee modificate cercare "Non linearità".

'2000 09 18
'Commentate tutte le linee contenenti
'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
'in quanto su piattaforme multimonitor non funzionano correttamente
'Adesso in fOptions la tStazione.text viene aggiornata con il
'nome stazione eventualmente già scaricato

'2000 09 19
'In fVslChnl attivata l'opzione che visualizza la casella
'con i count solamente nella versione debug
'Aggiunta la Private Sub Form_QueryUnload in fPhCond, fRubrica
'In fCounter aggiunto ProgressBar1.value = ProgressBar1.Max
'nel ciclo Do Loop in Scarica() perchè altrimenti
'rimane bloccato all'inizio in caso di scarico di emergenza.
'Sempre in fCounter cambiata la memorizzazione del file
'binario da
'For i = 0 To UBound(BloccoDati)
'    Put #Filnb, , BloccoDati(i)
'    'DoEvents
'    Next
'
'a
'
'Put #Filnb, , BloccoDati()
'
'molto più efficiente
'
'Corretto il controllo in caso di scarico con memoria
'vuota in fCounter
'Aggiunte le costanti  LeggiOrarioTFX e LeggiDFMAX per futuri
'miglioramenti

'2000 09 20
'In fIntervallo cambiato il controllo per il minimo
'intervallo di acquisizione. Prima era 30 secondi adesso è
'1 minuto. Se l'intervallo di campionamento era 0 il
'programma automaticamente lo impostava su 30 secondi che
'non erano visibili (il relativo controllo text è nascosto).
'Ciò provocava un indesiderato aumento del tempo di campionamento
'di 30 secondi, se l'utente cambiava il tempo impostato
'(ad esempio impostando 1 minuto otteneva 1 minuto e 30 secondi)
'o la partenza con un tempo che sembrava impostato a zero.

'2000 09 21
'Poichè nel firmware il controllo della tensione di
'alimentazione è stato ripristinato capitava che se
'si tentava di programmare la centralina con tensione
'inferiore al livello di guardia non si riuscisse
'nell'intento ottenendo un errore di mancata risposta
'"La centralina Multipar non risponde al CTRL C --"
'Questo perchè la centralina, al riavvio dopo
'la cancellazione della memoria, andava nella modalità
'di basso consumo (dormi) dalla quale per essere svegliata
'occorrono un paio di CTRL+C e non uno solo.
'La soluzione adottata è stata quella di controllare
'la tensione prima della programmazione. Ciò è stato fatto
'in fMain bProgramma_Click() all'inizio della sub
'
'in fCounter cambiato
'    Print #Filnb, "Tensione batteria "; fMain.StatusBar1.Panels(1).Text; " volt"
'in
'    Print #Filnb, "Tensione batteria "; fMain.StatusBar1.Panels(1).Text;
'
'In fIntervallo bContinua_Click() aggiunto
'    'Attivo tutti i canali perchè da
'    'qualche parte il canale 0 viene disattivato
'    For i = 0 To 5
'        Canale(i).Attivo = True
'    Next
'subito prima di mandare la configurazione dei canali
'
'Spostata l'istruzione
'Set frmOptions.tbsOptions.SelectedItem = frmOptions.tbsOptions.Tabs(1)
'da fMain bProgramma_Click() a frmOptions Form_Load()
'Questo evita il problema che il primo canale viene disattivato.

'2000 09 22
'Risolti (forse) i problemi creatisi con la disattivazione dei
'tabs in frmOptions e frmOptions2. Il problema principale era
'che in alcune condizioni il primo tabs veniva azzerato
'oppure sovrascritto con dati più vecchi
'Attivata l'opzione /debug che scrive su un file di
'log tutti i problemi. Per attivarla bisogna eseguire
'le stesse operazioni come per l'opzione /lab


'sincronizzazione orario
'In tara pare che il ph si spenga

'2001 02 14
'Corretto baco nelle tarature: dopo i calcoli aggiunte le
'istruzioni
'    Canale(x).Bitmin = q
'    Canale(x).Bitmax = 4095
'per temperatura, pH e conducibilità

'2001 02 15
'Aggiunto la generazione del nome del file dove salvare
'i dati secondo la data

'2001 03 21
'Aggiunta la registrazione e la lettura del percorso
'dell'ultimo file di programmazione usato
'Riattivata la visualizzazione delle schede

'2001 05 18
'Cambiati i TabIndex in frmOptions e frmOptions2
'In fTaratura cambiato il valore da inserire della tensione
'batteria da 12.5 al valore effettivamente misurato.
'Sempre in fTaratura tutte le variabili single di bTaraBatt_Click()
'sono diventate double perchè sembra che il sistema non funzionava
'In Trovaretta (FunzioniComuniNumerice.bas) tutti i single sono
'stati cambiati in double.
'In fTaratura cambiata
'    TrovaBitVal m, q, CSng(Canale(0).Bitmin), _
'    CSng(Canale(0).Bitmax), valMin, valMax, _
'    valOff
'
'in
'
'    TrovaBitVal m, q, 0, 4095, valMin, valMax, 0
'e Canale(3).Bitmin = q in Canale(3).Bitmin = 0
'per Conducibilità e pH
'In fCom la COM1 è stata preselezionata,
'il tasto Ok è stato abilitato ed è stato eliminato
'un loop che si innesca quando nel form fCom viene
'cliccata la x per chiudere il form. Inoltre è
'stato cambiato l'ordinamento del tab
'In frmOptions2 è stato cancellato il tasto di prova p
'In fIntervallo l'intervallo è predisposto su 2 ore ed è
'stato cambiato l'ordine dei TAB
'In fOrario è stato cambiato l'ordine dei TAB
'In fOrarioModem la proprietà MaxLength è stata
'cambiata da zero a 2. TabIndex modificati
'In fPhCond è stata cambiata la dimensione del font
'e corretta la scritta da Ph a pH
'Aggiunta la variabile globale TipoFile che indica
'il tipo di file di dati da salvare.
'In fCounter cambiati gli if then tra ascii e bin in
'select Case

'2001 05 24
'In fTara eliminate le diciture sull'asse X
'Risistemate le icone in fMain e aggiunta l'icona al tasto
'Tara2
'Trasformazione da versione Poseidon a versione Mista Poseidon/Multipar
'La costante TestataPrg diventa variabile e viene inizializzata nella
'Main() e cambiati tutti i messaggi
'Create 2 variabili globali InitDirPrg e InitDirSave
'La versione Poseidon allinea le ore della programmazione
'a quelle pari
'In fMain aggiunto il tasto bRemota in AbilitaTasti e DisabTasti
'Sempre in fMain aggiunta la scritta attendere quando si
'cancellano i dati
'Nella versione Poseidon vengono caricate a runtime le nuove icone

'2001 05 26
'In fMain eliminata la procedura bInvia in quanto non più utilizzata

'2002 02 25
'Corretto tbsOptions_Click in fOptions che nella versione Poseidon
'evitava lo scorrimento dei tabs
'Corretto in fOption lo ShowInTaskBar

'2002 03 18
'Aggiunta la variabile public LastFileSaved e il  pulsante per aprirlo.
'Aggiunta la spedizione di un CTRL+R in caso di fallimento della comunicazione
'per mancanza di risposta dalla centralina

'2002 03 19
'Corretto il CTRL+R da 19 a 18

'2002 03 22
'Cambiata la routine di cancellazione memoria, adesso si controllano le
'risposte. Cambiato l'aggiornamento automatico della data di partenza. Prima
'era allineata alle due ore, adesso è stata portata all'ora.

'2002 03 25
'Adesso la routine di cancellazione memoria sembra funzionare

'2002 09 03
'Cambiate definitivamente le icone
'Aggiunto il pulsante che manda un CTRL+R
'Nel Main Form scambiate le posizioni dei pulsanti scarica e programma

'2002 11 12
'Attivato il pulsante per la visione dell'ultimo file salvato

'2002 12 02
'Attivati tutti i pulsanti di fTara che prima erano disattivati
'Aggiunto il pulsante per il CTRLC e Expert mode
'Alcune scritte passate da normale 8 a grassetto 10
'L'intervallo di acquisizione standard in fIntervallo è adesso di un'ora.

'2003 02 25
'In fIntervallo cambiate alcune righe per implementare InputComTimeOut(2)

'2003 03 04
'Aggiunto il form frmTerminal
'Adesso c'è un terminale

'2003 03 05
'Modificati fIntervallo.bContinua_Click e fModem.Connetti
'in modo che vengano accettate anche risposte parzialmente errate come Pseidon

'2003 05 28
'Corretta bLastFile_Click() in fMain
'sostituita
'    Shell LastFileSaved, vbNormalFocus
'con
'    Stringa = "notepad " + LastFileSaved
'    Shell Stringa, vbNormalFocus
'Prima in notepad non partiva e il programma si inchiodava

'2003 06 26
'Aggiunta la gestione delle COM fino a 4 per
'la connessione via cavo
'Modificato il layout di fCom, fIntervallo, fTara
'In fOrario corretta la formula che programma la partenza alla prossima ora intera

'2003 08 06
'Corretto baco in Fmodem.frm. Il programma in mancanza del file ini si inchiodava
'Aggiunta la memorizzazione dell'ultima COM usata con il cavo. Modifiche
'in fCom e fMain.

'2003 08 07
'Aggiunta una modifica in fMain.CancellaFlash per gestire la
'mancata cancellazione della flash del TFX11


'2005 04 27
'Cambiate le scritte da Poseidon a MultiAmel
'Cambiate le routine di scarico per interpretare i dati sonda Amel
