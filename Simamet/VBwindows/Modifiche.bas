Attribute VB_Name = "Modifiche"
'Elenco modifiche e aggiunte

'2002 10 08
'In fCounter nell'intestazione dei dati Str(j)è stato cambiato
'in Trim(Str(j))
'Sempre in fCounter è stato modificato il ciclo di
'attesa per lo scarico dei dati da
'    If TimeOuts > 10 Then Exit Do
'a   If TimeOuts > 10 Or Bytes >= DFPNT Then Exit Do
'Modifica lasciata commentata perchè prima bisogna
'testarla.


'2002 09 30
'Allungati i tempi per lo scarico dei dati


'2002 09 09
'In fCounter nell'intestazione dei dati è stato aggiunto
'un contatore per i campi dei gruppi di misure

'2002 07 02
'Cambiato lo scarico dati per l'ennesima volta per
'riflettere le modifiche fatte nell'acquisizione.
'Ad esempio le misure sono fatte in gruppi di 5 invece di 16
'Aggiunta la nuova routine CountDec2value dove il valore
'passato è un single invece che un long

'2002 05 16
'Cambiato lo scarico dei dati che riflette il nuovo
'formato in cui la misurazione del vento è effettuata
'due volte con due sistemi diversi.
'Il significato di msxcount passa da m/s per impulso
'a Hz per m/s
'Cambiata l'intestazione del file di setup da
'"Simamet Sensors Setup File" a
'"Simamet3 Sensors Setup File"
'Adesso le misure sono in gruppi


'2001 01 22
'Il separatore adesso è il ; invece della ,

'2001 01 18
'Aggiunte i valori mmxcount e msxcount x il pluviometro
'e l'anemometro. Tali valori vengono salvati in coda al
'file di programmazione
'In questa versione il modem è stato diasabilitato
'Il pulsante tarabatteria appare solamente con l'opzione /lab
'I canali visibili sono solamente 8


'2001 01 11
'Creazione del programma da MH4_2 del 2001 01 11

