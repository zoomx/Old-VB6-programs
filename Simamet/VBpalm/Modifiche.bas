Attribute VB_Name = "Modifiche"
'Elenco modifiche e aggiunte

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

'2001 04 09
'Nel form fCounter il numero del canale è aumentato di uno.
'Fino adesso i canali, nel file .dat, partivano da zero.

'Versione consegnata a Magest

'2001 09 01
'Nel form fModem evento Load aggiunta la lettura dell'ultima porta COM
'utilizzata e conseguente settaggio degli option oCom

'2001 10 31
'Aggiunte alcune routine da MH42:
'Aggiunto il tipo variabile NUMBERFMT (serve per l'API
'Spostata la variabile SE Separatore Elenco vicino a separatore decimali
' e migliaia
'Aggiunte le variabili globali
'Public nft As NUMBERFMT     'Formato numeri custom
'Public Palm As Boolean	     'Si usa o no il Palm per scaricare
'Public TipoFile As String
'in fmain bScarica cambiata l'istruzione da
'Dummy = Dummy + (Stazione) a
'Dummy = Dummy + Trim(Stazione)
'Cambiata anche la gestione del tipofile (ascii bin excel..)
'Aggiunta la gestione del Palm
'

