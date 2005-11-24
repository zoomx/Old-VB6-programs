Attribute VB_Name = "FunzioniComuni"
Option Explicit
'PRINT PREVIEW articolo Q193379.html

Sub ShadeForm(Frm As Form)
' Description
'     Draws the "install-type" shaded background on a form
'
' Paramaters
'     Name                 Type     Value
'     -----------------------------------------------------------
'     Frm                  Form     The form to draw the shade on
'
' Returns
'     Nothing
'
' Last modified by Gord MacLeod 05.02.96

Dim i%
Dim NumberOfRects As Integer
Dim GradColor As Long
Dim GradValue As Integer

   Frm.ScaleMode = 3
   Frm.DrawStyle = 6
   Frm.DrawWidth = 2
   Frm.AutoRedraw = True

   NumberOfRects = 64
   
   'For i% = 64 To 1 Step -1
   For i% = 1 To 64
      GradValue = 255 - (i% * 4 - 1)
      ' Put GradValue in Red and Green for a different look
      GradColor = RGB(0, 0, GradValue)
      ' Draw the line
      Frm.Line (0, Frm.ScaleHeight * (i% - 1) / 64)-(Frm.ScaleWidth, Frm.ScaleHeight * i% / 64), GradColor, BF
   Next i%

   Frm.Refresh
End Sub

Sub Gradient(TheObject As Object, Redval&, Greenval&, Blueval&, TopToBottom As Boolean)
    'by John Rogers, June 19, 1996
    'Gradient Me, 0, 0, 255, 1
    'TheObject can be any object that supports the Line method (like forms and pictures).
    'Redval, Greenval, and Blueval are the Red, Green, and Blue starting values from 0 to 255.
    'TopToBottom determines whether the gradient will draw down or up.
    Dim Step%, Reps%, FillTop%, FillLeft%, FillRight%, FillBottom%, HColor$
    'This will create 63 steps in the gradient. This looks smooth on 16-bit and 24-bit color.
    'You can change this, but be careful. You can do some strange-looking stuff with it...
    Step = (TheObject.Height / 63)
    'This tells it whether to start on the top or the bottom and adjusts variables accordingly.
    If TopToBottom = True Then FillTop = 0 Else FillTop = TheObject.Height - Step
    FillLeft = 0
    FillRight = TheObject.Width
    FillBottom = FillTop + Step
    'If you changed the number of steps, change the number of reps to match it.
    'If you don't, the gradient will look all funny.
    For Reps = 1 To 63
        'This draws the colored bar.
        TheObject.Line (FillLeft, FillTop)-(FillRight, FillBottom), RGB(Redval, Greenval, Blueval), BF
        'This decreases the RGB values to darken the color.
        'Lower the value for "squished" gradients. Raise it for incomplete gradients.
        'Also, if you change the number of steps, you will need to change this number.
        Redval = Redval - 4
        Greenval = Greenval - 4
        Blueval = Blueval - 4
        'This prevents the RGB values from becoming negative, which causes a runtime error.
        If Redval <= 0 Then Redval = 0
        If Greenval <= 0 Then Greenval = 0
        If Blueval <= 0 Then Blueval = 0
        'More top or bottom stuff; Moves to next bar.
        If TopToBottom = True Then FillTop = FillBottom Else FillTop = FillTop - Step
        FillBottom = FillTop + Step
    Next
End Sub

Sub Sleeps(seconds As Double)
   'Wait Seconds seconds
   'There is a control for midnight
   Dim TempTime As Double
   TempTime = Timer
   While Timer - TempTime < seconds
      DoEvents
      If Timer < TempTime Then
         TempTime = TempTime - 24# * 3600#
      End If
   Wend
End Sub
'* Use a timer that has greater resolution (generally 1 millisecond).
'  Some of the other timers have values down to 1 millisecond, but you
'  can 't get the precise 1 millisecond resolution.
'
'   'declare
'   Declare Function timeGetTime Lib "MMSYSTEM" () As Long
'
'   'example
'   oldtime& = timeGetTime()
'
'    'code in here
'
'   deltamillisec& = timeGetTime() - oldtime&

Public Sub UnloadAllForms(sFormName As String)
'Unloading All Forms
'There has been a lot of stories about how Visual Basic
'doesn 't unload the forms when you exit the program. This
'is a 'resource killer'.
'This code unloads all of the forms in your program.
'This is a sub, that you would probably use from the
'Form_Unload of your Main form. So here is the code for
'that:
'
'Call UnloadAllForms Me.Name
'
'Also, here is the code if you're calling it from other
'Subs:
'
'Call UnloadAllForms ""

Dim Form As Form
   For Each Form In Forms
      If Form.Name <> sFormName Then
         Unload Form
         Set Form = Nothing
      End If
   Next Form
End Sub

Public Function OpenFile(File2Open As String, FileMode As String) _
     As Integer

'Then there's opening text files. No need to check if it exists or whatever - just call OpenFile with the
'right parameters (Thandle=OpenFile("TempFile","O") for example) and it will do all the error
'checking for you, passing back the file handle if OK, zero if not

     Dim WhatHandle As Integer
     On Local Error GoTo Op_Error
     WhatHandle = FreeFile()

     Select Case FileMode
     Case "I"
     Open File2Open For Input As WhatHandle
     Case "O"
     Open File2Open For Output As WhatHandle
     Case "A"
     Open File2Open For Append As WhatHandle
     Case "B"
     Open File2Open For Binary As WhatHandle
     End Select

     OpenFile = WhatHandle
     Exit Function

Op_Error:
     OpenFile = 0
End Function

Sub SeparaParole(testo$, parole$(), numParole)
    '===========================================================
    ' Separa le parole in una riga di testo
    '===========================================================
    Dim indice, codAscii, inizioParola, Separatore As Integer
    numParole = 0

    For indice = 1 To Len(testo$)
        ' determina se si tratta di un carattere separatore
        codAscii = Asc(Mid$(testo$, indice, 1))
        Select Case codAscii
            Case 48 To 57, 65 To 90, 97 To 122, 128 To 168, 224 To 238
                ' cifre, lettere, accentate e caratteri stranieri
                Separatore = 0
            Case Else
                ' tutto il resto può essere considerato un separatore
                Separatore = -1
        End Select

        If Separatore = 0 Then
            ' se non e' separatore potrebbe essere inizio di parola
            If inizioParola = 0 Then inizioParola = indice
        ElseIf inizioParola > 0 Then
            ' se e' un separatore che segue un non-separatore
            ' aggiungiamo la parola trovata al vettore
            numParole = numParole + 1
            parole$(numParole) = Mid$(testo$, inizioParola, indice - inizioParola)
            inizioParola = 0
        End If
    Next

    ' se l'ultimo carattere non era un separatore
    ' occorre aggiungere l'ultima parola del testo
    If inizioParola > 0 Then
        numParole = numParole + 1
        parole$(numParole) = Mid$(testo$, inizioParola)
    End If
End Sub

Private Sub Form_Unload()
'In applicazioni che includono più form, è possibile
'inserire il codice nella routine dell'evento Unload
'del form principale e utilizzare l'insieme Forms in
'modo che tutti i form vengano individuati e chiusi.
'Nel codice seguente per scaricare tutti i form viene
'utilizzato l'insieme Forms:

'In determinate situazioni potrebbe essere necessario
'terminare l'applicazione indipendentemente dallo stato
'dei form o degli oggetti esistenti.
'A tale scopo è possibile utilizzare l'istruzione End
'che consente di terminare un'applicazione immediatamente.
'Dopo l'esecuzione dell'istruzione End non viene eseguito
'più alcun codice e non viene generato alcun evento,
'in particolare non vengono eseguite le routine di eventi
'QueryUnload, Unload e Terminate. I riferimenti agli
'oggetti vengono liberati, ma se sono state definite
'classi personalizzate, gli eventi Terminate degli
'oggetti creati in base a tali classi non verranno
'generati.

    Dim i As Integer
    ' Esegue un ciclo nell'insieme Forms e scarica
    ' tutti i form.
    For i = 0 To Forms.count - 1
        Unload Forms(i)
    Next
End Sub

Public Function GetDecimal() As String
'Restituisce il separatore decimale
'C'e' anche la API per leggere direttamente dal registro
'di configurazione ma la stringa esiste solamente
'se si modofocano i valori standard
'La API è commentata perchè sembra che in win98 non funzioni sempre
    Dim Decimale As String
    'Decimale = QueryValue(HKEY_USERS, ".Default\Control Panel\International", "sDecimal")
    If Decimale <> "" Then
        Decimale = Left(Decimale, Len(Decimale) - 1)
    Else
        Decimale = Mid(Format(0.5, "0.0"), 2, 1)
    End If
    GetDecimal = Decimale
End Function

Public Function GetMigliaia() As String
'Restituisce il separatore delle migliaia
'C'e' anche la API per leggere direttamente dal registro
'di configurazione ma la stringa esiste solamente
'se si modofocano i valori standard
'La API è commentata perchè sembra che in win98 non funzioni sempre
    Dim Migliaia As String
    'Migliaia = QueryValue(HKEY_USERS, ".Default\Control Panel\International", "sThousand")
    If Migliaia <> "" Then
        Migliaia = Left(Migliaia, Len(Migliaia) - 1)
    Else
        Migliaia = Mid(Format(1000, "#,###"), 2, 1)
    End If

End Function

Public Sub NewPath(Stringa As String)
'Cambia drive e path contemporaneamente
'Modificare per i drive di rete
'Es. NewPath "d:\temp"
    ChDrive (Left(Stringa, 3))
    ChDir (Stringa)
End Sub

Public Function bMID(matrice() As Byte, inizio As Long, lunghezza As Long) As String
'Estrae una stringa da un vettore di bytes
'Sintassi come istruzione MID
    Dim Stringa As String
    Dim i As Long

    For i = inizio To inizio + lunghezza - 1
        Stringa = Stringa + Chr(matrice(i))
    Next
    bMID = Stringa
End Function

Public Function Dat2Ser2(d As String) As Double
    'Trasforma una data in una stringa in numero seriale
    Dim dat As String
    Dim tim As String
    Dim tims As Single
    dat = DateValue(d)
    tim = TimeValue(d)
    'tims = TimeSerial(Hour(tim), Minute(tim), Second(tim))
    'Debug.Print dat, tim, tims
    Dat2Ser2 = DateSerial(Year(dat), Month(dat), Day(dat)) + TimeSerial(Hour(tim), Minute(tim), Second(tim))
End Function

Public Function Dat2Ser(d As Date) As Double
    'Trasforma una data in numero seriale
    Dat2Ser = DateSerial(Year(d), Month(d), Day(d)) + TimeSerial(Hour(d), Minute(d), Second(d))
End Function

Public Function SwapString(Stringa As String) As String
    Dim lStringa As Long
    Dim Dummy As String
    Dim i As Long
    lStringa = Len(Stringa)
    'Capovolge la stringa
    Dummy = ""
    For i = lStringa To 1 Step -1
        Dummy = Dummy + Mid(Stringa, i, 1)
    Next
    SwapString = Dummy
End Function

Public Sub StampaAscii(Stringa As String)
    'Stampa il valore dei caratteri ASCII di una stringa
    'nella finestra di Debug
    Dim lStringa As Double
    Dim i As Integer
    lStringa = Len(Stringa)
    If lStringa = 0 Then Exit Sub
    Debug.Print "Risposta"; Stringa; " ";
    For i = 1 To lStringa
        Debug.Print Asc(Mid(Stringa, i, 1));
    Next
    Debug.Print
End Sub

Public Function String2Ascii(Stringa As String) As String
'Converte una stringa nei corrispondenti valori ASCII
'Non viene gestito il CHR$(0)
    Dim lStringa As Double
    Dim i As Integer
    Dim StringAscii As String
    lStringa = Len(Stringa)
    If lStringa = 0 Then Exit Function
    For i = 1 To lStringa
        String2Ascii = String2Ascii + Asc(Mid(Stringa, i, 1)) + " "
    Next
End Function

Public Function Char2ascii(Stringa As String) As String
'Trasforma una stringa contenente caratteri ASCII e non
'ASCII in stringa di codici di caratteri ASCII
'Viene gestito anche il chr$(0)
    Dim lStringa As Integer
    Dim tStringa As String
    Dim i As Integer
    
    lStringa = Len(Stringa)
    For i = 1 To lStringa
        If Mid(Stringa, i, 1) = Chr$(0) Then
            tStringa = tStringa + " " + "00"
        Else
            tStringa = tStringa + Str(Asc(Mid(Stringa, i, 1)))
        End If
    Next
    Char2ascii = tStringa
End Function

Public Function CeSpazio(Percorso As String, Nbytes As Long) As Boolean
    Dim iUnita As Integer
    Dim ok As Integer
    Dim BytesLiberi As Long
    Dim ClustersLiberi As Long
    Dim ClustersRichiesti As Long
    Dim Unita As String
    Dim SectorsPerCluster As Long
    Dim BytesPerSector As Long
    Dim NumberOfFreeClusters As Long
    Dim TtoalNumberOfClusters As Long
    Dim VaBene As Boolean
    VaBene = False
    'Identifichiamo l'unità
    'Cerchiamo i :
    iUnita = InStr(Percorso, ":")
    'Prendiamo la lettera prima dei :
    Unita = Mid(Percorso, iUnita - 1, 1) + ":\"
    
    ok = GetDiskFreeSpace(Unita, SectorsPerCluster, _
    BytesPerSector, NumberOfFreeClusters, _
    TtoalNumberOfClusters)
    If ok = 0 Then
        CeSpazio = False
        Exit Function
    End If
    BytesLiberi = NumberOfFreeClusters * SectorsPerCluster * BytesPerSector
    ClustersLiberi = NumberOfFreeClusters * SectorsPerCluster
    ClustersRichiesti = Nbytes / SectorsPerCluster / BytesPerSector
    If ClustersRichiesti > ClustersLiberi Then
        VaBene = False
    Else
        VaBene = True
    End If
    CeSpazio = VaBene
End Function

Public Function stripCrLf(Stringa As String) As String
'elimina i Cr e Lf finali in una stringa
    Dim i As Long
    
    For i = 1 To 2
        If Right(Stringa, 1) = vbCr Or Right(Stringa, 1) = vbLf Then
            Stringa = Left(Stringa, Len(Stringa) - 1)
        End If
    Next
    
    stripCrLf = Stringa
End Function

Public Sub FinePerErrore()
    Dim Mes As String
    CloseCom
    Mes = "Errore interno del programma " + App.Title
    Mes = Mes + Str$(Err.Number) + " " + Err.Description
    MsgBox (Mes)
    'Scaricare tutti i forms
    End
End Sub

Public Sub ErrHandler()
'Gestione errore non altrove gestito
    Dim NomeFileErrors As String
    Dim nfile As Integer
    NomeFileErrors = sGetAppPath() + "Errors.log"
    nfile = FreeFile
    Open NomeFileErrors For Append As nfile
    Print #nfile, "Errore in " + App.Title + " del "; Date$, " alle "; Time$
    Print #nfile, "numero "; Err.Number
    Print #nfile, Err.Description
    Print #nfile, Err.Source
    Print #nfile, "applicazione terminata"
    Close nfile
    NomeFileErrors = "Errore nel'applicazione " + App.Title + vbCrLf
    NomeFileErrors = NomeFileErrors + Str(Err.Number) + " " + Err.Description + vbCrLf
    NomeFileErrors = NomeFileErrors + "L'errore è stato salvato nel file errors.log" + vbCrLf
    NomeFileErrors = NomeFileErrors + "L'applicazione verrà chiusa"
    MsgBox (NomeFileErrors)
    'chiude tutti i forms e termina l'applicazione
    'Form_Unload
    End

End Sub

Public Sub ScriviErrore(Errore As String)
'Scrive un errore generico sul file errors.log
    Dim NomeFileErrors As String
    Dim nfile As Integer
    NomeFileErrors = sGetAppPath() + "Errors.log"
    nfile = FreeFile
    Open NomeFileErrors For Append As nfile
    Print #nfile, "Errore in "; App.Title; " del "; Date$; " alle "; Time$
    Print #nfile, Err.Description
    Print #nfile, Err.Source
    Print #nfile, "applicazione terminata"
    Close nfile

End Sub

Function sGetAppPath() As String
'*Returns the application path with a trailing \.      *
'*To use, call the function [SomeString=sGetAppPath()] *
Dim sTemp As String
        sTemp = App.Path
        If Right$(sTemp, 1) <> "\" Then sTemp = sTemp + "\"
        sGetAppPath = sTemp
End Function

Sub CenterForm(frmIn As Form)
'*This code will center a form in the center of the    *
'*screen. To use it, just call the sub and pass it the *
'*form name [Call CenterForm main]                     *
Dim iTop As Integer, iLeft As Integer
    If frmIn.WindowState <> 0 Then Exit Sub
    iTop = (Screen.Height - frmIn.Height) \ 2
    iLeft = (Screen.Width - frmIn.Width) \ 2
    frmIn.Move iLeft, iTop
End Sub

Public Function stringC(Stringa As String, lStringa As Integer) As String
    'Trasforma una stringa in una stringa di lunghezza lstringa
    'riempita da spazi e terminata con chr$(0)
    Dim lStringa2 As Integer
    lStringa2 = Len(Stringa)
    If lStringa2 < lStringa Then
        Stringa = Stringa + Space(lStringa - lStringa2 - 1)
    Else
        Stringa = Mid(Stringa, 1, lStringa)
    End If
    stringC = Stringa + Chr(0)
End Function

Public Function Formato(Numero As Double, StringaFormato As String) As String
'Modifica dell'istruzione format
'Tiene conto del fatto che l'istruzione Format mette come
'separatore dei decimali cio' che gli viene indicato
'dalle impostazioni internazionali. Quindi se il separatore non
'è un punto elimina il separatore e ci mette il punto.
'In pratica sostituisce la virgola col punto
'Richiede che in Migliaia e Decimale ci siano i separatori
'corrispondenti, letti dal registro di configurazione
'Si potrebbe evitare di sapere preventivamente di conoscere
'Il separatore cercando la virgola ma questa potrebbe
'essere presente come separatore delle migliaia

    Dim i As Integer
    Dim LungString2 As Integer
    Dim Stringa2 As String
    Stringa2 = Format(Numero, StringaFormato)
    LungString2 = Len(Stringa2)
    'il decimale e' un punto o una virgola?
    If Decimale <> "." Then
        'Si, sostituiamolo con il punto
        i = InStr(Stringa2, Decimale)
        Stringa2 = Left(Stringa2, i - 1) + "." + Right(Stringa2, LungString2 - i)
    End If
    Formato = Stringa2
End Function


Public Function SetInIDE() As Boolean
'Restituisce True (Vero) se si è in ambiente di programmazione
'False se il programma è compilato
    On Error GoTo DivideError
    Debug.Print 1 / 0
    SetInIDE = False
    Exit Function
    
DivideError:
    SetInIDE = True
End Function
