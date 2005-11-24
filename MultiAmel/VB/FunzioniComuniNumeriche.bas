Attribute VB_Name = "FunzioniComuniNumeriche"
Option Explicit
Public Function String2single(Stringa As String) As Single
    'Converte una stringa rappresentante un Float in standard
    'IEEE 754 32 bit in una variabile single (float in C)
    Dim PrimoByte As Byte
    Dim SecondoByte As Byte
    Dim TerzoByte As Byte
    Dim QuartoByte As Byte
    
    Dim Segno As Integer
    Dim Mantissa As Long
    Dim Esponente As Integer
    Dim dMantissa As Single
    Dim Risultato As Single
    Dim Risultato2 As Single
    Dim dMantissa2 As Single

    On Error GoTo GestErr

    PrimoByte = Asc(Right(Stringa, 1))
    SecondoByte = Asc(Mid(Stringa, 3, 1))
    TerzoByte = Asc(Mid(Stringa, 2, 1))
    QuartoByte = Asc(Left(Stringa, 1))
    
    '(-1)**S * 2**(E-Bias) * 1.F
       
    'caso 0
    If PrimoByte = 0 And SecondoByte = 0 And TerzoByte = 0 And QuartoByte = 0 Then
        String2single = 0
        Exit Function
    End If
    
    Segno = PrimoByte And 128
    If Segno = 128 Then
        Segno = -1
    Else
        Segno = 1
    End If
    
    Esponente = (PrimoByte And 127) * 2 + (SecondoByte And 128) / 128
    Esponente = Esponente - 127
    
    Mantissa = SecondoByte And 127
    Mantissa = Mantissa * 65535
    Mantissa = Mantissa + CLng(TerzoByte) * 256
    Mantissa = Mantissa + QuartoByte
    dMantissa = Mantissa / 8388480 + 1
    
    Mantissa = (SecondoByte And 127) + 1
    Mantissa = Mantissa * 65535
    Mantissa = Mantissa + CLng(TerzoByte) * 256
    Mantissa = Mantissa + QuartoByte
    
    
    Risultato = dMantissa * 2 ^ Esponente * Segno
    '(-1)**S * 2**(E-Bias) * 1.F
    Risultato2 = -1 ^ (Segno) * 2 ^ Esponente * Mantissa

    String2single = Risultato
    Exit Function
GestErr:
    If Err.Number = 6 Then
        String2single = 3.402823E+38
    End If

End Function

Public Function UnsInt(Lungo As Long) As Integer
    'trasforma un unsigned int (rappresentato in un long)
    'in un signed int, che è l'unico accettato da VB
    If Lungo > 32767 Then Lungo = Lungo - 65536
    UnsInt = CInt(Lungo)
    'Debug.Print UnsInt
End Function

Public Sub Int2Bytes(i As Integer, bh As Byte, bl As Byte)
    bh = Int(i / 256)
    bl = i Mod 256
    If bh * 256 + bl <> i Then
        MsgBox ("Errore nella funzione Int2Bytes")
        Debug.Print i, bh, bl
    End If
End Sub

Public Function bytes2long(Stringa As String) As Long
'converte una stringa rappresentante un numero long in binario
'(littel endian, basso-alto) nel numero stesso
    Dim lStringa As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Lungo As Long
    Dim a As String
    On Error GoTo GestErr
    'StampaAscii (Stringa)
    lStringa = Len(Stringa)
    If lStringa > 4 Then
        Stringa = Left(Stringa, 4)
        lStringa = Len(Stringa)
    End If
    Lungo = 0
    'For i = lstringa To 1 Step -1
    For i = 1 To lStringa
        a = Mid(Stringa, i, 1)
        j = Asc(a)
        Lungo = Lungo + j * 256 ^ (i - 1)
        
    Next
    bytes2long = Lungo
    Exit Function
GestErr:
    If Err.Number = 6 Then
        bytes2long = 2147483647
    End If
    
End Function

Public Function Val2(Valore As Variant) As Variant
'Simile alla val ma per separatore decimale usa sia il
'punto che la virgola
    Dim ip As Integer
    Dim iv As Integer
    Dim lStringa As Integer
    Dim temp As Variant
    Dim Stringa As String
    
    Stringa = CStr(Valore)
    'C'è il punto?
    ip = InStr(Stringa, ".")
    'C'è la virgola?
    iv = InStr(Stringa, ",")
    lStringa = Len(Stringa)
    If iv <> 0 Then 'Se c'è la virgola la sostituisce col punto
        Stringa = Left(Stringa, iv - 1) + "." + Right(Stringa, lStringa - iv)
        ip = iv
    End If
    temp = Val(Stringa)
    'If ip <> 0 And iv <> 0 Then
    'Se ci sono tutte e due?
    Val2 = temp
End Function

Function SwapBytes(num As Integer) As Integer
' Take an input integer, assumed to be in "left to right" byte order, and convert it to "standard" Intel format by swapping the two bytes.

Dim TextVal As String
Dim NewTextVal As String
Dim StringLength As Integer

TextVal = Hex$(num)
StringLength = Len(TextVal)
Select Case StringLength
Case 1
   NewTextVal = "&H" & "0" & TextVal & "00"
Case 2
   NewTextVal = "&H" & TextVal & "00"
Case 3
   NewTextVal = "&H" & Right$(TextVal, 2) & "0" & Left$(TextVal, 1)
Case 4
   NewTextVal = "&H" & Right$(TextVal, 2) & Left$(TextVal, 2)
End Select
SwapBytes = Val(NewTextVal)
End Function

Public Function Count2value(i As Byte, Valore As Long) As String
'ATTENZIONE: aggiunto + Canale(i).Valmin alla formula!!!
    Dim valore2 As Single
    valore2 = (Valore - Canale(i).Bitmin) / _
    (Canale(i).Bitmax - Canale(i).Bitmin) * _
    (Canale(i).valMax - Canale(i).valMin) + Canale(i).valMin + Canale(i).valOff

    'ValoreADC = (Valore - Valoff) * (Bitmax - Bitmin) / _
    '(Valmax - Valmin) + Bitmin
    
    Count2value = valore2
End Function

Public Function v2mUni(Valore_ADC As Integer, Vref As Single, Gain As Integer) As Double
'Volts to measure, Unipolar
    Dim v As Double
    v = (Valore_ADC * Vref / 65535) / Gain
    v2mUni = v
End Function
Public Function v2mBip(Valore_ADC As Integer, Vref As Single, Gain As Integer) As Double
'Volts to measure, Bipolar
    Dim v As Double
    v = ((Valore_ADC - 32768) * Vref / 32768) / Gain
    v2mBip = v
End Function



Public Function adc2value(Valore_ADC As Long, Bitmin As Long, _
Bitmax As Long, valMax As Double, valMin As Double, valOff _
As Double) As Double
'From ADCount to Value

    Dim Valore As Double
    Valore = (Valore_ADC - Bitmin) / (Bitmax - Bitmin) * _
    (valMax - valMin) + valMin + valOff
    adc2value = Valore
    'Debug.Print "adc2value-->"; Valore
End Function

Public Function adc2value2(Valore_ADC As Double, Bitmin As Double, _
Bitmax As Double, valMax As Double, valMin As Double, valOff _
As Double) As Double
'From ADCount to Value

    Dim Valore As Double
    Valore = (Valore_ADC - Bitmin) / (Bitmax - Bitmin) * _
    (valMax - valMin) + valMin + valOff
    'Float = (Float - V2) * (T1 - T2) / (V1 - V2) + T2
    'T1 = 15.7  'valmax
    'T2 = 52.7  'valmin
    'V1 = 2.546 'bitmax
    'V2 = 0.99  'bitmin

    '811    bitmin
    '2086   bitmax
    '52.7   valmin
    '15.7   valmax
    '52.7   valoff
    
    adc2value2 = Valore
    Debug.Print "adc2value2-->"; Valore
End Function


Public Function value2ADC(value As Single, Bitmin As Long, _
Bitmax As Long, valMax As Double, valMin As Double, valOff _
As Double) As Long
    Dim ADC As Long
    ADC = ((value - valMin - valOff) / (valMax - valMin) * (Bitmax - Bitmin)) + Bitmin
    value2ADC = ADC
End Function

Public Function String2long(Stringa As String) As Long
    Dim lStringa As Integer
    Dim Lungo As Long
    
'    lStringa = Len(Stringa)
'    If lStringa <> 4 Then
'        Messaggio = "La lunghezza del numero è errata ->" + Str(lStringa) + " invece di 4"
'        MsgBox (Messaggio)
'    End If

    Stringa = SwapString(Stringa)
    Lungo = bytes2long(Stringa)
    String2long = Lungo
End Function

Public Sub TrovaRetta(x1 As Double, x2 As Double, _
                      y1 As Double, y2 As Double, _
                      m As Single, q As Single)
'Trova i parametri m e q dati 2 punti x1,y1 e x2,y2
'dove x e' il valore fornito dal convertitore
'e y il valore fisico corrispondente

'In ingresso vanno fornite le coordinate dei 2 punti
'In uscita si ottengono m e q
    m = (y2 - y1) / (x2 - x1)
    q = y1 - m * x1
End Sub

Public Sub TrovaBitVal(m As Single, q As Single, _
                       Bitmin As Integer, Bitmax As Integer, _
                       valMin As Single, valMax As Single, _
                       valOff As Single)
'Data l'equazione di una retta in m e q trova
'i parametri per la formula alternativa
'usata in simapro

'In ingresso vanno forniti m, q e opzionalmente bitMax
'se diverso da 4095 e bitmin se diverso da 0
'In uscita si ottengono i parametri in formula
    'bitMin = 0
    valMin = q
    If m < 0 Then
        'valOff = valMin
    End If
    If Bitmax = 0 Then Bitmax = 4095
    valMax = m * Bitmax + q
    'valOff = valMin
End Sub

Public Function c255toV(count As Long) As Single
'Converte un dato del convertitore in volt
'Il dato e' nell'intervallo 0-255 8bit
    c255toV = count * 5 / 255
End Function

Public Function c4095toV(count As Long) As Single
'Converte un dato del convertitore in volt
'Il dato e' nell'intervallo 0-4095 12bit
    c4095toV = count * 5 / 4095
End Function

Public Function c65535toV(count As Long) As Single
'Converte un dato del convertitore in volt
'Il dato e' nell'intervallo 0-65535 16bit
    c65535toV = count * 5 / 65535
End Function

