Attribute VB_Name = "Module1"
Option Explicit
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public ComOk As Boolean
Public ComPort As Integer
Public Const GetDataFromProbe As String = &H81
Public GetReplyb() As Byte
Public AcquamasterOk As Boolean

Public Function SelectProbe(probe As Integer) As Boolean
    Dim Msg As Integer
    
    If probe < 1 Or probe > 6 Then
        SelectProbe = False
        Exit Function
    End If
    fMain.Text1.Text = fMain.Text1.Text + vbCrLf + "Sending command"
    Msg = &HF0 + probe
    'CloseCom
'    fMain.MSComm1.DTREnable = False
'    fMain.bDTRstatus.BackColor = vbRed
'    Sleep 100
'    fMain.MSComm1.DTREnable = True
'    fMain.bDTRstatus.BackColor = vbGreen
    fMain.MSComm1.DTREnable = True
    fMain.bDTRstatus.BackColor = vbGreen
    Sleep 100
    fMain.MSComm1.DTREnable = False
    fMain.bDTRstatus.BackColor = vbRed
    fMain.MSComm1.Output = Chr$(Msg)  'Trim(Str(Msg)) '+ vbCrLf
    fMain.Text1.Text = fMain.Text1.Text + vbCrLf + "Command sent" + Str(Msg)
    'fMain.MSComm1.DTREnable = False
End Function

Public Function GetReply(NumChar) As String
    Dim TimeStop As Long
    Dim Text As String
    
    Dim i As Integer
    TimeStop = Timer + 2
    Do
        DoEvents
    Loop Until (fMain.MSComm1.InBufferCount >= NumChar) Or (Timer > TimeStop)
    GetReplyb = fMain.MSComm1.Input
    If UBound(GetReplyb) = -1 Then
        GetReply = "No Reply"
        Text = "No Reply"
        fMain.Text1.Text = fMain.Text1.Text + Text
        Exit Function
    Else
        Text = Char2ascii(GetReplyb)
        AcquamasterOk = True
    End If
    
    fMain.Text1.Text = fMain.Text1.Text + vbCrLf
'    For i = 0 To UBound(GetReplyb)
'        fMain.Text1.Text = fMain.Text1.Text + Chr(GetReply(i))
'    Next
    fMain.Text1.Text = fMain.Text1.Text + Text + vbCrLf
    'fMain.MSComm1.DTREnable = False
End Function

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
            MsgBox Msg, vbOKOnly, "Errore"

            
            Err.Clear   ' Cancella i campi dell'oggetto
            ComOk = False
            Exit Sub
        Case 8002
            Msg = "Errore la porta Com" + Str$(ComPort) + " non esiste!"
            MsgBox Msg, vbOKOnly, "Errore"

            Err.Clear   ' Cancella i campi dell'oggetto
            ComOk = False
            Exit Sub
        Case Else
            Msg = Err.Description
            MsgBox Msg, vbOKOnly, "Errore"


            Exit Sub
    End Select

End Sub

Public Sub CloseCom()
    'Chiude la porta seriale se non è già chiusa
    fMain.MSComm1.InBufferCount = 0
    If fMain.MSComm1.PortOpen = True Then fMain.MSComm1.PortOpen = False
End Sub
Public Function Val2(Valore As String) As Single
'Simile alla val ma per separatore decimale usa sia il
'punto che la virgola
    Dim ip As Integer
    Dim iv As Integer
    Dim lStringa As Integer
    Dim Temp As Single
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
    Temp = CSng(Stringa)
    'If ip <> 0 And iv <> 0 Then
    'Se ci sono tutte e due?
    Val2 = Temp
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
Public Function Char2ascii(Stringa() As Byte) As String
'Trasforma una stringa contenente caratteri ASCII e non
'ASCII in stringa di codici di caratteri ASCII
'Viene gestito anche il chr$(0)
    Dim lStringa As Integer
    Dim tStringa As String
    Dim i As Integer
    
    lStringa = UBound(Stringa())
    For i = 0 To lStringa
        If Stringa(i) = 0 Then 'Mid(Stringa, i, 1) = Chr$(0) Then
            tStringa = tStringa + " " + "00"
        Else
            tStringa = tStringa + Str((Stringa(i)))
        End If
    Next
    Char2ascii = tStringa
End Function

Public Sub ParseReply(Message As String)
    Dim Field(7) As Byte
    Field(0) = Left(Message, 1)
    Field(1) = Mid(Message, 2, 1)
    Field(2) = Mid(Message, 3, 1)
    Field(1) = Mid(Message, 4, 1)
    Field(1) = Mid(Message, 5, 1)
    Field(1) = Mid(Message, 6, 1)
    Field(1) = Right(Message, 1)
    fMain.Text1.Text = fMain.Text1.Text + vbCrLf + "165->" + Str(Asc(Field(0)))
    fMain.Text1.Text = fMain.Text1.Text + vbCrLf + "170->" + Str(Asc(Field(7)))
End Sub

Public Sub ParseAcqReply()
    Dim Temperb As Long
    Dim Temp As Single
    Dim Press As Long
    fMain.Text1.Text = fMain.Text1.Text + vbCrLf
    If GetReplyb(0) = 165 Then fMain.Text1.Text = fMain.Text1.Text + "First Byte OK" + vbCrLf
    If GetReplyb(7) = 170 Then fMain.Text1.Text = fMain.Text1.Text + "Last Byte OK" + vbCrLf
    If GetReplyb(1) = 64 Then fMain.Text1.Text = fMain.Text1.Text + "Command unknown" + vbCrLf
    Temperb = GetReplyb(2) + 256 * CLng(GetReplyb(3))
    Temp = Temperb / 10
    Press = GetReplyb(4) + 256 * CLng(GetReplyb(5))
    fMain.Text1.Text = fMain.Text1.Text + Str(Temp) + " " + Str(Press)
End Sub

Public Sub ParseData()
    Dim ipH As Integer
    Dim pH As Single
    Dim iRedox As Integer
    Dim Redox As Single
    Dim iPP As Integer
    Dim PP As Single
    Dim iTP As Integer
    Dim TP As Single
    Dim iCond As Integer
    Dim CondScale As Integer
    Dim Cond As Single
    
    ipH = GetReplyb(0) + 256 * CLng(GetReplyb(1))
    pH = ipH / 100
    iRedox = GetReplyb(2) + 256 * CLng(GetReplyb(3))
    Redox = iRedox
    iPP = GetReplyb(4) + 256 * CLng(GetReplyb(5))
    PP = iPP
    iTP = GetReplyb(6) + 256 * CLng(GetReplyb(7))
    TP = iTP / 10
    fMain.Text1.Text = fMain.Text1.Text + vbCrLf + "pH=" + Str(pH) + vbCrLf
    fMain.Text1.Text = fMain.Text1.Text + "Redox=" + Str(Redox) + vbCrLf
    fMain.Text1.Text = fMain.Text1.Text + "Pressione=" + Str(PP) + vbCrLf
    fMain.Text1.Text = fMain.Text1.Text + "Temperatura=" + Str(TP) + vbCrLf
End Sub

Public Sub AskConducibility()
    Dim Comando As String
    Dim iCond As Long
    OpenCom
    Comando = Chr$(133) & Chr$(0) & Chr$(0) 'Scala conducibilità=0
    fMain.MSComm1.Output = Comando
    Sleep 200
    Comando = Chr$(130) & Chr$(4) & Chr$(0) 'Dammi la conducibilità
    fMain.MSComm1.Output = Comando
    GetReply 2
    Sleep 200
    Comando = Chr$(130) & Chr$(4) & Chr$(0) 'Dammi la conducibilità
    fMain.MSComm1.Output = Comando
    GetReply 2

    iCond = GetReplyb(0) + 256 * CLng(GetReplyb(1))
    
    If IsCondOk(iCond) = True Then
        PrintCond
        Exit Sub
   End If
    
    Comando = Chr$(133) & Chr$(0) & Chr$(1) 'Scala conducibilità=1
    fMain.MSComm1.Output = Comando
    Sleep 200
    Comando = Chr$(130) & Chr$(4) & Chr$(0) 'Dammi la conducibilità
    fMain.MSComm1.Output = Comando
    GetReply 2
    Sleep 200
    Comando = Chr$(130) & Chr$(4) & Chr$(0) 'Dammi la conducibilità
    fMain.MSComm1.Output = Comando
    GetReply 2

    iCond = GetReplyb(0) + 256 * CLng(GetReplyb(1))
    If IsCondOk(iCond) = True Then
        PrintCond
        Exit Sub
    End If
    
    Comando = Chr$(133) & Chr$(0) & Chr$(2) 'Scala conducibilità=2
    fMain.MSComm1.Output = Comando
    Sleep 200
    Comando = Chr$(130) & Chr$(4) & Chr$(0) 'Dammi la conducibilità
    fMain.MSComm1.Output = Comando
    GetReply 2
    Comando = Chr$(130) & Chr$(4) & Chr$(0) 'Dammi la conducibilità
    fMain.MSComm1.Output = Comando
    GetReply 2

    iCond = GetReplyb(0) + 256 * CLng(GetReplyb(1))
    If IsCondOk(iCond) = True Then
        PrintCond
        Exit Sub
    End If

    Comando = Chr$(133) & Chr$(0) & Chr$(3) 'Scala conducibilità=3
    fMain.MSComm1.Output = Comando
    Sleep 200
    Comando = Chr$(130) & Chr$(4) & Chr$(0) 'Dammi la conducibilità
    fMain.MSComm1.Output = Comando
    GetReply 2
    Comando = Chr$(130) & Chr$(4) & Chr$(0) 'Dammi la conducibilità
    fMain.MSComm1.Output = Comando
    GetReply 2

    PrintCond

'    GetReplyb(0) = CByte(164)
'    GetReplyb(1) = CByte(0)
'    fMain.Text1.Text = fMain.Text1.Text + vbCrLf + Char2ascii(GetReplyb)
'    iCond = GetReplyb(0) + 256 * CLng(GetReplyb(1))
'    PrintCond
'
'    GetReplyb(0) = CByte(183)
'    GetReplyb(1) = CByte(65)
'    fMain.Text1.Text = fMain.Text1.Text + vbCrLf + Char2ascii(GetReplyb)
'    iCond = GetReplyb(0) + 256 * CLng(GetReplyb(1))
'    PrintCond
'
'    GetReplyb(0) = CByte(164)
'    GetReplyb(1) = CByte(134)
'    fMain.Text1.Text = fMain.Text1.Text + vbCrLf + Char2ascii(GetReplyb)
'    iCond = GetReplyb(0) + 256 * CLng(GetReplyb(1))
'    PrintCond
'
'    GetReplyb(0) = CByte(183)
'    GetReplyb(1) = CByte(195)
'    fMain.Text1.Text = fMain.Text1.Text + vbCrLf + Char2ascii(GetReplyb)
'    iCond = GetReplyb(0) + 256 * CLng(GetReplyb(1))
'    PrintCond


End Sub

Public Function ParseConducibility(iValue As Long) As Single
    Dim iCond As Long
    Dim iScale As Long
    Dim scala As Single
    iCond = iValue And 2047
    iScale = (iValue And 49152) / 16384
    fMain.Text1.Text = fMain.Text1.Text + vbCrLf + Str(iCond) + " " + Str(iScale) + " "
    Select Case iScale
        Case 3
            scala = 10
        Case 2
            scala = 100
        Case 1
            scala = 1000
        Case 0
            scala = 10000
    End Select
    fMain.Text1.Text = fMain.Text1.Text + "Scala=" + Str(scala) + " "
    ParseConducibility = iCond / scala
End Function

Public Function IsCondOk(value As Long) As Boolean
    Dim scala As Long
    scala = (value And 49152) / 16384
    value = value And 2047
    IsCondOk = True
    If value < 150 Or value > 1900 Then
        IsCondOk = False
    End If
    fMain.Text1.Text = fMain.Text1.Text + "Scala=" + Str(scala) + " " + "valore=" + Str(value) + " " + Str(IsCondOk) + " "

End Function

Public Sub PrintCond()
    Dim Conducibilita As Single
    Dim iCond As Long
    iCond = GetReplyb(0) + 256 * CLng(GetReplyb(1))
    Conducibilita = ParseConducibility(iCond)
    fMain.Text1.Text = fMain.Text1.Text + "Conducibilità=" + Str(Conducibilita) + vbCrLf

End Sub
