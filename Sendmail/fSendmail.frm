VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form fSendmail 
   Caption         =   "Send Mail"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bInvia 
      Caption         =   "&Invia"
      Height          =   612
      Left            =   960
      TabIndex        =   0
      Top             =   1560
      Width           =   1692
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label StatusTxt 
      Alignment       =   2  'Center
      Caption         =   "Mail"
      Height          =   372
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   2052
   End
End
Attribute VB_Name = "fSendmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Routine di ritardo in millisecondi
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim Response As String
Dim Start As Single
Dim Tmr As Single

Private Sub Form_Load()
    Dim Comandi() As String
    Dim nComandi As Integer
    Dim i As Integer
'inizializzazioni
    If App.PrevInstance Then
        Stringa = App.Title
        App.Title = "... duplicate instance."
        fSendmail.Caption = "... duplicate instance."
        AppActivate Stringa
        SendKeys "% ~", True
        End
    End If

    FileIni = App.Path + "\" + App.EXEName + ".ini"
    FileLog = App.Path + "\" + App.EXEName + ".log"
    ErrorFile = App.Path + "\" + App.EXEName + "_errori.log"

    Comandi() = Split(Command$, " ")
    nComandi = UBound(Comandi())
    HostRemoto = sReadINI(App.EXEName, "RemoteHost", FileIni)
    If HostRemoto = "" Then
        Errore = "Error file " + Chr$(34) + FileIni + Chr$(34) + " not found"
        WriteLog Errore
    End If
    Debug.Print
End Sub

Private Sub bInvia_Click()
    bInvia.Enabled = False
If Winsock1.LocalIP = "127.0.0.1" Then
    MsgBox "Non sei connesso", vbCritical, "Linea non disponibile"
    bInvia.Enabled = True
    Exit Sub
Else
    Winsock1.RemoteHost = HostRemoto
    Winsock1.RemotePort = 25
    'Winsock1.Close
    Winsock1.Connect
    'Sleep (300)
    WaitFor ("220")

    StatusTxt.Caption = "Connecting...."
    StatusTxt.Refresh
    
    'On Error Resume Next
    Winsock1.SendData "helo" & " Roberto" & vbCrLf 'normalmente si mette l'IP, comunque va bene qualunque cosa....ad ogni modo, il server il vostro IP lo sa già, quindi in HELO si può mettere quel che si vuole
    WaitFor ("250")
    
    StatusTxt.Caption = "HELO"
    StatusTxt.Refresh

    Winsock1.SendData "mail from: <balubalu@none.it>" & vbCrLf
    WaitFor ("250")
    StatusTxt.Caption = "mail from:"
    StatusTxt.Refresh

    Winsock1.SendData "RCPT TO: <zoomx@none.it>" & vbCrLf
    WaitFor ("250")
    StatusTxt.Caption = "RCPT TO:"
    StatusTxt.Refresh

    Winsock1.SendData "data" & vbCrLf

    WaitFor ("354")
    StatusTxt.Caption = "data"
    StatusTxt.Refresh

    Winsock1.SendData "Subject: Allarme programma non funzionante" & vbCrLf
    'WaitFor ("250")
    Winsock1.SendData "From: CTACQH" & vbCrLf
    'WaitFor ("250")
    Winsock1.SendData "To: Zoomx" & vbCrLf
    'WaitFor ("250")
    Winsock1.SendData vbCrLf & "il programma si e' fermato!!!" & vbCrLf & "." & vbCrLf
    WaitFor ("250")
    StatusTxt.Caption = "Fine data"
    StatusTxt.Refresh

    Winsock1.SendData "quit" & vbCrLf

    WaitFor ("221")
    Winsock1.Close
    StatusTxt.Caption = "MAIL"
    StatusTxt.Refresh

    MsgBox "Fatto!"
    bInvia.Enabled = True
End If
End Sub
Sub WaitFor(ResponseCode As String)
'    Dim tmr As Long
'    While Len(Response) = 0
'        DoEvents
'        Debug.Print "ciclo1->"; Response
'        If tmr > 50 Then
'            MsgBox "SMTP service error, unable to get a response1", 64 ', MsgTitle
'            Exit Sub
'        End If
'    Wend
'    While Left(Response, 3) <> ResponseCode
'        DoEvents
'        Debug.Print "ciclo2->"; Response
'        If tmr > 50 Then
'            MsgBox "SMTP service error, unable to get a response2", 64 ', MsgTitle
'            Exit Sub
'        End If
'    Wend

    Start = Timer ' Time event so won't get stuck in loop
    While Len(Response) = 0
        Tmr = Start - Timer
        DoEvents ' Let System keep checking for incoming response **IMPORTANT**
        'Debug.Print "ciclo1->"; Response
        If Tmr > 50 Then ' Time in seconds to wait
            MsgBox "SMTP service error, timed out while waiting for" & Response, 64 ', MsgTitle
            Exit Sub
        End If
    Wend
    While Left(Response, 3) <> ResponseCode
        DoEvents
        'Debug.Print "ciclo1->"; Response
        If Tmr > 50 Then
            MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, 64 ',MsgTitle
            Exit Sub
        End If
    Wend
Response = "" ' Sent response code to blank **IMPORTANT**

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Winsock1.GetData Response
End Sub

