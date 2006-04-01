Attribute VB_Name = "FunzioniComuni"
Option Explicit

Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
As String, ByVal lpKeyName As String, ByVal lpDefault As _
String, ByVal lpReturnedString As String, ByVal nSize As _
Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" _
Alias "WritePrivateProfileStringA" (ByVal _
lpApplicationName As String, ByVal lpKeyName As String, _
ByVal lpString As String, ByVal lpFileName As String) As Long

Public Stringa As String
Public FileIni As String
Public FileLog As String
Public ErrorFile As String
Public HostRemoto As String
Public Errore As String

Function sReadINI(AppName, KeyName, filename As String) As String
'*Returns a string from an INI file. To use, call the  *
'*functions and pass it the AppName, KeyName and INI   *
'*File Name, [sReg=sReadINI(App1,Key1,INIFile)]. If you *
'*need the returned value to be a integer then use the *
'*val command.                                         *
'*******************************************************

Dim sRet As String
    sRet = String(255, Chr(0))
    sReadINI = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), filename))
End Function

Public Function WriteINI(sAppname, sKeyName, sNewString, sFileName As String) As Long
'*Writes a string to an INI file. To use, call the     *
'*function and pass it the sAppname, sKeyName, the New *
'*String and the INI File Name,                        *
'*[R=WriteINI(App1,Key1,sReg,INIFile)]. Returns a 1 if *
'*there were no errors and a 0 if there were errors.   *
'*******************************************************


    WriteINI = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)
End Function

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

Public Sub FinePerErrore()
    Dim Mes As String
    'CloseCom
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
    Print #nfile, "Errore in " + App.Title + App.Major + App.Minor + " del "; Date$, " alle "; Time$
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

Public Function WriteLog(Stringa As String) As Boolean
    Dim fn As Long
    Dim SubRoutine As String
    WriteLog = False
    SubRoutine = "WriteLog"
    On Error GoTo ErrHandler
    fn = FreeFile
    Open FileLog For Append As #fn
    Print #fn, "-----------------------------------------"
    Print #fn, Format(Now, "dd/mm/yyyy hh:mm:ss")
    Print #fn, Stringa
    DoEvents
    Close fn
    WriteLog = True
Exit Function
ErrHandler:
    ErrHandler
End Function



