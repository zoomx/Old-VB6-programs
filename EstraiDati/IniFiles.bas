Attribute VB_Name = "IniFiles"
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

Public INIFile As String


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



