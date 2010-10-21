Attribute VB_Name = "Module1"
Option Explicit
Public BasePath As String

Public Function Text2String(ByVal strPath As String, _
                            Optional ByVal strSearch As String, _
                            Optional ByVal CompareMode As VbCompareMethod _
                            ) As String
    Text2String = "ok"
    Exit Function
    Dim strBuff As String
    Dim strContent As String
    Dim rxSearch As New RegExp
    Dim rxMatches As MatchCollection
    Dim rxMatch As Match
    Dim strOutput As String
    'On Error GoTo ErrTrap
  
    rxSearch.IgnoreCase = (CompareMode = vbTextCompare)
    rxSearch.Global = True
    rxSearch.MultiLine = True
    rxSearch.Pattern = strSearch
   
    Open strPath For Binary As #1
    strBuff = Space(LOF(1))
    Get #1, , strBuff
    If strSearch = vbNullString Then
        strOutput = strBuff
    Else
        Set rxMatches = rxSearch.Execute(strBuff)
        For Each rxMatch In rxMatches
            strOutput = strOutput & rxMatch.Value
        Next
    End If
   
    Text2String = strOutput
   
    Close #1 ' close the connection before proceeding to the next step
ErrTrap:
   If Err Then Err.Raise Err.Number, , "Error from Functions.Text2String " & Err.Description
End Function


Public Function StripFile(Filename As String) As Boolean
    Dim FileNameOut As String
    Dim Path As String
    Dim Buffer As String
    Dim buffer2 As Variant
    Dim FileNumber As Integer
    Dim i As Long
    Dim stringa As String
    StripFile = False
    
    'prendi il file name dal path
    'prendi il path
    'lo trasformi
    'aggiungi al path il nuovo filename
    stringa = vbCrLf
    
    FileNumber = FreeFile
    Open Filename For Input As #FileNumber
    'Get #1, , buffer2
    Buffer = Input(LOF(FileNumber), #FileNumber)
    Close #FileNumber
    'i = InStr(1, stringa, Buffer, vbTextCompare)
    i = InStr(Buffer, stringa)
    'i = InStrB(stringa, Buffer)
    Debug.Print i; " OK!"
    Buffer = Mid(Buffer, i + 2, Len(Buffer) - i - 2)
    FileNumber = FreeFile
    FileNameOut = Mid(Filename, 1, Len(Filename) - 3) + "xl"
    Open FileNameOut For Output As #FileNumber
    Print #FileNumber, Buffer
    Close FileNumber
    
End Function

Public Function GetNameFromDir(Dir As String) As String
    Dim i As Long
    Dim lasti As Long
    Dim Dirr As String
    Dirr = Dir
    Do
        lasti = i
        i = InStr(Dir, "\")
        Dir = Right(Dir, Len(Dir) - i)
    Loop Until i = 0
    GetNameFromDir = Dir
End Function


Public Sub GetFileElements(PathFilename As String, Path As String, Filename As String, Extension As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Path = FSO.GetParentFolderName(PathFilename)
    Filename = FSO.GetFileName(PathFilename)
    Extension = FSO.GetExtensionName(PathFilename)
    
End Sub
