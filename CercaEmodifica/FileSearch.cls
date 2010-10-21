VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dim fs As New FileSearch
'Dim colOutput As Collection
'Set colOutput = fs.SearchFolders("C:\Documents and Settings", "txt|ini|log", "myname|myfriendsname", vbTextCompare)

'http://www.vbknowledgebase.com/?Id=45&Desc=VB6-Recursive-File-Content-Search

Private m_colResults As Collection
' ************************************************************
'*  Purpose :   Search the list of items from the folder
'*
'*  Inputs  :   strPath(String)                 Path to folder
'*              strSearchFiles(String)          Search file Name
'*              strSearchContent(String)        Search File Content
'*              CompareMode(VbCompareMethod)    Compare mode
'*  Returns :   list of items in the folder
'*
' ************************************************************
Public Function SearchFolders(strFolder As String, _
                        Optional ByVal strSearchFiles As String, _
                        Optional ByVal strSearchContent As String, _
                        Optional ByVal CompareMode As VbCompareMethod) As Collection
Set m_colResults = New Collection
ListFolders strFolder, strSearchFiles, strSearchContent, CompareMode
Set SearchFolders = m_colResults
End Function
 
' ************************************************************
'*  Purpose :   Get the list of items from the folder
'*
'*  Inputs  :   strPath(String)                 Path to folder
'*              strSearchFiles(String)          Search file Name
'*              strSearchContent(String)        Search File Content
'*              CompareMode(VbCompareMethod)    Compare mode
'*  Returns :   list of items in the folder
'*
' ************************************************************
Private Sub ListFolders(strFolder As String, _
                        Optional ByVal strSearchFiles As String, _
                        Optional ByVal strSearchContent As String, _
                        Optional ByVal CompareMode As VbCompareMethod)
 
Dim fold As Scripting.Folder
Dim foldSub As Scripting.Folder
Dim fil As File
Dim FSO As New FileSystemObject
Dim strOutput As String
 
'On Error GoTo ErrTrap
Set fold = FSO.GetFolder(strFolder)
For Each fil In fold.Files
     
    If IsMatchedRx(fil.Name, strSearchFiles, CompareMode) Or _
        Trim(strSearchFiles) = vbNullString Then
        If strSearchContent <> vbNullString Then
'            strOutput = Text2String(fil.Path, strSearchContent, vbTextCompare)
'            DoEvents
'             'Debug.Print "Searching the file " & fil.Name
'            If strOutput <> "" Then
            m_colResults.Add fil.Path & " " & strOutput
'            Else
'             'Debug.Print fil.Name & "," & fil.Path
'            End If
        End If
    End If
    DoEvents
Next
For Each foldSub In fold.SubFolders
    'Debug.Print foldSub.Path
    DoEvents
    ListFolders foldSub.Path, strSearchFiles, strSearchContent, CompareMode
Next
ErrTrap:
If Err Then Err.Raise Err.Number, , "Error form Functions.ListFolders " & Err.Description
End Sub
 
 
Public Property Get SearchResults() As Collection
    Set SearchResults = m_colResults
End Property

Public Function IsMatchedRx(strText As String, _
                            strSearch As String, _
                            Optional ByVal CompareMode As VbCompareMethod) As Boolean
                           
    Dim rxSearch As New RegExp
    On Error GoTo ErrTrap
    rxSearch.IgnoreCase = (CompareMode = vbTextCompare)
    rxSearch.Global = True
    rxSearch.MultiLine = True
    rxSearch.Pattern = strSearch
   
    IsMatchedRx = rxSearch.Execute(strText).Count > 0
ErrTrap:
    If Err Then Err.Raise Err.Number, , "Error form Functions.IsMatchedRx " & Err.Description
End Function

