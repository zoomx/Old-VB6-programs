VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "fMain"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bTest 
      Caption         =   "Test"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton bChange 
      Caption         =   "Change"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bChange_Click()
    StripFile "E:\CD\_INGV\Progetti\Per Lorenzo\01nist6101.csv"
End Sub

Private Sub bTest_Click()
    Dim Filename As String
    Dim Path As String
    Dim ext As String
    GetFileElements "C:\cartella1\cartella2\file.txt", Path, Filename, ext
    Debug.Print Path
    Debug.Print Filename
    Debug.Print ext
End Sub

Private Sub Command1_Click()
    Dim FileNumbers As Integer
    Dim fs As New FileSearch
    Dim colOutput As Collection
    Set colOutput = fs.SearchFolders("D:\downloads", "txt|zip|exe", "*", vbTextCompare)
    FileNumbers = colOutput.Count
    Debug.Print "filenumbers="; FileNumbers

End Sub
