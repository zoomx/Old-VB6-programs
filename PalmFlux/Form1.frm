VERSION 5.00
Object = "{9F4EED48-8EC7-4316-A47D-F6161874E478}#3.0#0"; "INGOTBUTTONCTL.DLL"
Object = "{84BE8A4A-3F9A-44E9-9B5E-E76D4888BA67}#3.0#0"; "INGOTTONECTL.DLL"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WEST Systems 2/00"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2400
   BeginProperty Font 
      Name            =   "AFPalm"
      Size            =   8.25
      Charset         =   2
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   StartUpPosition =   2  'CenterScreen
   Begin IngotToneCtl.AFTone AFTone1 
      Height          =   480
      Left            =   1680
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin IngotButtonCtl.AFButton InitBtn 
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "Form1.frx":0025
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

