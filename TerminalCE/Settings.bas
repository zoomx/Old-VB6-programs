Attribute VB_Name = "modSettings"
'********************************************************************************
'*
'*  VBCEComm - settings code
'*
'*  Copyright (c) 1998 Microsoft Corporation

Public Sub SetSettings()
    'Returns settings back to the Comm control
    If Len(frmConnect.txtRT.Text) Then
        frmMain.Comm1.RThreshold = frmConnect.txtRT.Text
    End If
    If Len(frmConnect.txtST.Text) Then
        frmMain.Comm1.SThreshold = frmConnect.txtST.Text
    End If
    If Len(frmConnect.cboPort.Text) Then
        frmMain.Comm1.CommPort = frmConnect.cboPort.Text
    End If
    If Len(frmConnect.cboSettings.Text) Then
        frmMain.Comm1.Settings = frmConnect.cboSettings.Text
    End If
    If Len(frmConnect.txtIL.Text) Then
        frmMain.Comm1.InputLen = frmConnect.txtIL.Text
    End If
End Sub
Public Sub ContinueStart()
    'Handles opening the com port, enabling/disabling
    'the proper controls and updating frmMain.
    If frmMain.Comm1.PortOpen = False Then
        frmMain.Comm1.PortOpen = True
    End If
    frmMain.cmdStart.Enabled = False
    frmMain.cmdEnd.Enabled = True
    ShowComm
    ShowErr
End Sub

Sub ShowComm()
    'This procedure updates the labels on frmMain to reflect
    'the current values of specified Comm control properties.
    frmMain.lblSThreshold.Caption = frmMain.Comm1.SThreshold
    frmMain.lblSThreshold.Refresh
    frmMain.lblRThreshold.Caption = frmMain.Comm1.RThreshold
    frmMain.lblRThreshold.Refresh
    frmMain.lblSettings.Caption = frmMain.Comm1.Settings
    frmMain.lblSettings.Refresh
    frmMain.lblInBuffCount.Caption = frmMain.Comm1.InBufferCount
    frmMain.lblInBuffCount.Refresh
    frmMain.lblOutBuffCount.Caption = frmMain.Comm1.OutBufferCount
    frmMain.lblOutBuffCount.Refresh
    frmMain.lblEvent.Caption = frmMain.Comm1.CommEvent
    frmMain.lblEvent.Refresh
    frmMain.lblInputLen.Caption = frmMain.Comm1.InputLen
    frmMain.lblInputLen.Refresh
    frmMain.lblComPort.Caption = frmMain.Comm1.CommPort
    frmMain.lblComPort.Refresh
End Sub
Sub ShowErr()
    'This procedure reports any errors that occur
    If Err.Number <> 0 Then
        frmMain.txtError.Text = Err.Number & " - " & Err.Description
        frmMain.txtError.Refresh
    End If
End Sub
