Attribute VB_Name = "MDL"
Option Explicit

Public Old_Info_Check As String
Public CRLF As String                  'VbCR+VbLF
Public RxMsg As String                 'Stringa buffer di ricezione
Public OldRx As String
Public TimeOut                         'Timeout variable
'Public debugTime
Public Info_SampInterval  As Double    'Intervallo tra un punto e la'ltro in secndi
Public Info_NumPoints As Integer       'Numero dei punti della curva
Public Info_AC_K As Double             'Costante della camera di accumulo
Public Info_unit As String             'Unita di misura che viene da ppm/sek*ac_k
Public info_FluxLineppmSec As Double
Public info_FluxLine As Double
Public Info_Ch_attivo(3) As Boolean    'dice se il canale è attivo
Public Info_Ch_FS(3) As Double          'Valore massimo del sensoref(x)=a*x+b
Public Info_Ch_LS(3) As Double         'Valore minimo del sensore
Public info_Ch_Type(3) As String       ' "4-20" "0-20" "0-5" "DIG" che tipo di output
Public Info_Ch_descr(3) As String      'Descrizione del canale
Public Info_Ch_FStr(3) As Integer      'Numero di cifre dopo il punto
Public Info_Ch_v(3) As Double          'valore dell'ADC dopo la correzione
Public Info_Check As String            'Variabile di controllo e indica che detector si usa
Public Info_RS232_Port As Integer      'Indica a palmflux che porta seriale usare (CASSIOPEIA=1 HP420=2)

Public Info_Pressure_A As Double       'Only for NI6b11_N
Public Info_Pressure_B As Double       'Only for NI6b11_N
Public Info_Pressure_A_Dev As Double   'Only for NI6b11_N
Public Info_Pressure_B_Dev As Double   'Only for NI6b11_N
Public Info_Pressure_Switch As Integer 'Only for NI6b11_N
Public Info_Decimazione As Boolean     ' se true il grafico viene decimato, altrimenti NO

Public Info_Area As String
Public Info_punto As String
Public Info_Lat As String
Public Info_Long As String
Public Info_Date As Date   'String
Public Info_Flux As Double
Public Info_Fluxppmsec As Double
Public Info_ErreQ As Double
Public Info_Fname As String
Public Info_RecordNumber As Long
Public Info_Native_Unit As String

Public Grd_MaxX As Single
Public Grd_MaxY As Single
Public Grd_MinX As Single
Public Grd_MinY As Single

Public Stop_Acquisition As Integer


Public DataIsNotSaved As Boolean       'Flag di stato: dati salvati o no
Public rec_rl As Double                'Limite sinistro
Public rec_ll As Double                'Limite destro

Public S_X As Double                   'Sommatorie per regressione
Public S_Y As Double
Public S_XQ As Double
Public S_YQ As Double
Public S_XY As Double
Public S_RQ As Double
Public S_v As Double



Public x1 As Single                    'Variabili di appoggio al grafico su GR
Public y1 As Single                    '
Public x2 As Single
Public y2 As Single

Public Counter As Integer              'Contatore dei punti della curva
Public Dati() As Double                'Dati(Info_NumPoints) della curva

Public Const Drager = "DRAGER IR CO2"
Public Const Licor = "LICOR LI-800"
Public Const Politron = "Politron II"
Public Const Riken = "Riken 100mV"
Public Const NI6B11 = "6B11 mV"
Public Const NI6B13 = "6B13 Pt100"
Public Const NI6B11_N = "6B11 mV native"


Public Const West = "WEST Systems: "
Public Status As Integer               'Variabile di gestione del bottone CmdBtn
Public Const wait = 1
Public Const flux = 2
Public Const settingL = 3
Public Const settingR = 4
Public Const Calculate = 5
Public Const Storing = 6

Public Ip       'VARIABILI di PARSE e PARSE6B
Public Ep
Public p As String
Public V As Double
Public a As Double
Public b As Double

Public LeggiSeriale_Ora As Double


Public ImBusy As Integer

Public ParseCode As Boolean

'Public Regressione_A As Double 'Variabili della funzione regressione
'Public Regressione_B As Double 'Variabili della funzione regressione

Public NI6BN_Tz As Long
Public NI6BN_p As String
Public NI6BN_ec As Integer
Public NI6BN_OldValue As Double
Public NI6BN_CheckError As Boolean

Public Stringa As String
Dim NomeFile As String

Sub SaveConfiguration(name As String)

    Dim Filetext As CFileTextWritable
    Dim FileMgr As New CFileManager

'StartUpFrm.File1.Open name, fsModeBinary, fsAccessWrite
    #If APPFORGE Then
        NomeFile = "Card:\" + name
    #Else
        NomeFile = name
    #End If

    Set Filetext = FileMgr.OpenAsText(NomeFile, afFileModeCreate)

    Filetext.Write (Str(Info_SampInterval) + vbCrLf) ' As Double    'Intervallo tra un punto e la'ltro in secndi
    Filetext.Write (Str(Info_NumPoints) + vbCrLf) 'As Integer       'Numero dei punti della curva
    Filetext.Write (Str(Info_AC_K) + vbCrLf) 'As Double             'Costante della camera di accumulo
    Filetext.Write (Info_unit + vbCrLf) 'As String             'Unita di misura che viene da ppm/sek*ac_k

    Filetext.Write (Str(Info_Ch_attivo(0)) + vbCrLf) ' As Boolean    'dice se il canale è attivo
    Filetext.Write (Str(Info_Ch_FS(0)) + vbCrLf) ' As Double          'f(x)=a*x+b
    Filetext.Write (Str(Info_Ch_LS(0)) + vbCrLf) ' As Double
    Filetext.Write (Info_Ch_descr(0) + vbCrLf) 'As String      'Descrizione del canale
    Filetext.Write (Str(Info_Ch_FStr(0)) + vbCrLf)
    Filetext.Write (Str(Info_Ch_v(0)) + vbCrLf)
    Filetext.Write (info_Ch_Type(0) + vbCrLf)

    Filetext.Write (Str(Info_Ch_attivo(1)) + vbCrLf) ' As Boolean    'dice se il canale è attivo
    Filetext.Write (Str(Info_Ch_FS(1)) + vbCrLf) ' As Double          'f(x)=a*x+b
    Filetext.Write (Str(Info_Ch_LS(1)) + vbCrLf) ' As Double
    Filetext.Write (Info_Ch_descr(1) + vbCrLf) 'As String      'Descrizione del canale
    Filetext.Write (Str(Info_Ch_FStr(1)) + vbCrLf)
    Filetext.Write (Str(Info_Ch_v(1)) + vbCrLf)
    Filetext.Write (info_Ch_Type(1) + vbCrLf)

    Filetext.Write (Str(Info_Ch_attivo(2)) + vbCrLf) ' As Boolean    'dice se il canale è attivo
    Filetext.Write (Str(Info_Ch_FS(2)) + vbCrLf) ' As Double          'f(x)=a*x+b
    Filetext.Write (Str(Info_Ch_LS(2)) + vbCrLf) ' As Double
    Filetext.Write (Info_Ch_descr(2) + vbCrLf) 'As String      'Descrizione del canale
    Filetext.Write (Str(Info_Ch_FStr(2)) + vbCrLf)
    Filetext.Write (Str(Info_Ch_v(2)) + vbCrLf)
    Filetext.Write (info_Ch_Type(2) + vbCrLf)

    Filetext.Write (Str(Info_Ch_attivo(3)) + vbCrLf) ' As Boolean    'dice se il canale è attivo
    Filetext.Write (Str(Info_Ch_FS(3)) + vbCrLf) ' As Double          'f(x)=a*x+b
    Filetext.Write (Str(Info_Ch_LS(3)) + vbCrLf) ' As Double
    Filetext.Write (Info_Ch_descr(3) + vbCrLf) 'As String      'Descrizione del canale
    Filetext.Write (Str(Info_Ch_FStr(3)) + vbCrLf)
    Filetext.Write (Str(Info_Ch_v(3)) + vbCrLf)
    Filetext.Write (info_Ch_Type(3) + vbCrLf)

    Filetext.Write (Info_Check + vbCrLf)
    Filetext.Write (Str(Info_RS232_Port) + vbCrLf)
    Set Filetext = Nothing

End Sub


Sub ReadConfiguration(name As String, frm As Form)
    Dim Filetext As CFileTextReadable
    Dim FileMgr As New CFileManager
    #If APPFORGE Then
        NomeFile = "Card:\" + name
    #Else
        NomeFile = name
    #End If

    
    Set Filetext = FileMgr.OpenReadOnlyAsText(NomeFile)
    'frm.File1.Open name, fsModeBinary, fsAccessRead
    Stringa = Filetext.ReadLine
    Info_SampInterval = CDbl(Val(Stringa)) ' As Double    'Intervallo tra un punto e la'ltro in secondi
    Stringa = Filetext.ReadLine
    Info_NumPoints = Val(Stringa) 'As Integer Numero dei punti della curva
    Stringa = Filetext.ReadLine
    Info_AC_K = CDbl(Val(Stringa))  ' As Double Costante della camera di accumulo
    Info_unit = Filetext.ReadLine 'As String Unita di misura che viene da ppm/sek*ac_

    Stringa = Trim(Filetext.ReadLine)
    If Stringa = "True" Or Stringa = "Vero" Then
        Info_Ch_attivo(0) = True ' As Boolean    'dice se il canale è attivo
    Else
        Info_Ch_attivo(0) = False
    End If
    Stringa = Filetext.ReadLine
    Info_Ch_FS(0) = CDbl(Val(Stringa)) ' As Double          'f(x)=a*x+b
    Stringa = Filetext.ReadLine
    Info_Ch_LS(0) = CDbl(Val(Stringa)) ' As Double
    Info_Ch_descr(0) = Filetext.ReadLine 'As String      'Descrizione del canale
    Stringa = Filetext.ReadLine
    Info_Ch_FStr(0) = Val(Stringa)
    Stringa = Filetext.ReadLine
    Info_Ch_v(0) = CDbl(Val(Stringa))
    info_Ch_Type(0) = Filetext.ReadLine

    Stringa = Trim(Filetext.ReadLine)
    If Stringa = "True" Or Stringa = "Vero" Then
        Info_Ch_attivo(1) = True ' As Boolean    'dice se il canale è attivo
    Else
        Info_Ch_attivo(1) = False
    End If
    Stringa = Filetext.ReadLine
    Info_Ch_FS(1) = CDbl(Val(Stringa)) ' As Double          'f(x)=a*x+b
    Stringa = Filetext.ReadLine
    Info_Ch_LS(1) = CDbl(Val(Stringa)) ' As Double
    Info_Ch_descr(1) = Filetext.ReadLine 'As String      'Descrizione del canale
    Stringa = Filetext.ReadLine
    Info_Ch_FStr(1) = Val(Stringa)
    Stringa = Filetext.ReadLine
    Info_Ch_v(1) = CDbl(Val(Stringa))
    info_Ch_Type(1) = Filetext.ReadLine

    Stringa = Trim(Filetext.ReadLine)
    If Stringa = "True" Or Stringa = "Vero" Then
        Info_Ch_attivo(2) = True ' As Boolean    'dice se il canale è attivo
    Else
        Info_Ch_attivo(2) = False
    End If
    Stringa = Filetext.ReadLine
    Info_Ch_FS(2) = CDbl(Val(Stringa)) ' As Double          'f(x)=a*x+b
    Stringa = Filetext.ReadLine
    Info_Ch_LS(2) = CDbl(Val(Stringa)) ' As Double
    Info_Ch_descr(2) = Filetext.ReadLine 'As String      'Descrizione del canale
    Stringa = Filetext.ReadLine
    Info_Ch_FStr(2) = Val(Stringa)
    Stringa = Filetext.ReadLine
    Info_Ch_v(2) = CDbl(Val(Stringa))
    info_Ch_Type(2) = Filetext.ReadLine

    Stringa = Trim(Filetext.ReadLine)
    If Stringa = "True" Or Stringa = "Vero" Then
        Info_Ch_attivo(0) = True ' As Boolean    'dice se il canale è attivo
    Else
        Info_Ch_attivo(0) = False
    End If
    Stringa = Filetext.ReadLine
    Info_Ch_FS(0) = CDbl(Val(Stringa)) ' As Double          'f(x)=a*x+b
    Stringa = Filetext.ReadLine
    Info_Ch_LS(0) = CDbl(Val(Stringa)) ' As Double
    Info_Ch_descr(0) = Filetext.ReadLine 'As String      'Descrizione del canale
    Stringa = Filetext.ReadLine
    Info_Ch_FStr(0) = Val(Stringa)
    Stringa = Filetext.ReadLine
    Info_Ch_v(0) = CDbl(Val(Stringa))
    info_Ch_Type(0) = Filetext.ReadLine

    Info_Check = Filetext.ReadLine
    Stringa = Filetext.ReadLine
    Info_RS232_Port = Val(Stringa)
    Set Filetext = Nothing
End Sub

Sub InitInfo()
 
Info_SampInterval = 1  'Sec.
Info_NumPoints = 360
    

Info_AC_K = 14.35
Info_unit = "gr/(m^2 day)"

Info_Ch_FS(0) = 10000
Info_Ch_LS(0) = 0
Info_Ch_attivo(0) = True
Info_Ch_descr(0) = "Drager CO2"
Info_Ch_FStr(0) = 0
info_Ch_Type(0) = "4-20"

Info_Ch_FS(1) = 200
Info_Ch_LS(1) = 0
Info_Ch_attivo(1) = False
Info_Ch_descr(1) = "Air T. °C"
Info_Ch_FStr(1) = 1
info_Ch_Type(1) = "DIG"

Info_Ch_FS(2) = 200
Info_Ch_LS(2) = 0
Info_Ch_attivo(2) = False
Info_Ch_descr(2) = "Soil T. °C"
Info_Ch_FStr(2) = 1
info_Ch_Type(2) = "DIG"


Info_Ch_FS(3) = 1060
Info_Ch_LS(3) = 600
Info_Ch_attivo(3) = False
Info_Ch_descr(3) = "B.P. HPa"
Info_Ch_FStr(3) = 1
info_Ch_Type(3) = "0-5"

Info_Check = "PALM"
Info_RS232_Port = 1
End Sub

Function CheckNI6B11_N_Informations() As Integer
'Solo per 6B11
'Prende i valori dal form e controlla che siano validi
Dim p As String
  CheckNI6B11_N_Informations = False
  Info_SampInterval = CDbl(StartUpFrm.SICmb.Text)   'Intervallo tra un punto e l'altro in secondi
  Info_NumPoints = Int(StartUpFrm.NPCmb.Text / Info_SampInterval) 'Numero dei punti della curva
  Info_AC_K = CDbl(StartUpFrm.AckFld.Text)  'Costante della camera di accumulo
  Info_unit = StartUpFrm.UnitCmb.Text   'Unita di misura che viene da ppm/sek*ac_k

  Info_Ch_attivo(0) = True  'dice se il canale è attivo
   p = StartUpFrm.FSCmb.Text
   Select Case p
     Case "15", "50", "100", "500", "1000", "5000", "5001"
        Info_Ch_FS(0) = p
     Case Else
        StartUpFrm.FSCmb.Text = "5000"
        Info_Ch_FS(0) = "5000"
   End Select
  If p = "5001" Then
      If Info_SampInterval < 1 Then
         Info_SampInterval = 1
         StartUpFrm.FSCmb.Text = p
         MsgBox "Sample Interval to short for autoranging"
         
         Exit Function
      End If
  End If
  
  
  
  
  Info_Ch_FS(0) = CDbl(StartUpFrm.FSCmb.Text)
  
            'f(x)=a*x+b
  Info_Ch_LS(0) = 0 '-Info_Ch_ a(0) / 4
  Info_Ch_descr(0) = "NI6B11: mV " 'Descrizione del canale"
  Info_Ch_FStr(0) = 3
  info_Ch_Type(0) = "DIG"
  

  Info_Ch_attivo(1) = -1 * StartUpFrm.SoilTChk.Value 'dice se il canale è attivo
  Info_Ch_FS(1) = 200 '  'f(x)=a*x+b
  Info_Ch_LS(1) = -50 '
  Info_Ch_descr(1) = "Soil T. [°C]: " 'Descrizione del canale
  Info_Ch_FStr(1) = 1
  info_Ch_Type(1) = "DIG"

  Info_Ch_attivo(2) = -1 * StartUpFrm.AirTChk.Value 'dice se il canale è attivo
  Info_Ch_FS(2) = 200 ' As Double          'f(x)=a*x+b
  Info_Ch_LS(2) = -50 ' ' As Double
  Info_Ch_descr(2) = "Air T. [°C]: " 'As String      'Descrizione del canale
  Info_Ch_FStr(2) = 1
  info_Ch_Type(2) = "DIG"

  Info_Ch_attivo(3) = -1 * StartUpFrm.VaisalaChk.Value ' As Boolean    'dice se il canale è attivo
  Info_Ch_FS(3) = 1060 ' As Double          'f(x)=a*x+b
  Info_Ch_LS(3) = 600  ' As Double
  Info_Ch_descr(3) = "Bar. P. HPa: " 'As String      'Descrizione del canale
  Info_Ch_FStr(3) = 1
  info_Ch_Type(3) = "0-5"
  CheckNI6B11_N_Informations = True

End Function
Function CheckRikenInformations()
'Solo per 6B11
'Prende i valori dal form e controlla che siano validi
Dim p As String
  CheckRikenInformations = False
  Info_SampInterval = CDbl(StartUpFrm.SICmb.Text)   'Intervallo tra un punto e l'altro in secondi
  Info_NumPoints = Int(StartUpFrm.NPCmb.Text / Info_SampInterval) 'Numero dei punti della curva
  Info_AC_K = CDbl(StartUpFrm.AckFld.Text)  'Costante della camera di accumulo
  Info_unit = StartUpFrm.UnitCmb.Text   'Unita di misura che viene da ppm/sek*ac_k

  Info_Ch_attivo(0) = True  'dice se il canale è attivo
   p = StartUpFrm.FSCmb.Text
   Select Case p
     Case "2000", "3000", "5000"
        Info_Ch_FS(0) = p
     Case Else
        StartUpFrm.FSCmb.Text = "5000"
        Info_Ch_FS(0) = "5000"
   End Select
  
  Info_Ch_FS(0) = CDbl(StartUpFrm.FSCmb.Text)
  
            'f(x)=a*x+b
  Info_Ch_LS(0) = 0 '-Info_Ch_ a(0) / 4
  Info_Ch_descr(0) = "Riken: ppm " 'Descrizione del canale"
  Info_Ch_FStr(0) = 3
  info_Ch_Type(0) = "DIG"
  

  Info_Ch_attivo(1) = -1 * StartUpFrm.SoilTChk.Value 'dice se il canale è attivo
  Info_Ch_FS(1) = 200 '  'f(x)=a*x+b
  Info_Ch_LS(1) = -50 '
  Info_Ch_descr(1) = "Soil T. [°C]: " 'Descrizione del canale
  Info_Ch_FStr(1) = 1
  info_Ch_Type(1) = "DIG"

  Info_Ch_attivo(2) = -1 * StartUpFrm.AirTChk.Value 'dice se il canale è attivo
  Info_Ch_FS(2) = 200 ' As Double          'f(x)=a*x+b
  Info_Ch_LS(2) = -50 ' ' As Double
  Info_Ch_descr(2) = "Air T. [°C]: " 'As String      'Descrizione del canale
  Info_Ch_FStr(2) = 1
  info_Ch_Type(2) = "DIG"

  Info_Ch_attivo(3) = -1 * StartUpFrm.VaisalaChk.Value ' As Boolean    'dice se il canale è attivo
  Info_Ch_FS(3) = 1060 ' As Double          'f(x)=a*x+b
  Info_Ch_LS(3) = 600  ' As Double
  Info_Ch_descr(3) = "Bar. P. HPa: " 'As String      'Descrizione del canale
  Info_Ch_FStr(3) = 1
  info_Ch_Type(3) = "0-5"
  CheckRikenInformations = True


End Function
Function CheckLICORInformations() As Integer
'Solo per LICOR
'Prende i valori dal form e controlla che siano validi
  CheckLICORInformations = False
  
  Info_SampInterval = StartUpFrm.SICmb.Text   'Intervallo tra un punto e la'ltro in secndi
  Info_NumPoints = Int(StartUpFrm.NPCmb.Text / Info_SampInterval) 'Numero dei punti della curva
  Info_AC_K = (StartUpFrm.AckFld.Text)  'Costante della camera di accumulo
  Info_unit = StartUpFrm.UnitCmb.Text   'Unita di misura che viene da ppm/sek*ac_k

  Info_Ch_attivo(0) = True  'dice se il canale è attivo
  
  Info_Ch_FS(0) = StartUpFrm.FSCmb.Text
  
            'f(x)=a*x+b
  Info_Ch_LS(0) = 0
  Info_Ch_descr(0) = "CO2 [ppm]: " 'Descrizione del canale"
  Info_Ch_FStr(0) = 1
  info_Ch_Type(0) = "DIG"

  Info_Ch_attivo(1) = -1 * StartUpFrm.SoilTChk.Value 'dice se il canale è attivo
  Info_Ch_FS(1) = 200 '  'f(x)=a*x+b
  Info_Ch_LS(1) = -50 '
  Info_Ch_descr(1) = "Soil T. [°C]: " 'Descrizione del canale
  Info_Ch_FStr(1) = 1
  info_Ch_Type(1) = "DIG"

  Info_Ch_attivo(2) = -1 * StartUpFrm.AirTChk.Value 'dice se il canale è attivo
  Info_Ch_FS(2) = 200 ' As Double          'f(x)=a*x+b
  Info_Ch_LS(2) = -50 ' ' As Double
  Info_Ch_descr(2) = "Air T. [°C]: " 'As String      'Descrizione del canale
  Info_Ch_FStr(2) = 1
  info_Ch_Type(2) = "DIG"

  Info_Ch_attivo(3) = True ' As Boolean    'dice se il canale è attivo
  Info_Ch_FS(3) = 1 ' As Double          'f(x)=a*x+b
  Info_Ch_LS(3) = 0 ' As Double
  Info_Ch_descr(3) = "Bar. P. HPa: " 'As String      'Descrizione del canale
  Info_Ch_FStr(3) = 1
  info_Ch_Type(3) = "DIG"
  CheckLICORInformations = True
  'Info_Check = "LICOR"
End Function
Function CheckDragerInformations() As Integer
'Solo per DRAGER
'Prende i valori dal form e controlla che siano validi
  CheckDragerInformations = False
  Info_SampInterval = StartUpFrm.SICmb.Text   'Intervallo tra un punto e la'ltro in secndi
  Info_NumPoints = Int(StartUpFrm.NPCmb.Text / Info_SampInterval) 'Numero dei punti della curva
  Info_AC_K = (StartUpFrm.AckFld.Text)  'Costante della camera di accumulo
  Info_unit = StartUpFrm.UnitCmb.Text   'Unita di misura che viene da ppm/sek*ac_k

  Info_Ch_attivo(0) = True  'dice se il canale è attivo
  
  Info_Ch_FS(0) = CDbl(StartUpFrm.FSCmb.Text)
  
            'f(x)=a*x+b
  Info_Ch_LS(0) = 0 '-Info_Ch_ a(0) / 4
  
  Info_Ch_descr(0) = "CO2 [ppm]: " 'Descrizione del canale"
  Info_Ch_FStr(0) = 1
  info_Ch_Type(0) = "4-20"

  Info_Ch_attivo(1) = -1 * StartUpFrm.SoilTChk.Value 'dice se il canale è attivo
  Info_Ch_FS(1) = 200 '  'f(x)=a*x+b
  Info_Ch_LS(1) = -50 '
  Info_Ch_descr(1) = "Soil T. [°C]: " 'Descrizione del canale
  Info_Ch_FStr(1) = 1
  info_Ch_Type(1) = "DIG"


  Info_Ch_attivo(2) = -1 * StartUpFrm.AirTChk.Value 'dice se il canale è attivo
  Info_Ch_FS(2) = 200 ' As Double          'f(x)=a*x+b
  Info_Ch_LS(2) = -50 ' ' As Double
  Info_Ch_descr(2) = "Air T. [°C]: " 'As String      'Descrizione del canale
  Info_Ch_FStr(2) = 1
  info_Ch_Type(2) = "DIG"
  Info_Ch_attivo(3) = -1 * StartUpFrm.VaisalaChk.Value ' As Boolean    'dice se il canale è attivo
  Info_Ch_FS(3) = 1060 ' As Double          'f(x)=a*x+b
  Info_Ch_LS(3) = 600  ' As Double
  Info_Ch_descr(3) = "Bar. P. HPa: " 'As String      'Descrizione del canale
  Info_Ch_FStr(3) = 1
  info_Ch_Type(3) = "0-5"
  CheckDragerInformations = True
 ' Info_Check = "DRAGER"
End Function
Function CheckPolitronInformations() As Integer
'Solo per DRAGER
'Prende i valori dal form e controlla che siano validi
  CheckPolitronInformations = False
  Info_SampInterval = StartUpFrm.SICmb.Text   'Intervallo tra un punto e la'ltro in secndi
  Info_NumPoints = Int(StartUpFrm.NPCmb.Text / Info_SampInterval) 'Numero dei punti della curva
  Info_AC_K = (StartUpFrm.AckFld.Text)  'Costante della camera di accumulo
  Info_unit = StartUpFrm.UnitCmb.Text   'Unita di misura che viene da ppm/sek*ac_k

  Info_Ch_attivo(0) = True  'dice se il canale è attivo
  
  Info_Ch_FS(0) = StartUpFrm.FSCmb.Text
  
            'f(x)=a*x+b
  Info_Ch_LS(0) = 0 '-Info_Ch_ a(0) / 4
  Info_Ch_descr(0) = "CO2 [ppm]: " 'Descrizione del canale"
  Info_Ch_FStr(0) = 1
  info_Ch_Type(0) = "4-20"
  

  Info_Ch_attivo(1) = -1 * StartUpFrm.SoilTChk.Value 'dice se il canale è attivo
  Info_Ch_FS(1) = 200 '  'f(x)=a*x+b
  Info_Ch_LS(1) = -50 '
  Info_Ch_descr(1) = "Soil T. [°C]: " 'Descrizione del canale
  Info_Ch_FStr(1) = 1
  info_Ch_Type(1) = "DIG"

  Info_Ch_attivo(2) = -1 * StartUpFrm.AirTChk.Value 'dice se il canale è attivo
  Info_Ch_FS(2) = 200 ' As Double          'f(x)=a*x+b
  Info_Ch_LS(2) = -50 ' ' As Double
  Info_Ch_descr(2) = "Air T. [°C]: " 'As String      'Descrizione del canale
  Info_Ch_FStr(2) = 1
  info_Ch_Type(2) = "DIG"

  Info_Ch_attivo(3) = -1 * StartUpFrm.VaisalaChk.Value ' As Boolean    'dice se il canale è attivo
  Info_Ch_FS(3) = 1060 ' As Double          'f(x)=a*x+b
  Info_Ch_LS(3) = 600  ' As Double
  Info_Ch_descr(3) = "Bar. P. HPa: " 'As String      'Descrizione del canale
  Info_Ch_FStr(3) = 1
  info_Ch_Type(3) = "0-5"
  CheckPolitronInformations = True
End Function
Function CheckNI6B13Informations() As Integer
Dim p As String
'Solo per 6B13
'Prende i valori dal form e controlla che siano validi
  CheckNI6B13Informations = False
  Info_SampInterval = StartUpFrm.SICmb.Text   'Intervallo tra un punto e la'ltro in secndi
  Info_NumPoints = Int(StartUpFrm.NPCmb.Text / Info_SampInterval) 'Numero dei punti della curva
  Info_AC_K = (StartUpFrm.AckFld.Text)  'Costante della camera di accumulo
  Info_unit = StartUpFrm.UnitCmb.Text   'Unita di misura che viene da ppm/sek*ac_k

  Info_Ch_attivo(0) = True  'dice se il canale è attivo
  
  p = StartUpFrm.FSCmb.Text
  Select Case p
     Case "100", "200", "600"
        Info_Ch_FS(0) = p
     Case Else
        StartUpFrm.FSCmb.Text = "100"
        Info_Ch_FS(0) = "100"
  End Select
  
            'f(x)=a*x+b
  Info_Ch_LS(0) = 0 '-Info_Ch_ a(0) / 4
  Info_Ch_descr(0) = "NI6B13: °C " 'Descrizione del canale"
  Info_Ch_FStr(0) = 3
  info_Ch_Type(0) = "DIG"
  

  Info_Ch_attivo(1) = -1 * StartUpFrm.SoilTChk.Value 'dice se il canale è attivo
  Info_Ch_FS(1) = 200 '  'f(x)=a*x+b
  Info_Ch_LS(1) = -50 '
  Info_Ch_descr(1) = "Soil T. [°C]: " 'Descrizione del canale
  Info_Ch_FStr(1) = 1
  info_Ch_Type(1) = "DIG"

  Info_Ch_attivo(2) = -1 * StartUpFrm.AirTChk.Value 'dice se il canale è attivo
  Info_Ch_FS(2) = 200 ' As Double          'f(x)=a*x+b
  Info_Ch_LS(2) = -50 ' ' As Double
  Info_Ch_descr(2) = "Air T. [°C]: " 'As String      'Descrizione del canale
  Info_Ch_FStr(2) = 1
  info_Ch_Type(2) = "DIG"

  Info_Ch_attivo(3) = -1 * StartUpFrm.VaisalaChk.Value ' As Boolean    'dice se il canale è attivo
  Info_Ch_FS(3) = 1060 ' As Double          'f(x)=a*x+b
  Info_Ch_LS(3) = 600  ' As Double
  Info_Ch_descr(3) = "Bar. P. HPa: " 'As String      'Descrizione del canale
  Info_Ch_FStr(3) = 1
  info_Ch_Type(3) = "0-5"
  CheckNI6B13Informations = True


End Function

Function CheckNI6B11Informations() As Integer
'Solo per 6B11
'Prende i valori dal form e controlla che siano validi
Dim p As String
  CheckNI6B11Informations = False
  Info_SampInterval = CDbl(StartUpFrm.SICmb.Text)   'Intervallo tra un punto e la'ltro in secndi
  Info_NumPoints = Int(StartUpFrm.NPCmb.Text / Info_SampInterval) 'Numero dei punti della curva
  Info_AC_K = CDbl(StartUpFrm.AckFld.Text)  'Costante della camera di accumulo
  Info_unit = StartUpFrm.UnitCmb.Text   'Unita di misura che viene da ppm/sek*ac_k

  Info_Ch_attivo(0) = True  'dice se il canale è attivo
   p = StartUpFrm.FSCmb.Text
   Select Case p
     Case "15", "50", "100", "500", "1000", "5000", "5001"
        Info_Ch_FS(0) = p
     Case Else
        StartUpFrm.FSCmb.Text = "5000"
        Info_Ch_FS(0) = "5000"
   End Select
  If p = "5001" Then
      If Info_SampInterval < 1 Then
         Info_SampInterval = 1
         StartUpFrm.FSCmb.Text = p
         MsgBox "Sample Interval to short for autoranging"
         
         Exit Function
      End If
  End If
  
  
  
  
  Info_Ch_FS(0) = CDbl(StartUpFrm.FSCmb.Text)
  
            'f(x)=a*x+b
  Info_Ch_LS(0) = 0 '-Info_Ch_ a(0) / 4
  Info_Ch_descr(0) = "NI6B11: mV " 'Descrizione del canale"
  Info_Ch_FStr(0) = 3
  info_Ch_Type(0) = "DIG"
  

  Info_Ch_attivo(1) = -1 * StartUpFrm.SoilTChk.Value 'dice se il canale è attivo
  Info_Ch_FS(1) = 200 '  'f(x)=a*x+b
  Info_Ch_LS(1) = -50 '
  Info_Ch_descr(1) = "Soil T. [°C]: " 'Descrizione del canale
  Info_Ch_FStr(1) = 1
  info_Ch_Type(1) = "DIG"

  Info_Ch_attivo(2) = -1 * StartUpFrm.AirTChk.Value 'dice se il canale è attivo
  Info_Ch_FS(2) = 200 ' As Double          'f(x)=a*x+b
  Info_Ch_LS(2) = -50 ' ' As Double
  Info_Ch_descr(2) = "Air T. [°C]: " 'As String      'Descrizione del canale
  Info_Ch_FStr(2) = 1
  info_Ch_Type(2) = "DIG"

  Info_Ch_attivo(3) = -1 * StartUpFrm.VaisalaChk.Value ' As Boolean    'dice se il canale è attivo
  Info_Ch_FS(3) = 1060 ' As Double          'f(x)=a*x+b
  Info_Ch_LS(3) = 600  ' As Double
  Info_Ch_descr(3) = "Bar. P. HPa: " 'As String      'Descrizione del canale
  Info_Ch_FStr(3) = 1
  info_Ch_Type(3) = "0-5"
  CheckNI6B11Informations = True

End Function
Sub ShowConfiguration()
 
 If (Info_Check <> Licor) And (Info_Check <> Drager) And (Info_Check <> Politron) And (Info_Check <> Riken) And (Info_Check <> NI6B11) And (Info_Check <> NI6B13) And (Info_Check <> NI6B11_N) Then
    MsgBox "ERR:003 Attention: Not supported: " + Info_Check + "run DetectorSETUP", , "WEST Systems"
 End If
  StartUpFrm.FSCmb.Clear
  StartUpFrm.SICmb.Clear
  StartUpFrm.VaisalaChk.Visible = True
  StartUpFrm.SICmb.Enabled = True
        
  Select Case Info_Check
    Case Licor
        StartUpFrm.FSCmb.AddItem "20000"
        StartUpFrm.FSCmb.AddItem "10000"
        StartUpFrm.FSCmb.AddItem "5000"
        StartUpFrm.FSCmb.AddItem "2000"
        StartUpFrm.FSCmb.AddItem "1000"
        
        If Info_Check <> Old_Info_Check Then
            StartUpFrm.FSCmb.Text = "20000"
            StartUpFrm.SICmb.Text = "1"
        Else
            StartUpFrm.FSCmb.Text = FormatNumber(Info_Ch_FS(0), 0)
            StartUpFrm.SICmb.Text = FormatNumber(Info_SampInterval, 1)
        End If
        
        
        StartUpFrm.SICmb.AddItem "1"
        StartUpFrm.SICmb.Text = "1"
        StartUpFrm.SICmb.Enabled = False
        StartUpFrm.VaisalaChk.Value = False
        StartUpFrm.VaisalaChk.Visible = False
    
    Case Drager
        StartUpFrm.FSCmb.AddItem "3000"
        StartUpFrm.FSCmb.AddItem "3000"
        StartUpFrm.FSCmb.AddItem "5000"
        StartUpFrm.FSCmb.AddItem "7500"
        StartUpFrm.FSCmb.AddItem "9999"
        StartUpFrm.FSCmb.AddItem "10000"
        StartUpFrm.FSCmb.AddItem "20000"
        StartUpFrm.FSCmb.AddItem "30000"
        StartUpFrm.FSCmb.AddItem "40000"
        StartUpFrm.FSCmb.AddItem "50000"
        StartUpFrm.FSCmb.AddItem "70000"
        StartUpFrm.FSCmb.AddItem "100000"
        StartUpFrm.FSCmb.AddItem "200000"
        StartUpFrm.FSCmb.AddItem "500000"
        StartUpFrm.FSCmb.AddItem "1000000"
        
       
        StartUpFrm.SICmb.AddItem "2.0"
        StartUpFrm.SICmb.AddItem "1.0"
        StartUpFrm.SICmb.AddItem "0.8"
        StartUpFrm.SICmb.AddItem "0.6"
        StartUpFrm.SICmb.AddItem "0.4"
        StartUpFrm.SICmb.AddItem "0.2"
        
        If Info_Check <> Old_Info_Check Then
            StartUpFrm.FSCmb.Text = "10000"
            StartUpFrm.SICmb.Text = "1"
        Else
            StartUpFrm.FSCmb.Text = FormatNumber(Info_Ch_FS(0), 0)
            StartUpFrm.SICmb.Text = FormatNumber(Info_SampInterval, 1)
        End If
                
    Case NI6B11
        StartUpFrm.FSCmb.AddItem "5001"
        StartUpFrm.FSCmb.AddItem "5000"
        StartUpFrm.FSCmb.AddItem "1000"
        StartUpFrm.FSCmb.AddItem "500"
        StartUpFrm.FSCmb.AddItem "100"
        StartUpFrm.FSCmb.AddItem "50"
        StartUpFrm.FSCmb.AddItem "15"
        
        
       
        StartUpFrm.SICmb.AddItem "2.0"
        StartUpFrm.SICmb.AddItem "1.0"
        StartUpFrm.SICmb.AddItem "0.8"
        StartUpFrm.SICmb.AddItem "0.6"
        StartUpFrm.SICmb.AddItem "0.4"
        StartUpFrm.SICmb.AddItem "0.2"
        
        
        If Info_Check <> Old_Info_Check Then
            StartUpFrm.FSCmb.Text = "5000"
            StartUpFrm.SICmb.Text = "1"
        Else
            StartUpFrm.FSCmb.Text = FormatNumber(Info_Ch_FS(0), 0)
            StartUpFrm.SICmb.Text = FormatNumber(Info_SampInterval, 1)
        End If
    Case Riken
        StartUpFrm.FSCmb.AddItem "4975"
        StartUpFrm.FSCmb.AddItem "5000"
        StartUpFrm.FSCmb.AddItem "3000"
        StartUpFrm.FSCmb.AddItem "2000"
        
        
        
        StartUpFrm.SICmb.AddItem "4.0"
        StartUpFrm.SICmb.AddItem "3.0"
        StartUpFrm.SICmb.AddItem "2.0"
        StartUpFrm.SICmb.AddItem "1.0"
        
        
        If Info_Check <> Old_Info_Check Then
            StartUpFrm.FSCmb.Text = "4975"
            StartUpFrm.SICmb.Text = "1"
        Else
            StartUpFrm.FSCmb.Text = FormatNumber(Info_Ch_FS(0), 0)
            StartUpFrm.SICmb.Text = FormatNumber(Info_SampInterval, 1)
        End If
        
    
    
    
    
    Case NI6B11_N
        StartUpFrm.FSCmb.AddItem "5000"
        StartUpFrm.FSCmb.AddItem "1000"
        StartUpFrm.FSCmb.AddItem "500"
        StartUpFrm.FSCmb.AddItem "100"
        StartUpFrm.FSCmb.AddItem "50"
        StartUpFrm.FSCmb.AddItem "15"
       
        StartUpFrm.SICmb.AddItem "2.0"
        StartUpFrm.SICmb.AddItem "1.0"
        StartUpFrm.SICmb.AddItem "0.8"
        StartUpFrm.SICmb.AddItem "0.6"
        StartUpFrm.SICmb.AddItem "0.4"
        StartUpFrm.SICmb.AddItem "0.2"
        
        
        If Info_Check <> Old_Info_Check Then
            StartUpFrm.FSCmb.Text = "5000"
            StartUpFrm.SICmb.Text = "1"
        Else
            StartUpFrm.FSCmb.Text = FormatNumber(Info_Ch_FS(0), 0)
            StartUpFrm.SICmb.Text = FormatNumber(Info_SampInterval, 1)
        End If
        
    
    Case NI6B13
        StartUpFrm.FSCmb.AddItem "100"
        StartUpFrm.FSCmb.AddItem "200"
        StartUpFrm.FSCmb.AddItem "600"
        
        
        
        StartUpFrm.SICmb.AddItem "2.0"
        StartUpFrm.SICmb.AddItem "1.0"
        StartUpFrm.SICmb.AddItem "0.8"
        StartUpFrm.SICmb.AddItem "0.6"
        StartUpFrm.SICmb.AddItem "0.4"
        StartUpFrm.SICmb.AddItem "0.2"
    
        If Info_Check <> Old_Info_Check Then
            StartUpFrm.FSCmb.Text = "100"
            StartUpFrm.SICmb.Text = "1"
        Else
            StartUpFrm.FSCmb.Text = FormatNumber(Info_Ch_FS(0), 0)
            StartUpFrm.SICmb.Text = FormatNumber(Info_SampInterval, 1)
        End If
    
    
    
    Case Politron
        StartUpFrm.FSCmb.AddItem "10"
        StartUpFrm.FSCmb.AddItem "20"
        StartUpFrm.FSCmb.AddItem "50"
        StartUpFrm.FSCmb.AddItem "100"
        StartUpFrm.FSCmb.AddItem "1000"
        
    
        StartUpFrm.SICmb.AddItem "1.0"
        StartUpFrm.SICmb.AddItem "0.8"
        StartUpFrm.SICmb.AddItem "0.6"
        StartUpFrm.SICmb.AddItem "0.4"
        StartUpFrm.SICmb.AddItem "0.2"
        
        If Info_Check <> Old_Info_Check Then
            StartUpFrm.FSCmb.Text = "100"
            StartUpFrm.SICmb.Text = "1"
        Else
            StartUpFrm.FSCmb.Text = FormatNumber(Info_Ch_FS(0), 0)
            StartUpFrm.SICmb.Text = FormatNumber(Info_SampInterval, 1)
        End If

    End Select
    
  StartUpFrm.NPCmb.Clear
  StartUpFrm.NPCmb.AddItem "720"
  StartUpFrm.NPCmb.AddItem "360"
  StartUpFrm.NPCmb.AddItem "180"
  StartUpFrm.NPCmb.AddItem "120"
  StartUpFrm.NPCmb.AddItem "90"
  StartUpFrm.NPCmb.Text = FormatNumber(Info_NumPoints * Info_SampInterval, 0)
  
  StartUpFrm.AckFld.Text = FormatNumber(Info_AC_K, 2) 'As Double             'Costante della camera di accumulo
  StartUpFrm.UnitCmb.Clear
  
  Select Case Info_Check
    Case Licor, Politron, Drager
        StartUpFrm.UnitCmb.AddItem "gr/(m^2 day)"
        StartUpFrm.UnitCmb.AddItem "moles/(m^2 day)"
        StartUpFrm.UnitCmb.AddItem "ppm/sec"
        StartUpFrm.UnitCmb.AddItem "cm/sec"
        
        If Info_Check = Old_Info_Check Then
            StartUpFrm.UnitCmb.Text = Info_unit
        Else
            StartUpFrm.UnitCmb.Text = "gr/(m^2 day)"
        End If
    
    Case NI6B11
        Info_unit = "mV/sec"
        StartUpFrm.UnitCmb.AddItem "mV/sec"
        StartUpFrm.UnitCmb.Text = "mV/sec"
        StartUpFrm.AckFld = "1"
    
    Case NI6B11_N
        Info_unit = "mV"
        StartUpFrm.UnitCmb.AddItem "mV"
        StartUpFrm.UnitCmb.Text = "mV"
        StartUpFrm.AckFld = "1"
        StartUpFrm.VaisalaChk.Visible = False
    Case NI6B13
        Info_unit = "°C/sec"
        StartUpFrm.UnitCmb.AddItem "°C/sec"
        StartUpFrm.UnitCmb.Text = "°C/sec"
        StartUpFrm.AckFld = "1"
   
    Case Riken
        StartUpFrm.UnitCmb.AddItem "gr/(m^2 day)"
        StartUpFrm.UnitCmb.AddItem "moles/(m^2 day)"
        StartUpFrm.UnitCmb.AddItem "ppm/sec"
        StartUpFrm.UnitCmb.AddItem "cm/sec"
        
        If Info_Check = Old_Info_Check Then
            StartUpFrm.UnitCmb.Text = Info_unit
        Else
            StartUpFrm.UnitCmb.Text = "gr/(m^2 day)"
        End If
        StartUpFrm.VaisalaChk.Visible = False
   
   End Select
    
    
  'Info_unit = StartUpFrm.UnitCmb.List(StartUpFrm.UnitCmb.ListIndex)      'As String             'Unita di misura che viene da ppm/sek*ac_k

  If Info_Ch_attivo(1) Then
        StartUpFrm.SoilTChk.Value = 1
  Else
        StartUpFrm.SoilTChk.Value = 0
  End If
  
  If Info_Ch_attivo(2) Then
        StartUpFrm.AirTChk.Value = 1
  Else
        StartUpFrm.AirTChk.Value = 0
  End If
  
  If Info_Ch_attivo(3) Then
     StartUpFrm.VaisalaChk.Value = 1
  Else
     StartUpFrm.VaisalaChk.Value = 0
  End If
  StartUpFrm.RS232Cmb.Clear
  StartUpFrm.RS232Cmb.AddItem "1"
  StartUpFrm.RS232Cmb.AddItem "2"
  StartUpFrm.RS232Cmb.AddItem "3"
  StartUpFrm.RS232Cmb.AddItem "4"
  If (Info_RS232_Port >= 1) And (Info_RS232_Port <= 4) Then
    StartUpFrm.RS232Cmb.Text = FormatNumber(Info_RS232_Port, 0)
  Else
    StartUpFrm.RS232Cmb.Text = "1"
  End If
  
End Sub

