Attribute VB_Name = "Modem"
'Aggiunte per la comunicazione via modem
Type TipoConnessione
    Locale As Boolean
    nTelefono As String
    Manuale As Boolean
    Ora As String
    Password As String
    ComPort As Integer
    ModemString As String
    PortConfiguration As String
End Type

'Variabile che contiene il tipo di connessione
Public CfgCon As TipoConnessione

