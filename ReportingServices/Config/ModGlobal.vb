Public Module ModGlobal

    Public SQLHandler As Queries

    Public WinDecSep As String
    Public SBODecSep As String
    Public SBOThousSep As String

    Public oApplicationConfig As ApplicationConfig


    Public isInit As Boolean = False

    Public UserCode As String
    Public GlobalUserName As String
    Public GlobalWeighbridge As String
    Public GlobalWeighbridgeBackdate As String
    Public UserBranch As String
    Public Params() As String
    Public oConfigXMLParser As New ConfigXMLParser
    Public oSBOConnection As SBOConnection
    Public oDBConnection As Connection
    Public oLogConnection As Connection
    Public errori As Integer
    Public Global_tipeP As String = ""
    Public date_now As Date
    Public time_now As Integer
    Public toleransi As Double
    Public toleransiWB As Double
    Public min As Double
    Public max As Double
    Public load_stat As String
    Public dt_shift As New DataTable

End Module
