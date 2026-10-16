Public MustInherit Class Connection

    Public Enum ConnectionType
        SBO
        SQL
        HANA
    End Enum

    Protected _Name As String
    Protected _DatabaseName As String

    Protected Sub New(ByVal Name As String, ByVal DatabaseName As String)
        Me._Name = Name
        Me._DatabaseName = DatabaseName
    End Sub

    Public Function Name() As String
        Return _Name
    End Function

    Public Function DatabaseName() As String
        Return _DatabaseName
    End Function

    Public MustOverride Function SQLHandler() As Queries

    Public MustOverride Function Valid() As Boolean
    Public MustOverride Function Connected() As Boolean
    Public MustOverride Function Connect() As Boolean
    Public MustOverride Function Disconnect() As Boolean
    Public MustOverride Function toXML(ByVal xmlDocument As System.Xml.XmlDocument) As System.Xml.XmlNode
    Public MustOverride Function clone() As Connection

    Public MustOverride Function executeQuery(ByVal sql As String) As DataTable
    Public MustOverride Function executeQueryRS(ByVal sql As String) As ResultSet
    Public MustOverride Function executeNonQuery(ByVal sql As String) As Integer
    Public MustOverride Function executeScalar(ByVal sql As String) As Object

    Public MustOverride Sub BeginTransaction()
    Public MustOverride Sub Commit()
    Public MustOverride Sub Rollback()

    Public MustOverride Function Type() As Integer

End Class
