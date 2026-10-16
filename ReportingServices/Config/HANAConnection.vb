Public Class HANAConnection
    Inherits Connection

    Private _ServerName As String
    Private _Port As String
    Private _UserName As String
    Private _Password As String

    Private _SQLHandler As Queries

    Private _Connection As Sap.Data.Hana.HanaConnection
    Private _Transaction As Sap.Data.Hana.HanaTransaction

    Public Sub New(ByVal Name As String, ByVal ServerName As String, ByVal Port As String, ByVal DatabaseName As String, ByVal UserName As String, ByVal Password As String)
        MyBase.New(Name, DatabaseName)

        Me._ServerName = ServerName
        Me._Port = Port
        Me._UserName = UserName
        Me._Password = Encryption.Decrypt(Password)
    End Sub

#Region "Properties"

    Public Function ServerName() As String
        Return _ServerName
    End Function

    Public Function Port() As String
        Return _Port
    End Function

    Public Function UserName() As String
        Return _UserName
    End Function

    Public Function Password() As String
        Return _Password
    End Function

    Public Overrides Function SQLHandler() As Queries
        Return _SQLHandler
    End Function

    Public Function Connection() As Sap.Data.Hana.HanaConnection
        Return _Connection
    End Function

#End Region

#Region "Connection"

    Public Overrides Function Valid() As Boolean
        If (Name() Is Nothing Or String.Equals(Trim(Name), "")) Then
            Return False
        ElseIf (ServerName() Is Nothing Or String.Equals(Trim(ServerName), "")) Then
            Return False
        ElseIf (Port() Is Nothing Or String.Equals(Trim(Port), "")) Then
            Return False
        ElseIf (DatabaseName() Is Nothing Or String.Equals(Trim(DatabaseName), "")) Then
            Return False
        ElseIf (UserName() Is Nothing Or String.Equals(Trim(UserName), "")) Then
            Return False
        ElseIf (Password() Is Nothing Or String.Equals(Trim(Password), "")) Then
            Return False
        End If

        Return True
    End Function

    Public Overrides Function Connected() As Boolean
        If Not Connection() Is Nothing Then
            If Connection.State = ConnectionState.Open Then
                Return True
            End If
        End If

        Return False
    End Function

    Public Overrides Function Connect() As Boolean
        If Not Valid() Then
            Return False
        End If

        If Connected() Then
            Return True
        End If

        Dim ConnectionString As String

        ConnectionString = "Server=" + ServerName() + ":" + Port() +
                            ";UserID=" + UserName() +
                            ";Password=" + Password() + ";"

        Try
            _Connection = New Sap.Data.Hana.HanaConnection(ConnectionString)
            Connection.Open()

            _SQLHandler = New HANAVersion(DatabaseName)

            Return True
        Catch ex As Exception
            _Connection = Nothing
            Throw New Exception("Connection " + ToString() + " : " + ex.Message)
        End Try

        Return False
    End Function

    Public Overrides Function Disconnect() As Boolean
        Try
            If Connection() Is Nothing Then
                Return True
            Else
                If Connected() Then
                    Connection.Close()
                End If

                Return True
            End If
        Catch ex As Exception
            _Connection = Nothing
            Throw New Exception("Connection " + ToString() + " : " + ex.Message)
        End Try

        Return False
    End Function

#End Region

#Region "Execution"

    Public Overrides Function executeNonQuery(ByVal sql As String) As Integer
        If Not Connected() Then
            Throw New Exception("Connection " + ToString() + " is not connected")
        End If

        Dim command As New Sap.Data.Hana.HanaCommand(sql, Connection)

        If Not IsNothing(_Transaction) Then
            command.Transaction = _Transaction
        End If

        Return command.ExecuteNonQuery
    End Function

    Public Overrides Function executeScalar(ByVal sql As String) As Object
        If Not Connected() Then
            Throw New Exception("Connection " + ToString() + " is not connected")
        End If

        Dim command As New Sap.Data.Hana.HanaCommand(sql, Connection)

        If Not IsNothing(_Transaction) Then
            command.Transaction = _Transaction
        End If

        Return command.ExecuteScalar
    End Function

    Public Overrides Function executeQueryRS(ByVal sql As String) As ResultSet
        If Not Connected() Then
            Throw New Exception("Connection " + ToString() + " is not connected")
        End If

        Dim command As New Sap.Data.Hana.HanaCommand(sql, Connection)

        If Not IsNothing(_Transaction) Then
            command.Transaction = _Transaction
        End If

        Dim dataReader As Sap.Data.Hana.HanaDataReader
        dataReader = command.ExecuteReader()

        Dim headerList As New List(Of Object)

        For i As Integer = 0 To dataReader.FieldCount - 1
            headerList.Add(dataReader.GetName(i))
        Next

        Dim dataList As New List(Of List(Of Object))

        While dataReader.Read And dataReader.HasRows
            Dim innerDataList As New List(Of Object)

            For i As Integer = 0 To dataReader.FieldCount - 1
                innerDataList.Add(dataReader.GetValue(i))
            Next

            dataList.Add(innerDataList)
        End While

        Return New ResultSet(headerList, dataList)
    End Function

    Public Overrides Function executeQuery(ByVal sql As String) As DataTable
        If Not Connected() Then
            Throw New Exception("Connection " + ToString() + " is not connected")
        End If

        Dim command As New Sap.Data.Hana.HanaCommand(sql, Connection)

        If Not IsNothing(_Transaction) Then
            command.Transaction = _Transaction
        End If

        Dim da As New Sap.Data.Hana.HanaDataAdapter(command)
        Dim dt As New DataTable


        da.Fill(dt)

        Return dt
    End Function

#End Region

#Region "Transaction"

    Public Overrides Sub BeginTransaction()
        Rollback()
        _Transaction = Connection.BeginTransaction
    End Sub

    Public Overrides Sub Commit()
        If Not IsNothing(_Transaction) Then
            _Transaction.Commit()
        End If

        _Transaction = Nothing
    End Sub

    Public Overrides Sub Rollback()
        If Not IsNothing(_Transaction) Then
            _Transaction.Rollback()
        End If

        _Transaction = Nothing
    End Sub

#End Region

#Region "Inheritance"

    Public Overrides Function toXML(ByVal xmlDocument As System.Xml.XmlDocument) As System.Xml.XmlNode
        Dim xmlStagingDatabase As System.Xml.XmlNode = xmlDocument.CreateElement("HANA_DATABASE")

        Dim xmlName As System.Xml.XmlNode = xmlDocument.CreateElement("NAME")
        xmlName.InnerText = Trim(Name)
        xmlStagingDatabase.AppendChild(xmlName)

        Dim xmlServerName As System.Xml.XmlNode = xmlDocument.CreateElement("SERVER_NAME")
        xmlServerName.InnerText = Trim(ServerName)
        xmlStagingDatabase.AppendChild(xmlServerName)

        Dim xmlPort As System.Xml.XmlNode = xmlDocument.CreateElement("PORT")
        xmlPort.InnerText = Trim(Port)
        xmlStagingDatabase.AppendChild(xmlPort)

        Dim xmlDatabaseName As System.Xml.XmlNode = xmlDocument.CreateElement("DATABASE_NAME")
        xmlDatabaseName.InnerText = Trim(DatabaseName)
        xmlStagingDatabase.AppendChild(xmlDatabaseName)

        Dim xmlUserName As System.Xml.XmlNode = xmlDocument.CreateElement("USER_NAME")
        xmlUserName.InnerText = Trim(UserName)
        xmlStagingDatabase.AppendChild(xmlUserName)

        Dim xmlPassword As System.Xml.XmlNode = xmlDocument.CreateElement("PASSWORD")
        xmlPassword.InnerText = Encryption.Encrypt(Trim(Password))
        xmlStagingDatabase.AppendChild(xmlPassword)

        Return xmlStagingDatabase
    End Function

    Public Overrides Function Type() As Integer
        Return ConnectionType.HANA
    End Function

    Public Overrides Function ToString() As String
        Dim text As String
        text = "HANA : " + ServerName() + ":" + Port() + " ; " + DatabaseName()
        text = text.ToUpper

        Return text
    End Function

    Public Overrides Function clone() As Connection
        Return New HANAConnection(Name, ServerName, Port, DatabaseName, UserName, Encryption.Encrypt(Password))
    End Function

#End Region

End Class
