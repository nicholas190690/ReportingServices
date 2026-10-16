Public Class SBOConnection
    Inherits Connection

    Private _ServerName As String
    Private _Port As String

    Private _DBUserName As String
    Private _DBPassword As String
    Private _DBServerType As String

    Private _SBOUserName As String
    Private _SBOPassword As String
    Private _SBOLicenseServer As String
    Private _SBOAddOnIdentifier As String

    Private _SQLHandler As Queries

    Private _Company As SAPbobsCOM.Company



    Public Sub New(ByVal Name As String, ByVal ServerName As String, ByVal Port As String, ByVal DBUserName As String, ByVal DBPassword As String, ByVal DBServerType As String,
                    ByVal SBODatabaseName As String, ByVal SBOUserName As String, ByVal SBOPassword As String, ByVal SBOLicenseServer As String, ByVal SBOAddOnIdentifier As String)
        MyBase.New(Name, SBODatabaseName)

        Me._ServerName = ServerName
        Me._Port = Port
        Me._DBUserName = DBUserName
        Me._DBPassword = Encryption.Decrypt(DBPassword)
        Me._DBServerType = DBServerType
        Me._SBOUserName = SBOUserName
        Me._SBOPassword = Encryption.Decrypt(SBOPassword)
        Me._SBOLicenseServer = SBOLicenseServer
        Me._SBOAddOnIdentifier = SBOAddOnIdentifier
    End Sub

#Region "Properties"

    Public Function ServerName() As String
        Return _ServerName
    End Function

    Public Function Port() As String
        Return _Port
    End Function

    Public Function DBUserName() As String
        Return _DBUserName
    End Function

    Public Function DBPassword() As String
        Return _DBPassword
    End Function

    Public Function DBServerType() As String
        Return _DBServerType
    End Function

    Public Function SBODatabaseName() As String
        Return DatabaseName()
    End Function

    Public Function SBOUserName() As String
        Return _SBOUserName
    End Function

    Public Function SBOPassword() As String
        Return _SBOPassword
    End Function

    Public Function SBOLicenseServer() As String
        Return _SBOLicenseServer
    End Function

    Public Function SBOAddOnIdentifier() As String
        Return _SBOAddOnIdentifier
    End Function

    Public Overrides Function SQLHandler() As Queries
        Return _SQLHandler
    End Function

    Public Function Company() As SAPbobsCOM.Company
        Return _Company
    End Function

#End Region

#Region "Connection"

    Public Overrides Function Valid() As Boolean
        If (Name() Is Nothing Or String.Equals(Trim(Name), "")) Then
            Return False
        ElseIf (ServerName() Is Nothing Or String.Equals(Trim(ServerName), "")) Then
            Return False
        ElseIf (DBUserName() Is Nothing Or String.Equals(Trim(DBUserName), "")) Then
            Return False
        ElseIf (DBPassword() Is Nothing Or String.Equals(Trim(DBPassword), "")) Then
            Return False
        ElseIf (DBServerType() Is Nothing Or String.Equals(Trim(DBServerType), "")) Then
            Return False
        ElseIf (SBODatabaseName() Is Nothing Or String.Equals(Trim(SBODatabaseName), "")) Then
            Return False
        ElseIf (SBOUserName() Is Nothing Or String.Equals(Trim(SBOUserName), "")) Then
            Return False
        ElseIf (SBOPassword() Is Nothing Or String.Equals(Trim(SBOPassword), "")) Then
            Return False
        ElseIf (SBOLicenseServer() Is Nothing Or String.Equals(Trim(SBOLicenseServer), "")) Then
            Return False
        End If

        Return True
    End Function

    Public Overrides Function Connected() As Boolean
        If Not Company() Is Nothing Then
            If Company.Connected Then
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

        _Company = New SAPbobsCOM.Company

        Dim iError As Integer
        Dim iError2 As Integer
        Dim sErrMsg As String = ""

        Try
            If Company.Connected = False Then
                If String.Equals("2005", DBServerType) Then
                    Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
                ElseIf String.Equals("2008", DBServerType) Then
                    Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
                ElseIf String.Equals("2012", DBServerType) Then
                    Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
                ElseIf String.Equals("2014", DBServerType) Then
                    Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
                ElseIf String.Equals("HANA", DBServerType) Then
                    Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
                End If

                If Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Company.Server = ServerName() + ":" + Port()
                Else
                    Company.Server = ServerName()
                End If

                Company.DbUserName = DBUserName()
                Company.DbPassword = DBPassword()

                Company.CompanyDB = SBODatabaseName()
                Company.UserName = SBOUserName()
                Company.Password = SBOPassword()
                Company.LicenseServer = SBOLicenseServer()
                Company.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                If IsNothing(SBOAddOnIdentifier) Then
                    _SBOAddOnIdentifier = ""
                End If

                If SBOAddOnIdentifier.Trim.Length > 0 Then
                    Company.AddonIdentifier = SBOAddOnIdentifier()
                End If

                iError = Company.Connect()
                If iError <> 0 Then
                    If iError = 100000085 Then
                        Company.Disconnect()
                    End If
                    Company.GetLastError(iError, sErrMsg)
                    Throw New Exception(iError.ToString + " - " + sErrMsg)
                Else
                    sErrMsg = "Connected to SAP Database - " & SBODatabaseName()

                    If Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        _SQLHandler = New HANAVersion(SBODatabaseName)
                    Else
                        _SQLHandler = New SQLVersion(SBODatabaseName)
                    End If

                    Return True
                End If
            Else
                sErrMsg = "SAP Business One already connected to - " & SBODatabaseName()
                Return True
            End If
        Catch ex As Exception
            _Company = Nothing
            Throw New Exception(ex.Message)
        End Try

        Return False
    End Function

    Public Overrides Function Disconnect() As Boolean
        If Company() Is Nothing Then
            Return True
        End If

        Try
            If Company.Connected Then
                Company.Disconnect()
                Return True
            End If
        Catch ex As Exception
            _Company = Nothing
            Throw New Exception("Connection " + ToString() + " : " + ex.Message)
        End Try

        Return True
    End Function

#End Region

#Region "Execution"

    Public Overrides Function executeNonQuery(ByVal sql As String) As Integer
        If Not Connected() Then
            Throw New Exception("Connection " + ToString() + " is not connected")
        End If

        Dim oRec As SAPbobsCOM.Recordset
        oRec = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oRec.DoQuery(sql)

        Return oRec.RecordCount
    End Function

    Public Overrides Function executeScalar(ByVal sql As String) As Object
        If Not Connected() Then
            Throw New Exception("Connection " + ToString() + " is not connected")
        End If

        Dim oRec As SAPbobsCOM.Recordset
        oRec = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oRec.DoQuery(sql)

        If oRec.RecordCount > 0 Then
            oRec.MoveFirst()
            Return oRec.Fields.Item(0).Value
        End If

        Return Nothing
    End Function

    Public Overrides Function executeQueryRS(ByVal sql As String) As ResultSet
        If Not Connected() Then
            Throw New Exception("Connection " + ToString() + " is not connected")
        End If

        Dim oRec As SAPbobsCOM.Recordset
        oRec = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oRec.DoQuery(sql)

        Dim headerList As New List(Of Object)

        For i As Integer = 0 To oRec.Fields.Count - 1
            headerList.Add(oRec.Fields.Item(i).Name)
        Next

        Dim dataList As New List(Of List(Of Object))

        While Not oRec.EoF
            Dim innerDataList As New List(Of Object)

            For i As Integer = 0 To oRec.Fields.Count - 1
                innerDataList.Add(oRec.Fields.Item(i).Value)
            Next

            dataList.Add(innerDataList)
        End While

        Return New ResultSet(headerList, dataList)
    End Function

    Public Overrides Function executeQuery(ByVal sql As String) As DataTable
        If Not Connected() Then
            Throw New Exception("Connection " + ToString() + " is not connected")
        End If

        Dim oRec As SAPbobsCOM.Recordset
        oRec = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oRec.DoQuery(sql)

        Dim dt As New DataTable

        For i As Integer = 0 To oRec.Fields.Count - 1
            dt.Columns.Add(oRec.Fields.Item(i).Name)
        Next

        While Not oRec.EoF
            Dim innerDataList As New List(Of Object)

            For i As Integer = 0 To oRec.Fields.Count - 1
                innerDataList.Add(oRec.Fields.Item(i).Value)
            Next

            dt.Rows.Add(innerDataList.ToArray)
            oRec.MoveNext()
        End While

        Return dt
    End Function

#End Region

#Region "Transaction"

    Public Overrides Sub BeginTransaction()
        Rollback()
        Company.StartTransaction()
    End Sub

    Public Overrides Sub Commit()
        If Company.InTransaction Then
            Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
    End Sub

    Public Overrides Sub Rollback()
        If Company.InTransaction Then
            Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        End If
    End Sub

#End Region

#Region "Inheritance"

    Public Overrides Function toXML(ByVal xmlDocument As System.Xml.XmlDocument) As System.Xml.XmlNode
        Dim xmlSBODatabase As System.Xml.XmlNode = xmlDocument.CreateElement("SBO_DATABASE")

        Dim xmlName As System.Xml.XmlNode = xmlDocument.CreateElement("NAME")
        xmlName.InnerText = Trim(Name)
        xmlSBODatabase.AppendChild(xmlName)

        Dim xmlServerName As System.Xml.XmlNode = xmlDocument.CreateElement("SERVER_NAME")
        xmlServerName.InnerText = Trim(ServerName)
        xmlSBODatabase.AppendChild(xmlServerName)

        Dim xmlPort As System.Xml.XmlNode = xmlDocument.CreateElement("PORT")
        xmlPort.InnerText = Trim(Port)
        xmlSBODatabase.AppendChild(xmlPort)

        Dim xmlDBUserName As System.Xml.XmlNode = xmlDocument.CreateElement("DATABASE_USER_NAME")
        xmlDBUserName.InnerText = Trim(DBUserName)
        xmlSBODatabase.AppendChild(xmlDBUserName)

        Dim xmlDBPassword As System.Xml.XmlNode = xmlDocument.CreateElement("DATABASE_PASSWORD")
        xmlDBPassword.InnerText = Encryption.Encrypt(Trim(DBPassword))
        xmlSBODatabase.AppendChild(xmlDBPassword)

        Dim xmlDBServerType As System.Xml.XmlNode = xmlDocument.CreateElement("DATABASE_SERVER_TYPE")
        xmlDBServerType.InnerText = Trim(DBServerType)
        xmlSBODatabase.AppendChild(xmlDBServerType)

        Dim xmlSBODBName As System.Xml.XmlNode = xmlDocument.CreateElement("SBO_DATABASE_NAME")
        xmlSBODBName.InnerText = Trim(SBODatabaseName)
        xmlSBODatabase.AppendChild(xmlSBODBName)

        Dim xmlSBOUserName As System.Xml.XmlNode = xmlDocument.CreateElement("SBO_USER_NAME")
        xmlSBOUserName.InnerText = Trim(SBOUserName)
        xmlSBODatabase.AppendChild(xmlSBOUserName)

        Dim xmlSBOPassword As System.Xml.XmlNode = xmlDocument.CreateElement("SBO_PASSWORD")
        xmlSBOPassword.InnerText = Encryption.Encrypt(Trim(SBOPassword))
        xmlSBODatabase.AppendChild(xmlSBOPassword)

        Dim xmlSBOLicenseServer As System.Xml.XmlNode = xmlDocument.CreateElement("SBO_LICENSE_SERVER")
        xmlSBOLicenseServer.InnerText = Trim(SBOLicenseServer)
        xmlSBODatabase.AppendChild(xmlSBOLicenseServer)

        Dim xmlSBOAddOnIdentifier As System.Xml.XmlNode = xmlDocument.CreateElement("SBO_ADDON_IDENTIFIER")
        xmlSBOAddOnIdentifier.InnerText = Trim(SBOAddOnIdentifier)
        xmlSBODatabase.AppendChild(xmlSBOAddOnIdentifier)

        Return xmlSBODatabase
    End Function

    Public Overrides Function Type() As Integer
        Return ConnectionType.SBO
    End Function

    Public Overrides Function ToString() As String
        Dim text As String

        If Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            text = "SBO : " + ServerName() + ":" + Port() + " ; " + DBServerType() + " ; " + SBODatabaseName()
        Else
            text = "SBO : " + ServerName() + " ; " + DBServerType() + " ; " + SBODatabaseName()
        End If

        text = text.ToUpper

        Return text
    End Function

    Public Overrides Function clone() As Connection
        Return New SBOConnection(Name, ServerName, Port, DBUserName, Encryption.Encrypt(DBPassword), DBServerType, SBODatabaseName, SBOUserName, Encryption.Encrypt(SBOPassword), SBOLicenseServer, SBOAddOnIdentifier)
    End Function

#End Region

End Class
