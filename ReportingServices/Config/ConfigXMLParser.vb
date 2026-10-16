Imports System.IO
Public Class ConfigXMLParser

    '<CONFIGURATION>
    '	<APPLICATION>
    '		<PASSWORD/>
    '	</APPLICATION>
    '	<SBO_DATABASE>
    '		<NAME/>
    '		<SERVER_NAME/>
    '		<DATABASE_USER_NAME/>
    '		<DATABASE_PASSWORD/>
    '       <DATABASE_SERVER_TYPE/>
    '		<SBO_DATABASE_NAME/>
    '		<SBO_USER_NAME/>
    '		<SBO_PASSWORD/>
    '		<SBO_LICENSE_SERVER/>
    '		<SBO_ADDON_IDENTIFIER/>
    '	</SBO_DATABASE>
    '	<SQL_DATABASE>                      <HANA_DATABASE>
    '		<NAME/>                             <NAME/>
    '		<SERVER_NAME/>                      <SERVER_NAME/>
    '		<DATABASE_NAME/>                    <PORT/>
    '		<USER_NAME/>                        <DATABASE_NAME/>
    '		<PASSWORD/>                         <USER_NAME/>
    '	</SQL_DATABASE>                         <PASSWORD/>
    '                                       <HANA_DATABASE>
    '</CONFIGURATION>
    Private _Password As String
    Private _PortName As String
    Private _BaudRate As Integer
    Private _DataBits As Integer
    Private _ReadTimeout As Integer
    Private _PaperHeight As Integer
    Private _PaperWidth As Integer
    Private _Separator As String
    Private _PathMDB As String
    Private _Menu As String
    Private _AddonDB As String
    Private _ItemCodeJarak As String
    Private _ItemNameJarak As String
    Private _ItemCodeBerat As String
    Private _ItemNameBerat As String
    Private _ItemCodeMesin As String
    Private _ItemNameMesin As String
    Private _TipeIssueOnFarm As String
    Private _RptName As String
    'Private CONFIGURATION_FILE As String = " C:\Users\Nicholas\Documents\Visual Studio 2010\Projects\ImportDataService\ImportDataService\bin\Debug\Configuration.xml"
    Private CONFIGURATION_FILE As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).ToString.Substring(6) & "\Configuration.xml"

    Public Function isConfigFileExists() As Boolean
        Return System.IO.File.Exists(CONFIGURATION_FILE)
    End Function

    Public Function read() As Boolean

        If Not isConfigFileExists() Then
            Return False
        End If

        Try
            Dim xmlDocument As System.Xml.XmlDocument
            Dim xmlNodeList As System.Xml.XmlNodeList
            Dim xmlNode As System.Xml.XmlNode
            xmlDocument = New System.Xml.XmlDocument()
            xmlDocument.Load(CONFIGURATION_FILE)
            xmlNodeList = xmlDocument.SelectNodes("/ADDON_CONFIGURATION/CONFIGURATION")
            For Each xmlNode In xmlNodeList
                Dim xmlNodeEnumerator As IEnumerator = xmlNode.ChildNodes.GetEnumerator
                While (xmlNodeEnumerator.MoveNext)
                    Dim xmlNodeInner As System.Xml.XmlNode
                    xmlNodeInner = CType(xmlNodeEnumerator.Current, System.Xml.XmlNode)

                    If String.Equals("APPLICATION", xmlNodeInner.Name) Then
                        Dim ApplicationXMLHandler As ApplicationXMLHandler = New ApplicationXMLHandler

                        oApplicationConfig = CType(ApplicationXMLHandler.handle(xmlNodeInner), ApplicationConfig)
                    ElseIf String.Equals("SBO_DATABASE", xmlNodeInner.Name) Then
                        Dim SBOXMLHandler As SBOXMLHandler = New SBOXMLHandler

                        oSBOConnection = CType(SBOXMLHandler.handle(xmlNodeInner), SBOConnection)
                    ElseIf String.Equals("SQL_DATABASE", xmlNodeInner.Name) Then
                        Dim SQLXMLHandler As SQLXMLHandler = New SQLXMLHandler

                        oDBConnection = CType(SQLXMLHandler.handle(xmlNodeInner), SQLConnection)
                    ElseIf String.Equals("HANA_DATABASE", xmlNodeInner.Name) Then
                        Dim HANAXMLHandler As HANAXMLHandler = New HANAXMLHandler

                        oDBConnection = CType(HANAXMLHandler.handle(xmlNodeInner), HANAConnection)
                    End If
                End While
            Next
        Catch ex As Exception
            Dim message As String
            message = "Read Connection's XML : " + Environment.NewLine + ex.ToString()
            ' MessageBox.Show(Me, message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)

            '  System.Windows.Forms.Application.Exit()
        End Try

        Return True
    End Function

    Public Function write() As Boolean
        Try
            Dim xmlDocument As New System.Xml.XmlDocument()

            ' Use the XmlDeclaration class to place the
            ' <?xml version="1.0" encoding="utf-8"?> declaration at the top of our XML file
            Dim xmlDeclaration As System.Xml.XmlDeclaration = xmlDocument.CreateXmlDeclaration("1.0", "UTF-8", Nothing)
            xmlDocument.AppendChild(xmlDeclaration)

            Dim xmlDocRoot As System.Xml.XmlElement = xmlDocument.CreateElement("ADDON_CONFIGURATION")
            xmlDocRoot.SetAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
            xmlDocRoot.SetAttribute("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
            xmlDocument.AppendChild(xmlDocRoot)

            Dim xmlConfiguration As System.Xml.XmlNode = xmlDocument.CreateElement("CONFIGURATION")

            If Not IsNothing(oApplicationConfig) Then
                xmlConfiguration.AppendChild(oApplicationConfig.toXML(xmlDocument))
            End If

            If Not IsNothing(oSBOConnection) Then
                xmlConfiguration.AppendChild(oSBOConnection.toXML(xmlDocument))
            End If

            If Not IsNothing(oDBConnection) Then
                xmlConfiguration.AppendChild(oDBConnection.toXML(xmlDocument))
            End If


            xmlDocRoot.AppendChild(xmlConfiguration)
            xmlDocument.Save(CONFIGURATION_FILE)
            Return True
        Catch ex As Exception
            Dim message As String
            message = ex.Message
            '  MessageBox.Show(Me, message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return False
    End Function




End Class
