Public Class SBOXMLHandler
    Inherits XMLHandler

    Public Overrides Function handle(ByVal xmlNode As System.Xml.XmlNode) As Object
        If Not String.Equals("SBO_DATABASE", xmlNode.Name) Then
            Return Nothing
        End If

        Dim ConName As String = ""
        Dim ServerName As String = ""
        Dim Port As String = ""
        Dim DBUserName As String = ""
        Dim DBPassword As String = ""
        Dim DBServerType As String = ""
        Dim SBODatabaseName As String = ""
        Dim SBOUserName As String = ""
        Dim SBOPassword As String = ""
        Dim SBOLicenseServer As String = ""
        Dim SBOAddOnIdentifier As String = ""

        Dim xmlNodeEnumerator As IEnumerator = xmlNode.ChildNodes.GetEnumerator
        While (xmlNodeEnumerator.MoveNext)
            Dim xmlElement As System.Xml.XmlElement
            xmlElement = CType(xmlNodeEnumerator.Current, System.Xml.XmlElement)

            Dim name As String = xmlElement.Name
            Dim value As String = xmlElement.InnerText

            If String.Equals("NAME", name) Then
                ConName = value
            ElseIf String.Equals("SERVER_NAME", name) Then
                ServerName = value
            ElseIf String.Equals("PORT", name) Then
                Port = value
            ElseIf String.Equals("DATABASE_USER_NAME", name) Then
                DBUserName = value
            ElseIf String.Equals("DATABASE_PASSWORD", name) Then
                DBPassword = value
            ElseIf String.Equals("DATABASE_SERVER_TYPE", name) Then
                DBServerType = value
            ElseIf String.Equals("SBO_DATABASE_NAME", name) Then
                SBODatabaseName = value
            ElseIf String.Equals("SBO_USER_NAME", name) Then
                SBOUserName = value
            ElseIf String.Equals("SBO_PASSWORD", name) Then
                SBOPassword = value
            ElseIf String.Equals("SBO_LICENSE_SERVER", name) Then
                SBOLicenseServer = value
            ElseIf String.Equals("SBO_ADDON_IDENTIFIER", name) Then
                SBOAddOnIdentifier = value
            End If
        End While

        Return New SBOConnection(ConName, ServerName, Port, DBUserName, DBPassword, DBServerType, SBODatabaseName, SBOUserName, SBOPassword, SBOLicenseServer, SBOAddOnIdentifier)
    End Function

End Class
