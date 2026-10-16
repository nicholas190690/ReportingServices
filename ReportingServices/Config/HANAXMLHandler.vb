Public Class HANAXMLHandler
    Inherits XMLHandler

    Public Overrides Function handle(ByVal xmlNode As System.Xml.XmlNode) As Object
        If Not String.Equals("HANA_DATABASE", xmlNode.Name) Then
            Return Nothing
        End If

        Dim ConName As String = ""
        Dim ServerName As String = ""
        Dim Port As String = ""
        Dim DatabaseName As String = ""
        Dim UserName As String = ""
        Dim Password As String = ""

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
            ElseIf String.Equals("DATABASE_NAME", name) Then
                DatabaseName = value
            ElseIf String.Equals("USER_NAME", name) Then
                UserName = value
            ElseIf String.Equals("PASSWORD", name) Then
                Password = value
            End If
        End While

        Return New HANAConnection(ConName, ServerName, Port, DatabaseName, UserName, Password)
    End Function

End Class
