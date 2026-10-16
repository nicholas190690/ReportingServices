Imports System.Xml

Public Class ApplicationXMLHandler
    Inherits XMLHandler

    Public Overrides Function handle(ByVal xmlNode As XmlNode) As Object
        If Not String.Equals("APPLICATION", xmlNode.Name) Then
            Return Nothing
        End If

        Dim Application As ApplicationConfig
        Application = New ApplicationConfig()

        Dim xmlNodeEnumerator As IEnumerator = xmlNode.ChildNodes.GetEnumerator
        While (xmlNodeEnumerator.MoveNext)
            Dim xmlElement As XmlElement
            xmlElement = CType(xmlNodeEnumerator.Current, XmlElement)

            Dim name As String = xmlElement.Name
            Dim value As String = xmlElement.InnerText

            If String.Equals("PORT_NAME", name) Then
                If Not Application Is Nothing Then
                    Application.PortName = value
                End If
            ElseIf String.Equals("BAUD_RATE", name) Then
                If Not Application Is Nothing Then
                    Application.BaudRate = value
                End If
            ElseIf String.Equals("DATA_BITS", name) Then
                If Not Application Is Nothing Then
                    Application.DataBits = value
                End If
            ElseIf String.Equals("READ_TIMEOUT", name) Then
                If Not Application Is Nothing Then
                    Application.ReadTimeout = value
                End If
            ElseIf String.Equals("SEPARATOR", name) Then
                If Not Application Is Nothing Then
                    Application.Separator = value
                End If
            ElseIf String.Equals("PathMDB", name) Then
                If Not Application Is Nothing Then
                    Application.PathMDB = value
                End If

            ElseIf String.Equals("ItemIssueCode", name) Then
                If Not Application Is Nothing Then
                    Application.ItemIssueCode = value
                End If


            ElseIf String.Equals("ItemIssueName", name) Then
                If Not Application Is Nothing Then
                    Application.ItemIssueName = value
                End If


            ElseIf String.Equals("ItemReceiptCode", name) Then
                If Not Application Is Nothing Then
                    Application.ItemReceiptCode = value
                End If


            ElseIf String.Equals("ItemReceiptName", name) Then
                If Not Application Is Nothing Then
                    Application.ItemReceiptName = value
                End If

            ElseIf String.Equals("TipeIssue", name) Then
                If Not Application Is Nothing Then
                    Application.TipeIssue = value
                End If


            ElseIf String.Equals("TipeReceipt", name) Then
                If Not Application Is Nothing Then
                    Application.TipeReceipt = value
                End If


            ElseIf String.Equals("StartID", name) Then
                If Not Application Is Nothing Then
                    Application.StartID = value
                End If

            ElseIf String.Equals("Timbangan", name) Then
                If Not Application Is Nothing Then
                    Application.Timbangan = value
                End If


            ElseIf String.Equals("PathBag", name) Then
                If Not Application Is Nothing Then
                    Application.PathBag = value
                End If

            ElseIf String.Equals("Menu", name) Then
                If Not Application Is Nothing Then
                    Application.Menu = value
                End If


            ElseIf String.Equals("AddonDB", name) Then
                If Not Application Is Nothing Then
                    Application.AddonDB = value
                End If


            ElseIf String.Equals("ItemCodeJarak", name) Then
                If Not Application Is Nothing Then
                    Application.ItemCodeJarak = value
                End If

            ElseIf String.Equals("ItemNameJarak", name) Then
                If Not Application Is Nothing Then
                    Application.ItemNameJarak = value
                End If

            ElseIf String.Equals("ItemCodeBerat", name) Then
                If Not Application Is Nothing Then
                    Application.ItemCodeBerat = value
                End If

            ElseIf String.Equals("ItemNameBerat", name) Then
                If Not Application Is Nothing Then
                    Application.ItemNameBerat = value
                End If

            ElseIf String.Equals("ItemCodeMesin", name) Then
                If Not Application Is Nothing Then
                    Application.ItemCodeMesin = value
                End If

            ElseIf String.Equals("ItemNameMesin", name) Then
                If Not Application Is Nothing Then
                    Application.ItemNameMesin = value
                End If

            ElseIf String.Equals("TipeIssueOnFarm", name) Then
                If Not Application Is Nothing Then
                    Application.TipeIssueOnFarm = value
                End If

            ElseIf String.Equals("TipeGrpoDocOnFarm", name) Then
                If Not Application Is Nothing Then
                    Application.TipeGrpoDocOnFarm = value
                End If

            ElseIf String.Equals("TipeGrpoPrOnFarm", name) Then
                If Not Application Is Nothing Then
                    Application.TipeGrpoPrOnFarm = value
                End If

            ElseIf String.Equals("RptName", name) Then
                If Not Application Is Nothing Then
                    Application.RptName = value
                End If


            ElseIf String.Equals("WorkFolder", name) Then
                If Not Application Is Nothing Then
                    Application.WorkFolder = value
                End If
            End If






        End While

        Return Application
    End Function

End Class
