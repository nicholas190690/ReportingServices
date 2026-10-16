Imports System.Xml

Public Class ApplicationConfig

    Private _Password As String
    Private _PortName As String
    Private _BaudRate As Integer
    Private _DataBits As Integer
    Private _ReadTimeout As Integer
    Private _Separator As Integer
    Private _PathMDB As String
    Private _ItemIssueCode As String
    Private _ItemIssueName As String
    Private _ItemReceiptCode As String
    Private _ItemReceiptName As String
    Private _TipeIssue As String
    Private _TipeReceipt As String
    Private _StartID As String
    Private _Timbangan As String
    Private _PathBag As String
    Private _Menu As String
    Private _AddonDB As String
    Private _ItemCodeJarak As String
    Private _ItemNameJarak As String
    Private _ItemCodeBerat As String
    Private _ItemNameBerat As String
    Private _ItemCodeMesin As String
    Private _ItemNameMesin As String
    Private _TipeIssueOnFarm As String
    Private _TipeGrpoDocOnFarm As String
    Private _TipeGrpoPrOnFarm As String
    Private _RptName As String
    Private _WorkFolder As String

#Region "Properties"

    Public Property PortName() As String
        Get
            Return _PortName
        End Get
        Set(ByVal value As String)
            _PortName = value
        End Set
    End Property

    Public Property BaudRate() As Integer
        Get
            Return _BaudRate
        End Get
        Set(ByVal value As Integer)
            _BaudRate = value
        End Set
    End Property

    Public Property DataBits() As Integer
        Get
            Return _DataBits
        End Get
        Set(ByVal value As Integer)
            _DataBits = value
        End Set
    End Property

    Public Property ReadTimeout() As Integer
        Get
            Return _ReadTimeout
        End Get
        Set(ByVal value As Integer)
            _ReadTimeout = value
        End Set
    End Property

    Public Property Separator() As Integer
        Get
            Return _Separator
        End Get
        Set(ByVal value As Integer)
            _Separator = value
        End Set
    End Property

    Public Property Password As String
        Get
            Return _Password
        End Get
        Set(ByVal value As String)
            _Password = value
        End Set
    End Property


    Public Property PathMDB As String
        Get
            Return _PathMDB
        End Get
        Set(ByVal value As String)
            _PathMDB = value
        End Set
    End Property


    Public Property ItemIssueCode As String
        Get
            Return _ItemIssueCode
        End Get
        Set(ByVal value As String)
            _ItemIssueCode = value
        End Set
    End Property

    Public Property ItemIssueName As String
        Get
            Return _ItemIssueName
        End Get
        Set(ByVal value As String)
            _ItemIssueName = value
        End Set
    End Property



    Public Property ItemReceiptCode As String
        Get
            Return _ItemReceiptCode
        End Get
        Set(ByVal value As String)
            _ItemReceiptCode = value
        End Set
    End Property

    Public Property ItemReceiptName As String
        Get
            Return _ItemReceiptName
        End Get
        Set(ByVal value As String)
            _ItemReceiptName = value
        End Set
    End Property

    Public Property TipeIssue As String
        Get
            Return _TipeIssue
        End Get
        Set(ByVal value As String)
            _TipeIssue = value
        End Set
    End Property

 Public Property TipeReceipt As String
        Get
            Return _TipeReceipt
        End Get
        Set(ByVal value As String)
            _TipeReceipt = value
        End Set
    End Property


    Public Property StartID As String
        Get
            Return _StartID
        End Get
        Set(ByVal value As String)
            _StartID = value
        End Set
    End Property

    Public Property Timbangan As String
        Get
            Return _Timbangan
        End Get
        Set(ByVal value As String)
            _Timbangan = value
        End Set
    End Property


    Public Property PathBag As String
        Get
            Return _PathBag
        End Get
        Set(ByVal value As String)
            _PathBag = value
        End Set
    End Property


    Public Property Menu As String
        Get
            Return _Menu
        End Get
        Set(ByVal value As String)
            _Menu = value
        End Set
    End Property


    Public Property AddonDB As String
        Get
            Return _AddonDB
        End Get
        Set(ByVal value As String)
            _AddonDB = value
        End Set
    End Property


    Public Property ItemCodeJarak As String
        Get
            Return _ItemCodeJarak
        End Get
        Set(ByVal value As String)
            _ItemCodeJarak = value
        End Set
    End Property

    Public Property ItemNameJarak As String
        Get
            Return _ItemNameJarak
        End Get
        Set(ByVal value As String)
            _ItemNameJarak = value
        End Set
    End Property



    Public Property ItemCodeBerat As String
        Get
            Return _ItemCodeBerat
        End Get
        Set(ByVal value As String)
            _ItemCodeBerat = value
        End Set
    End Property

    Public Property ItemNameBerat As String
        Get
            Return _ItemNameBerat
        End Get
        Set(ByVal value As String)
            _ItemNameBerat = value
        End Set
    End Property

    Public Property ItemCodeMesin As String
        Get
            Return _ItemCodeMesin
        End Get
        Set(ByVal value As String)
            _ItemCodeMesin = value
        End Set
    End Property

    Public Property ItemNameMesin As String
        Get
            Return _ItemNameMesin
        End Get
        Set(ByVal value As String)
            _ItemNameMesin = value
        End Set
    End Property

    Public Property TipeIssueOnFarm As String
        Get
            Return _TipeIssueOnFarm
        End Get
        Set(ByVal value As String)
            _TipeIssueOnFarm = value
        End Set
    End Property

    Public Property TipeGrpoDocOnFarm As String
        Get
            Return _TipeGrpoDocOnFarm
        End Get
        Set(ByVal value As String)
            _TipeGrpoDocOnFarm = value
        End Set
    End Property

    Public Property TipeGrpoPrOnFarm As String
        Get
            Return _TipeGrpoPrOnFarm
        End Get
        Set(ByVal value As String)
            _TipeGrpoPrOnFarm = value
        End Set
    End Property



    Public Property RptName As String
        Get
            Return _RptName
        End Get
        Set(ByVal value As String)
            _RptName = value
        End Set
    End Property

    Public Property WorkFolder As String
        Get
            Return _WorkFolder
        End Get
        Set(ByVal value As String)
            _WorkFolder = value
        End Set
    End Property

#End Region




    Public Function toXML(ByVal xmlDocument As XmlDocument) As XmlNode
        Dim xmlAppConfig As XmlNode = xmlDocument.CreateElement("APPLICATION")

        Dim xmlPassword As XmlNode = xmlDocument.CreateElement("PASSWORD")
        xmlPassword.InnerText = Trim(Password)
        xmlAppConfig.AppendChild(xmlPassword)

        Dim xmlPortName As XmlNode = xmlDocument.CreateElement("PORT_NAME")
        xmlPortName.InnerText = PortName.ToString.Trim
        xmlAppConfig.AppendChild(xmlPortName)

        Dim xmlBaudRate As XmlNode = xmlDocument.CreateElement("BAUD_RATE")
        xmlBaudRate.InnerText = BaudRate.ToString.Trim
        xmlAppConfig.AppendChild(xmlBaudRate)

        Dim xmlDataBits As XmlNode = xmlDocument.CreateElement("DATA_BITS")
        xmlDataBits.InnerText = DataBits.ToString.Trim
        xmlAppConfig.AppendChild(xmlDataBits)

        Dim xmlReadTimeout As XmlNode = xmlDocument.CreateElement("READ_TIMEOUT")
        xmlReadTimeout.InnerText = ReadTimeout.ToString.Trim
        xmlAppConfig.AppendChild(xmlReadTimeout)

        Dim xmlSeparator As XmlNode = xmlDocument.CreateElement("SEPARATOR")
        xmlSeparator.InnerText = Separator.ToString.Trim
        xmlAppConfig.AppendChild(xmlSeparator)

        Dim xmlPathMDB As XmlNode = xmlDocument.CreateElement("PathMDB")
        xmlPathMDB.InnerText = PathMDB.ToString.Trim
        xmlAppConfig.AppendChild(xmlPathMDB)


        Dim xmlItemIssueCode As XmlNode = xmlDocument.CreateElement("ItemIssueCode")
        xmlItemIssueCode.InnerText = ItemIssueCode.ToString.Trim
        xmlAppConfig.AppendChild(xmlItemIssueCode)



        Dim xmlItemIssueName As XmlNode = xmlDocument.CreateElement("ItemIssueName")
        xmlItemIssueName.InnerText = ItemIssueName.ToString.Trim
        xmlAppConfig.AppendChild(xmlItemIssueName)


        Dim xmlItemReceiptCode As XmlNode = xmlDocument.CreateElement("ItemReceiptCode")
        xmlItemReceiptCode.InnerText = ItemReceiptCode.ToString.Trim
        xmlAppConfig.AppendChild(xmlItemReceiptCode)



        Dim xmlItemReceiptName As XmlNode = xmlDocument.CreateElement("ItemReceiptName")
        xmlItemReceiptName.InnerText = ItemReceiptName.ToString.Trim
        xmlAppConfig.AppendChild(xmlItemReceiptName)



        Dim xmlTipeIssue As XmlNode = xmlDocument.CreateElement("TipeIssue")
        xmlTipeIssue.InnerText = TipeIssue.ToString.Trim
        xmlAppConfig.AppendChild(xmlTipeIssue)

        Dim xmlTipeReceipt As XmlNode = xmlDocument.CreateElement("TipeReceipt")
        xmlTipeReceipt.InnerText = TipeReceipt.ToString.Trim
        xmlAppConfig.AppendChild(xmlTipeReceipt)



        Dim xmlStartID As XmlNode = xmlDocument.CreateElement("StartID")
        xmlStartID.InnerText = StartID.ToString.Trim
        xmlAppConfig.AppendChild(xmlStartID)

        Dim xmlTimbangan As XmlNode = xmlDocument.CreateElement("Timbangan")
        xmlTimbangan.InnerText = Timbangan.ToString.Trim
        xmlAppConfig.AppendChild(xmlTimbangan)


        Dim xmlPathBag As XmlNode = xmlDocument.CreateElement("PathBag")
        xmlPathBag.InnerText = PathBag.ToString.Trim
        xmlAppConfig.AppendChild(xmlPathBag)

        Dim xmlMenu As XmlNode = xmlDocument.CreateElement("Menu")
        xmlMenu.InnerText = Menu.ToString.Trim
        xmlAppConfig.AppendChild(xmlMenu)

        Dim xmlAddonDB As XmlNode = xmlDocument.CreateElement("AddonDB")
        xmlAddonDB.InnerText = AddonDB.ToString.Trim
        xmlAppConfig.AppendChild(xmlAddonDB)

        Dim xmlItemCodeJarak As XmlNode = xmlDocument.CreateElement("ItemCodeJarak")
        xmlItemCodeJarak.InnerText = ItemCodeJarak.ToString.Trim
        xmlAppConfig.AppendChild(xmlItemCodeJarak)

        Dim xmlItemNameJarak As XmlNode = xmlDocument.CreateElement("ItemNameJarak")
        xmlItemNameJarak.InnerText = ItemNameJarak.ToString.Trim
        xmlAppConfig.AppendChild(xmlItemNameJarak)

        Dim xmlItemCodeBerat As XmlNode = xmlDocument.CreateElement("ItemCodeBerat")
        xmlItemCodeBerat.InnerText = ItemCodeBerat.ToString.Trim
        xmlAppConfig.AppendChild(xmlItemCodeBerat)


        Dim xmlItemNameBerat As XmlNode = xmlDocument.CreateElement("ItemNameBerat")
        xmlItemNameBerat.InnerText = ItemNameBerat.ToString.Trim
        xmlAppConfig.AppendChild(xmlItemNameBerat)



        Dim xmlItemCodeMesin As XmlNode = xmlDocument.CreateElement("ItemCodeMesin")
        xmlItemCodeMesin.InnerText = ItemCodeMesin.ToString.Trim
        xmlAppConfig.AppendChild(xmlItemCodeMesin)

        Dim xmlItemNameMesin As XmlNode = xmlDocument.CreateElement("ItemNameMesin")
        xmlItemNameMesin.InnerText = ItemNameMesin.ToString.Trim
        xmlAppConfig.AppendChild(xmlItemNameMesin)

        Dim xmlTipeIssueOnFarm As XmlNode = xmlDocument.CreateElement("TipeIssueOnFarm")
        xmlTipeIssueOnFarm.InnerText = TipeIssueOnFarm.ToString.Trim
        xmlAppConfig.AppendChild(xmlTipeIssueOnFarm)

        Dim xmlTipeGrpoDocOnFarm As XmlNode = xmlDocument.CreateElement("TipeGrpoDocOnFarm")
        xmlTipeGrpoDocOnFarm.InnerText = TipeGrpoDocOnFarm.ToString.Trim
        xmlAppConfig.AppendChild(xmlTipeGrpoDocOnFarm)

        Dim xmlTipeGrpoPrOnFarm As XmlNode = xmlDocument.CreateElement("TipeGrpoPrOnFarm")
        xmlTipeGrpoPrOnFarm.InnerText = TipeGrpoPrOnFarm.ToString.Trim
        xmlAppConfig.AppendChild(xmlTipeGrpoPrOnFarm)

        Dim xmlRptName As XmlNode = xmlDocument.CreateElement("RptName")
        xmlRptName.InnerText = RptName.ToString.Trim
        xmlAppConfig.AppendChild(xmlRptName)


        Dim xmlWorkFolder As XmlNode = xmlDocument.CreateElement("WorkFolder")
        xmlWorkFolder.InnerText = WorkFolder.ToString.Trim
        xmlAppConfig.AppendChild(xmlWorkFolder)

        Return xmlAppConfig

    End Function

End Class
