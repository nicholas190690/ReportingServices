Public Class Utils

    Public Shared Sub CreateUDF(
            ByVal oSBOCompany As SAPbobsCOM.Company,
            ByVal TableName As String,
            ByVal FieldName As String,
            ByVal FieldDescription As String,
            ByVal Type As SAPbobsCOM.BoFieldTypes)

        CreateUDF(oSBOCompany, TableName, FieldName, FieldDescription, Type, SAPbobsCOM.BoFldSubTypes.st_None)

    End Sub

    Public Shared Sub CreateUDF(
            ByVal oSBOCompany As SAPbobsCOM.Company,
            ByVal TableName As String,
            ByVal FieldName As String,
            ByVal FieldDescription As String,
            ByVal Type As SAPbobsCOM.BoFieldTypes,
            ByVal SubType As SAPbobsCOM.BoFldSubTypes)

        CreateUDF(oSBOCompany, TableName, FieldName, FieldDescription, Type, SubType, 0)

    End Sub

    Public Shared Sub CreateUDF(
            ByVal oSBOCompany As SAPbobsCOM.Company,
            ByVal TableName As String,
            ByVal FieldName As String,
            ByVal FieldDescription As String,
            ByVal Type As SAPbobsCOM.BoFieldTypes,
            ByVal EditSize As Integer)

        CreateUDF(oSBOCompany, TableName, FieldName, FieldDescription, Type, SAPbobsCOM.BoFldSubTypes.st_None, EditSize)

    End Sub

    Public Shared Sub CreateUDF(
            ByVal oSBOCompany As SAPbobsCOM.Company,
            ByVal TableName As String,
            ByVal FieldName As String,
            ByVal FieldDescription As String,
            ByVal Type As SAPbobsCOM.BoFieldTypes,
            ByVal SubType As SAPbobsCOM.BoFldSubTypes,
            ByVal EditSize As Integer)

        CreateUDF(oSBOCompany, TableName, FieldName, FieldDescription, Type, SubType, EditSize, New List(Of String()), "")

    End Sub

    Public Shared Sub CreateUDF(
            ByVal oSBOCompany As SAPbobsCOM.Company,
            ByVal TableName As String,
            ByVal FieldName As String,
            ByVal FieldDescription As String,
            ByVal Type As SAPbobsCOM.BoFieldTypes,
            ByVal SubType As SAPbobsCOM.BoFldSubTypes,
            ByVal EditSize As Integer,
            ByVal ValidValues As List(Of String()),
            ByVal DefaultValue As String)

        CreateUDF(oSBOCompany, TableName, FieldName, FieldDescription, Type, SubType, EditSize, ValidValues, DefaultValue, SAPbobsCOM.BoYesNoEnum.tNO)

    End Sub

    Public Shared Sub CreateUDF(
            ByVal oSBOCompany As SAPbobsCOM.Company,
            ByVal TableName As String,
            ByVal FieldName As String,
            ByVal FieldDescription As String,
            ByVal Type As SAPbobsCOM.BoFieldTypes,
            ByVal SubType As SAPbobsCOM.BoFldSubTypes,
            ByVal EditSize As Integer,
            ByVal ValidValues As List(Of String()),
            ByVal DefaultValue As String,
            ByVal Mandatory As SAPbobsCOM.BoYesNoEnum)

        Dim oUFields As SAPbobsCOM.UserFieldsMD
        Dim oRec As SAPbobsCOM.Recordset
        Dim lRetCode As Integer = 0
        Dim lErrCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim bIsContinue As Boolean = True

        oRec = CType(oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim haa As String = SQLHandler.CheckUDFSQL(TableName, FieldName)
        oRec.DoQuery(SQLHandler.CheckUDFSQL(TableName, FieldName))

        If oRec.RecordCount = 0 Then
            bIsContinue = True
        Else
            bIsContinue = False
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec)
        oRec = Nothing
        GC.Collect()

        If bIsContinue Then
            oUFields = CType(oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields), SAPbobsCOM.UserFieldsMD)
            oUFields.TableName = TableName
            oUFields.Name = FieldName
            oUFields.Description = FieldDescription
            oUFields.Type = Type
            oUFields.SubType = SubType

            If EditSize > 0 Then
                oUFields.EditSize = EditSize
            End If

            If Not IsNothing(ValidValues) Then
                For i As Integer = 0 To ValidValues.Count - 1
                    Dim values As String() = ValidValues.Item(i)

                    oUFields.ValidValues.Value = values(0)
                    oUFields.ValidValues.Description = values(1)
                    oUFields.ValidValues.Add()
                Next
            End If

            oUFields.DefaultValue = DefaultValue
            oUFields.Mandatory = Mandatory

            lRetCode = oUFields.Add()

            If lRetCode <> 0 Then
                oSBOCompany.GetLastError(lErrCode, sErrMsg)
                Throw New Exception(lErrCode.ToString & " : " & sErrMsg)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.Collect()
        End If

    End Sub

    Public Shared Sub CreateUDF(
            ByVal oSBOCompany As SAPbobsCOM.Company,
            ByVal TableName As String,
            ByVal FieldName As String,
            ByVal FieldDescription As String,
            ByVal Type As SAPbobsCOM.BoFieldTypes,
            ByVal SubType As SAPbobsCOM.BoFldSubTypes,
            ByVal EditSize As Integer,
            ByVal LinkedTable As String,
            ByVal DefaultValue As String,
            ByVal Mandatory As SAPbobsCOM.BoYesNoEnum)

        Dim oUFields As SAPbobsCOM.UserFieldsMD
        Dim oRec As SAPbobsCOM.Recordset
        Dim lRetCode As Integer = 0
        Dim lErrCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim bIsContinue As Boolean = True

        oRec = CType(oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        oRec.DoQuery(SQLHandler.CheckUDFSQL(TableName, FieldName))

        If oRec.RecordCount = 0 Then
            bIsContinue = True
        Else
            bIsContinue = False
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec)
        oRec = Nothing
        GC.Collect()

        If bIsContinue Then
            oUFields = CType(oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields), SAPbobsCOM.UserFieldsMD)
            oUFields.TableName = TableName
            oUFields.Name = FieldName
            oUFields.Description = FieldDescription
            oUFields.Type = Type
            oUFields.SubType = SubType

            If EditSize > 0 Then
                oUFields.EditSize = EditSize
            End If

            oUFields.LinkedTable = LinkedTable
            oUFields.DefaultValue = DefaultValue
            oUFields.Mandatory = Mandatory

            lRetCode = oUFields.Add()

            If lRetCode <> 0 Then
                oSBOCompany.GetLastError(lErrCode, sErrMsg)
                Throw New Exception(lErrCode.ToString & " : " & sErrMsg)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.Collect()
        End If

    End Sub

    Public Shared Sub CreateUDT(
            ByVal oSBOCompany As SAPbobsCOM.Company,
            ByVal TableName As String,
            ByVal TableDescription As String,
            ByVal TableType As SAPbobsCOM.BoUTBTableType)

        Dim oUTables As SAPbobsCOM.UserTablesMD
        Dim lRetCode As Integer = 0
        Dim lErrCode As Integer = 0
        Dim sErrMsg As String = ""

        oUTables = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)

        If Not oUTables.GetByKey(TableName) Then
            oUTables.TableName = TableName
            oUTables.TableDescription = TableDescription
            oUTables.TableType = TableType
            lRetCode = oUTables.Add()
        End If

        If lRetCode <> 0 Then
            oSBOCompany.GetLastError(lErrCode, sErrMsg)
            Throw New Exception(lErrCode.ToString & " : " & sErrMsg)
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUTables)
        oUTables = Nothing
        GC.Collect()

    End Sub

    

    Public Shared Sub CreateUDO(
            ByVal oSBOCompany As SAPbobsCOM.Company,
            ByVal ObjectName As String,
            ByVal ObjectType As SAPbobsCOM.BoUDOObjType,
            ByVal TableName As String,
            ByVal CanApprove As SAPbobsCOM.BoYesNoEnum,
            ByVal CanArchive As SAPbobsCOM.BoYesNoEnum,
            ByVal CanCancel As SAPbobsCOM.BoYesNoEnum,
            ByVal CanClose As SAPbobsCOM.BoYesNoEnum,
            ByVal CanCreateDefaultForm As SAPbobsCOM.BoYesNoEnum,
            ByVal CanDelete As SAPbobsCOM.BoYesNoEnum,
            ByVal CanFind As SAPbobsCOM.BoYesNoEnum,
            ByVal CanLog As SAPbobsCOM.BoYesNoEnum,
            ByVal CanYearTransfer As SAPbobsCOM.BoYesNoEnum,
            ByVal CanSeries As SAPbobsCOM.BoYesNoEnum)
        '  ByVal FindColumns As String()) ',
        ' ByVal ChildTables As String())

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Dim oRec As SAPbobsCOM.Recordset
        Dim lRetCode As Integer = 0
        Dim lErrCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim bIsContinue As Boolean = True

        oRec = CType(oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        oRec.DoQuery(SQLHandler.CheckUDOSQL(ObjectName))

        If oRec.RecordCount = 0 Then
            bIsContinue = True
        Else
            bIsContinue = False
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec)
        oRec = Nothing
        GC.Collect()

        If bIsContinue Then
            oUserObjectMD = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

            oUserObjectMD.Code = ObjectName
            oUserObjectMD.Name = ObjectName
            oUserObjectMD.ObjectType = ObjectType
            oUserObjectMD.TableName = TableName

            oUserObjectMD.CanApprove = CanApprove
            oUserObjectMD.CanArchive = CanArchive

            oUserObjectMD.CanCancel = CanCancel
            oUserObjectMD.CanClose = CanClose
            oUserObjectMD.CanCreateDefaultForm = CanCreateDefaultForm
            oUserObjectMD.CanDelete = CanDelete
            oUserObjectMD.CanFind = CanFind
            oUserObjectMD.CanLog = CanLog
            oUserObjectMD.CanYearTransfer = CanYearTransfer
            oUserObjectMD.ManageSeries = CanSeries

            'For i As Integer = 0 To FindColumns.Length - 1
            '    oUserObjectMD.FindColumns.ColumnAlias = FindColumns(i)
            '    oUserObjectMD.FindColumns.Add()
            'Next

            'For i As Integer = 0 To ChildTables.Length - 1
            '    oUserObjectMD.ChildTables.TableName = ChildTables(i)
            '    oUserObjectMD.ChildTables.Add()
            'Next

            lRetCode = oUserObjectMD.Add()

            If lRetCode <> 0 Then
                oSBOCompany.GetLastError(lErrCode, sErrMsg)
                Throw New Exception(lErrCode.ToString & " : " & sErrMsg)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing

            GC.Collect() 'Release the handle to the table
        End If
    End Sub

    Public Shared Sub CreateUDO(
            ByVal oSBOCompany As SAPbobsCOM.Company,
            ByVal ObjectName As String,
            ByVal ObjectType As SAPbobsCOM.BoUDOObjType,
            ByVal TableName As String,
            ByVal CanApprove As SAPbobsCOM.BoYesNoEnum,
            ByVal CanArchive As SAPbobsCOM.BoYesNoEnum,
            ByVal CanCancel As SAPbobsCOM.BoYesNoEnum,
            ByVal CanClose As SAPbobsCOM.BoYesNoEnum,
            ByVal CanCreateDefaultForm As SAPbobsCOM.BoYesNoEnum,
            ByVal CanDelete As SAPbobsCOM.BoYesNoEnum,
            ByVal CanFind As SAPbobsCOM.BoYesNoEnum,
            ByVal CanLog As SAPbobsCOM.BoYesNoEnum,
            ByVal CanYearTransfer As SAPbobsCOM.BoYesNoEnum,
            ByVal CanSeries As SAPbobsCOM.BoYesNoEnum,
            ByVal ChildTables As String())

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Dim oRec As SAPbobsCOM.Recordset
        Dim lRetCode As Integer = 0
        Dim lErrCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim bIsContinue As Boolean = True

        oRec = CType(oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        oRec.DoQuery(SQLHandler.CheckUDOSQL(ObjectName))

        If oRec.RecordCount = 0 Then
            bIsContinue = True
        Else
            bIsContinue = False
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec)
        oRec = Nothing
        GC.Collect()

        If bIsContinue Then
            oUserObjectMD = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

            oUserObjectMD.Code = ObjectName
            oUserObjectMD.Name = ObjectName
            oUserObjectMD.ObjectType = ObjectType
            oUserObjectMD.TableName = TableName

            oUserObjectMD.CanApprove = CanApprove
            oUserObjectMD.CanArchive = CanArchive

            oUserObjectMD.CanCancel = CanCancel
            oUserObjectMD.CanClose = CanClose
            oUserObjectMD.CanCreateDefaultForm = CanCreateDefaultForm
            oUserObjectMD.CanDelete = CanDelete
            oUserObjectMD.CanFind = CanFind
            oUserObjectMD.CanLog = CanLog
            oUserObjectMD.CanYearTransfer = CanYearTransfer
            oUserObjectMD.ManageSeries = CanSeries

            'For i As Integer = 0 To FindColumns.Length - 1
            '    oUserObjectMD.FindColumns.ColumnAlias = FindColumns(i)
            '    oUserObjectMD.FindColumns.Add()
            'Next

            For i As Integer = 0 To ChildTables.Length - 1
                oUserObjectMD.ChildTables.TableName = ChildTables(i)
                oUserObjectMD.ChildTables.Add()
            Next

            lRetCode = oUserObjectMD.Add()

            If lRetCode <> 0 Then
                oSBOCompany.GetLastError(lErrCode, sErrMsg)
                Throw New Exception(lErrCode.ToString & " : " & sErrMsg)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing

            GC.Collect() 'Release the handle to the table
        End If
    End Sub

    Public Shared Sub CreateUDO(
            ByVal oSBOCompany As SAPbobsCOM.Company,
            ByVal ObjectName As String,
            ByVal ObjectType As SAPbobsCOM.BoUDOObjType,
            ByVal TableName As String,
            ByVal CanApprove As SAPbobsCOM.BoYesNoEnum,
            ByVal CanArchive As SAPbobsCOM.BoYesNoEnum,
            ByVal CanCancel As SAPbobsCOM.BoYesNoEnum,
            ByVal CanClose As SAPbobsCOM.BoYesNoEnum,
            ByVal CanCreateDefaultForm As SAPbobsCOM.BoYesNoEnum,
            ByVal CanDelete As SAPbobsCOM.BoYesNoEnum,
            ByVal CanFind As SAPbobsCOM.BoYesNoEnum,
            ByVal CanLog As SAPbobsCOM.BoYesNoEnum,
            ByVal CanYearTransfer As SAPbobsCOM.BoYesNoEnum,
            ByVal FatherMenuID As String,
            ByVal MenuCaption As String,
            ByVal FindColumns As String(),
            ByVal ChildTables As String(),
            ByVal FormColumns As String(,))

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Dim oRec As SAPbobsCOM.Recordset
        Dim lRetCode As Integer = 0
        Dim lErrCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim bIsContinue As Boolean = True

        oRec = CType(oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        oRec.DoQuery(SQLHandler.CheckUDOSQL(ObjectName))

        If oRec.RecordCount = 0 Then
            bIsContinue = True
        Else
            bIsContinue = False
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec)
        oRec = Nothing
        GC.Collect()

        If bIsContinue Then
            oUserObjectMD = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

            oUserObjectMD.Code = ObjectName
            oUserObjectMD.Name = ObjectName
            oUserObjectMD.ObjectType = ObjectType
            oUserObjectMD.TableName = TableName

            oUserObjectMD.CanApprove = CanApprove
            oUserObjectMD.CanArchive = CanArchive
            oUserObjectMD.CanCancel = CanCancel
            oUserObjectMD.CanClose = CanClose
            oUserObjectMD.CanCreateDefaultForm = CanCreateDefaultForm
            oUserObjectMD.CanDelete = CanDelete
            oUserObjectMD.CanFind = CanFind
            oUserObjectMD.CanLog = CanLog
            oUserObjectMD.CanYearTransfer = CanYearTransfer

            oUserObjectMD.EnableEnhancedForm = CanCreateDefaultForm
            oUserObjectMD.MenuItem = CanCreateDefaultForm
            oUserObjectMD.FatherMenuID = FatherMenuID
            oUserObjectMD.MenuCaption = MenuCaption
            oUserObjectMD.Position = 0
            oUserObjectMD.MenuUID = ObjectName

            For i As Integer = 0 To FindColumns.Length - 1
                oUserObjectMD.FindColumns.ColumnAlias = FindColumns(i)
                oUserObjectMD.FindColumns.Add()
            Next

            For i As Integer = 0 To FormColumns.GetLength(0) - 1
                oUserObjectMD.FormColumns.FormColumnAlias = FormColumns(i, 0)
                oUserObjectMD.FormColumns.FormColumnDescription = FormColumns(i, 1)
                oUserObjectMD.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.FormColumns.Add()
            Next

            For i As Integer = 0 To ChildTables.Length - 1
                oUserObjectMD.ChildTables.TableName = ChildTables(i)
                oUserObjectMD.ChildTables.Add()
            Next

            lRetCode = oUserObjectMD.Add()

            If lRetCode <> 0 Then
                oSBOCompany.GetLastError(lErrCode, sErrMsg)
                Throw New Exception(lErrCode.ToString & " : " & sErrMsg)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing

            GC.Collect() 'Release the handle to the table
        End If
    End Sub

    Public Shared Sub CreateQueryCategory(
            ByVal oSBOCompany As SAPbobsCOM.Company,
            ByVal CategoryName As String)

        CreateQueryCategory(oSBOCompany, CategoryName, "YYYYYYYYYYYYYYY")

    End Sub

    Public Shared Sub CreateQueryCategory(
            ByVal oSBOCompany As SAPbobsCOM.Company,
            ByVal CategoryName As String,
            ByVal Permissions As String)

        Dim oQueryCategory As SAPbobsCOM.QueryCategories
        Dim oRec As SAPbobsCOM.Recordset
        Dim lRetCode As Integer = 0
        Dim lErrCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim bIsContinue As Boolean = True

        oRec = CType(oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        oRec.DoQuery(SQLHandler.CheckQueryCategorySQL(CategoryName))

        If oRec.RecordCount = 0 Then
            bIsContinue = True
        Else
            bIsContinue = False
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec)
        oRec = Nothing
        GC.Collect()

        If bIsContinue Then
            oQueryCategory = CType(oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQueryCategories), SAPbobsCOM.QueryCategories)
            oQueryCategory.Name = CategoryName
            oQueryCategory.Permissions = Permissions

            lRetCode = oQueryCategory.Add()
            If lRetCode <> 0 Then
                oSBOCompany.GetLastError(lErrCode, sErrMsg)
                Throw New Exception(lErrCode.ToString & " : " & sErrMsg)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oQueryCategory)
            oQueryCategory = Nothing
            GC.Collect()
        End If

    End Sub

    Public Shared Sub CreateQuery(
            ByVal oSBOCompany As SAPbobsCOM.Company,
            ByVal CategoryName As String,
            ByVal QueryName As String,
            ByVal SQL As String)

        Dim oUserQuery As SAPbobsCOM.UserQueries
        Dim oRec As SAPbobsCOM.Recordset
        Dim lRetCode As Integer = 0
        Dim lErrCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim bIsContinue As Boolean = True
        Dim CategoryId As Integer = 0

        oRec = CType(oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        oRec.DoQuery(SQLHandler.CheckQueryCategorySQL(CategoryName))

        If oRec.RecordCount <> 0 Then
            CategoryId = CType(oRec.Fields.Item("CategoryId").Value, Integer)
        Else
            Throw New Exception("Query Category (" + CategoryName + ") Not Exists!!!")
        End If

        oRec.DoQuery(SQLHandler.CheckQuerySQL(CategoryName, QueryName))

        If oRec.RecordCount = 0 Then
            bIsContinue = True
        Else
            bIsContinue = False
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec)
        oRec = Nothing
        GC.Collect()

        If bIsContinue Then
            oUserQuery = CType(oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries), SAPbobsCOM.UserQueries)
            oUserQuery.Query = SQL
            oUserQuery.QueryCategory = CategoryId
            oUserQuery.QueryDescription = QueryName

            lRetCode = oUserQuery.Add()
            If lRetCode <> 0 Then
                oSBOCompany.GetLastError(lErrCode, sErrMsg)
                Throw New Exception(lErrCode.ToString & " : " & sErrMsg)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserQuery)
            oUserQuery = Nothing
            GC.Collect()
        End If

    End Sub

    Public Shared Sub CreateFMS(
            ByVal oSBOCompany As SAPbobsCOM.Company,
            ByVal CategoryName As String,
            ByVal QueryName As String,
            ByVal FormID As String,
            ByVal ItemID As String)

        Dim oFormattedSearch As SAPbobsCOM.FormattedSearches
        Dim oRec As SAPbobsCOM.Recordset
        Dim lRetCode As Integer = 0
        Dim lErrCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim bIsContinue As Boolean = True
        Dim QueryId As Integer
        Dim CategoryId As Integer

        oRec = CType(oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        oRec.DoQuery(SQLHandler.CheckQueryCategorySQL(CategoryName))
        If oRec.RecordCount <> 0 Then
            CategoryId = CType(oRec.Fields.Item("CategoryId").Value, Integer)
        Else
            Throw New Exception("Query Category (" + CategoryName + ") Not Exists!!!")
        End If

        oRec.DoQuery(SQLHandler.CheckQuerySQL(CategoryName, QueryName))
        If oRec.RecordCount <> 0 Then
            QueryId = CType(oRec.Fields.Item("IntrnalKey").Value, Integer)
        Else
            Throw New Exception("Query (" + QueryName + ") Not Exists!!!")
        End If

        oRec.DoQuery(SQLHandler.CheckFMSSQL(FormID, ItemID))

        If oRec.RecordCount = 0 Then
            bIsContinue = True
        Else
            bIsContinue = False
        End If

        Dim IndexID As String
        IndexID = oRec.Fields.Item("IndexID").Value.ToString.Trim

        Dim resultQueryID As String
        resultQueryID = oRec.Fields.Item("QueryId").Value.ToString.Trim

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec)
        oRec = Nothing
        GC.Collect()

        If bIsContinue Then
            oFormattedSearch = CType(oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches), SAPbobsCOM.FormattedSearches)
            oFormattedSearch.FormID = FormID
            oFormattedSearch.ItemID = ItemID
            oFormattedSearch.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery
            oFormattedSearch.QueryID = QueryId

            lRetCode = oFormattedSearch.Add()

            If lRetCode <> 0 Then
                oSBOCompany.GetLastError(lErrCode, sErrMsg)
                Throw New Exception(lErrCode.ToString & " : " & sErrMsg)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oFormattedSearch)
            oFormattedSearch = Nothing
            GC.Collect()
        Else
            If QueryId <> resultQueryID Then
                oFormattedSearch = CType(oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches), SAPbobsCOM.FormattedSearches)
                If oFormattedSearch.GetByKey(IndexID) Then
                    oFormattedSearch.FormID = FormID
                    oFormattedSearch.ItemID = ItemID
                    oFormattedSearch.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery
                    oFormattedSearch.QueryID = QueryId

                    lRetCode = oFormattedSearch.Update()

                    If lRetCode <> 0 Then
                        oSBOCompany.GetLastError(lErrCode, sErrMsg)
                        Throw New Exception(lErrCode.ToString & " : " & sErrMsg)
                    End If

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oFormattedSearch)
                    oFormattedSearch = Nothing
                    GC.Collect()
                End If
            End If
        End If

    End Sub

    ' 20140225
    Public Shared Function ConvertToDateTime(ByVal SBODate As String) As DateTime
        Return New DateTime(SBODate.Substring(0, 4), SBODate.Substring(4, 2), SBODate.Substring(6, 2))
    End Function

    Public Shared Function WindowsToSBONumber(ByVal value As Double) As String
        Return value.ToString.Replace(WinDecSep, SBODecSep)
    End Function

    Public Shared Function SBOToWindowsNumberWithCurrency(ByVal value As String) As Double
        Return SBOToWindowsNumberWithoutCurrency(value.Substring(4).ToString)
    End Function

    Public Shared Function SBOToWindowsNumberWithoutCurrency(ByVal value As String) As Double
        Return value.Replace(SBOThousSep, "").Replace(SBODecSep, WinDecSep)
    End Function

    Public Shared Function FormattedStringAmount(ByVal value As Double) As String
        Dim valueS As String = value.ToString("G17")
        Dim DecSepIndex As Integer = valueS.Trim.IndexOf(WinDecSep)

        Dim a As String = valueS
        Dim b As String = ""

        If DecSepIndex >= 0 Then
            a = valueS.Trim.Substring(0, valueS.Trim.IndexOf(WinDecSep))
            b = valueS.Trim.Substring(valueS.Trim.IndexOf(WinDecSep) + 1)
        End If

        Dim c As Integer = Math.Floor(a.Length / 3)

        Dim d As String = StrReverse(a)

        Dim e As New List(Of String)

        Dim ctr As Integer = 0
        For i As Integer = 0 To c - 1
            Dim f As String = d.Substring(i + ctr, 3)

            ctr = ctr + 2

            e.Add(StrReverse(f))
        Next

        Dim g As String = ""
        Dim h As String = ""

        For i As Integer = e.Count - 1 To 0 Step -1
            g = g + e.Item(i).ToString
            h = h + e.Item(i).ToString + SBOThousSep
        Next

        Dim result As String = a.Substring(0, a.Length - g.Length) + SBOThousSep + h
        result = result.Substring(0, result.Length - 1)

        If result.StartsWith(SBOThousSep) Then
            result = result.Substring(1)
        End If

        If b <> "" Then
            result = result + SBODecSep + b
        End If

        Return result
    End Function


    Public Shared Sub releaseobject(ByVal obj As Object)
        If obj Is Nothing = False Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
            GC.Collect()
        End If
    End Sub

End Class
