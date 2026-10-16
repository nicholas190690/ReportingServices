Public Class SQLVersion
    Implements Queries

    Private DatabaseName As String

    Public Sub New(ByVal DatabaseName As String)
        Me.DatabaseName = DatabaseName
    End Sub

#Region " STANDARD "

    Public Function SeparatorSQL() As String Implements Queries.SeparatorSQL
        Return "SELECT DecSep, ThousSep FROM OADM"
    End Function

    Public Function CheckUDFSQL(ByVal tableName As String, ByVal fieldName As String) As String Implements Queries.CheckUDFSQL
        Return "SELECT TableID FROM CUFD WHERE TableID = '" + tableName.Trim + "' AND AliasID = '" + fieldName.Trim + "'"
    End Function

    Public Function CheckUDOSQL(ByVal tableName As String) As String Implements Queries.CheckUDOSQL
        Return "SELECT Code FROM OUDO WHERE Code = '" + tableName.Trim + "'"
    End Function

    Public Function CheckQueryCategorySQL(ByVal categoryName As String) As String Implements Queries.CheckQueryCategorySQL
        Return "SELECT CategoryId FROM OQCN WHERE CatName = '" + categoryName + "'"
    End Function

    Public Function CheckQuerySQL(ByVal categoryName As String, ByVal queryName As String) As String Implements Queries.CheckQuerySQL
        Return "SELECT IntrnalKey FROM OUQR WHERE QName = '" + queryName + "' AND QCategory IN (" + CheckQueryCategorySQL(categoryName) + ")"
    End Function

    Public Function CheckFMSSQL(ByVal formID As String, ByVal itemID As String) As String Implements Queries.CheckFMSSQL
        Return "SELECT IndexID, QueryId FROM CSHS WHERE FormID = '" + formID + "' AND ItemID = '" + itemID + "'"
    End Function

    Public Function CheckFunctionExistsSQL(ByVal DBName As String, ByVal FunctionName As String) As String Implements Queries.CheckFunctionExistsSQL
        Return "SELECT TOP 1 1 FROM INFORMATION_SCHEMA.ROUTINES WHERE SPECIFIC_CATALOG = '" + DatabaseName + "' AND SPECIFIC_SCHEMA = 'dbo' AND SPECIFIC_NAME = '" + FunctionName + "' AND ROUTINE_TYPE = 'FUNCTION'"
    End Function

    Public Function CheckSPExistsSQL(ByVal DBName As String, ByVal SPName As String) As String Implements Queries.CheckSPExistsSQL
        Return "SELECT TOP 1 1 FROM INFORMATION_SCHEMA.ROUTINES WHERE SPECIFIC_CATALOG = '" + DatabaseName + "' AND SPECIFIC_SCHEMA = 'dbo' AND SPECIFIC_NAME = '" + SPName + "' AND ROUTINE_TYPE = 'PROCEDURE'"
    End Function

    Public Function GetSeries(ByVal SeriesName As String, ByVal ObjectCode As String) As String Implements Queries.GetSeries
        Return "SELECT  Series  as ""Series"" FROM NNM1 WHERE SeriesName = '" + SeriesName + "' AND ObjectCode = '" + ObjectCode + "'"
    End Function


    Public Function GetRawSugarConfirmList(ByVal mesin As String, ByVal date1 As String, ByVal shift As String) As String Implements Queries.GetRawSugarConfirmList
        If mesin = "All" Then
            mesin = "T0.""U_SOL_MESIN"""
        Else
            mesin = "'" & mesin & "'"
        End If

        Return "SELECT T0.""DocEntry"" AS ""DocEntry"" , T0.""U_SOL_SHIFT"" as ""Shift"" ,T1.""Name"" as ""Mesin Bagging"",RIGHT('00' ||T0.""U_SOL_TIME"",2) || ':00'  as ""Waktu"" , T0.""U_SOL_SUM"" AS ""Jumlah Karung"" , T0.""U_SOL_SUM"" AS ""Jumlah Karung Confirmed"",'0' AS ""Jumlah Karung Cleaning"",'' as ""Comments"", IFNULL(T0.""U_SOL_SUMOVER"",0) AS ""Jumlah Karung Over"", IFNULL(T0.""U_SOL_SUMUNDER"",0) AS ""Jumlah Karung Under"", ""U_SOL_MAXCOUNT"" AS ""Counter"" FROM " + oSBOConnection.DatabaseName.ToString + ".""@SOL_BAGGING"" T0 LEFT JOIN " + oSBOConnection.DatabaseName.ToString + ".""@SOL_MESIN"" T1 ON T0.""U_SOL_MESIN"" = T1.""Code"" AND T1.""U_SOL_TYPE"" = '2'  WHERE T0.""U_SOL_STAT"" = 'Open' and T0.""U_SOL_SHIFT"" = '" & shift & "' AND T0.""U_SOL_DATE"" = '" & date1 & "' AND T0.""U_SOL_MESIN"" = " & mesin & " ORDER BY T1.""Name"" asc, t0.""U_SOL_TIME"" ASC"
    End Function


    Public Function GetShiftFilter(ByVal Type As String) As String Implements Queries.GetShiftFilter
        Return "SELECT ""U_SOL_SHIFT"" AS ""Shift"",""U_SOL_START"" AS ""Start"" ,""U_SOL_END"" as ""End"", ""U_SOL_TYPE"" as ""Type"" FROM " & oSBOConnection.DatabaseName.ToString & ".""@SOL_SHIFT"" WHERE ""U_SOL_TYPE"" = '" & Type.ToString & "'"
    End Function
    Public Function GetTransaksi(ByVal TipeJual As String, ByVal Date1 As String) As String Implements Queries.GetTransaksi

        If TipeJual = "1" Then
            Return "SELECT " _
            & "'1' AS ""SPPB"" " _
            & ",'2' AS ""SPT"" " _
            & ",'3' AS ""PO"" " _
            & ",'4' AS ""SO"" " _
            & ",'v001' AS ""CardCode"" " _
            & ",'NicholAS' AS ""CardName"" " _
            & ",'kapalxxx' AS ""Kapal"" " _
            & ",'a1021' AS ""NoSJ"" " _
            & ",'1' AS ""Kendaraan"" " _
            & ",'B 123 ASD' AS ""Nopol"" " _
            & ",'JNE' AS ""Transporter"" " _
            & ",'Fiqr' AS ""Supir"" " _
            & ",'R1' AS ""ItemCode"" " _
            & ",'1' AS ""NoSystem"" " _
            & ",'1800001' AS ""NoSpm"" " _
            & ",0 AS ""QtySJ"" " _
            & ", 10000 AS ""QtySPM"" UNION ALL " _
            & "SELECT " _
            & "'1' AS ""SPPB"" " _
            & ",'2' AS ""SPT"" " _
            & ",'3' AS ""PO"" " _
            & ",'4' AS ""SO"" " _
            & ",'v001' AS ""CardCode"" " _
            & ",'NicholAS' AS ""CardName"" " _
            & ",'kapalxxx' AS ""Kapal"" " _
            & ",'a1021' AS ""NoSJ"" " _
            & ",'2' AS ""Kendaraan"" " _
            & ",'B 123 ASD' AS ""Nopol"" " _
            & ",'JNE' AS ""Transporter"" " _
            & ",'Fiqr' AS ""Supir"" " _
            & ",'R1' AS ""ItemCode"" " _
            & ",'1' AS ""NoSystem"" " _
            & ",'1800001' AS ""NoSpm"" " _
            & ",0 AS ""QtySJ"" " _
            & ", 10000 AS ""QtySPM"""
        Else
            Return "SELECT 1"
        End If



    End Function

    Public Function GetListSPMSPB(ByVal DateFrom As String, ByVal DateTo As String, ByVal Status As String, ByVal flag As String) As String Implements Queries.GetListSPMSPB
        Return "SELECT T0.""DocNum"" as ""Nomor SPPB"" " _
            & ", T0.""U_SOL_SPT"" as ""Nomor SPT"" " _
            & ", T0.""U_SOL_SO"" as ""Nomor SO"" " _
            & ", T0.""U_SOL_CARDNAME"" as ""Customer"" " _
            & ", T0.""U_SOL_ITEMCODE"" as ""Item"" " _
            & ", T0.""U_SOL_QTY_SPM""  as ""Quantity"" " _
            & ", T0.""U_SOL_SUPIR""  as ""Supir"" " _
            & ", T0.""U_SOL_TRANSPORTIR""  as ""Transportir"" " _
            & ", T0.""U_SOL_NOPOL""  as ""Plat Nomor"" " _
            & ", T0.""U_SOL_TIPE""  as ""Status"" " _
            & "FROM ""@SOL_WB"" T0 WHERE T0.""CreateDate"" >= '" + DateFrom + "' " _
            & "AND T0.""CreateDate"" <= '" + DateTo + "' " _
            & "AND T0.""U_SOL_STATUS"" in (" + Status + ") " _
            & "ORDER BY T0.""DocEntry"" asc "
    End Function

    Public Function GetCurrent() As String Implements Queries.GetCurrent
        Return "SELECT TO_NVARCHAR(CURRENT_DATE,'YYYYMMDD') AS ""CDate"",TO_NVARCHAR(CURRENT_TIME,'HH24MI') as ""Ctime"" FROM DUMMY"
    End Function

    Public Function GetToleransi() As String Implements Queries.GetToleransi
        Return "SELECT TOP 1 ""U_SOL_PERCENTAGE"" FROM ""@SOL_TOLERANSI"" "
    End Function

    Public Function GetForm(ByVal Docentry As Integer, ByVal type As String) As String Implements Queries.GetForm
        Return "CALL " + oSBOConnection.DatabaseName.ToString + ".SOL_FORM_WB(" + Docentry + ") "
    End Function

    Public Function GetFormTicket(ByVal Docentry As Integer, ByVal type As String) As String Implements Queries.GetFormTicket
        Return "CALL " + oSBOConnection.DatabaseName.ToString + ".SOL_FORM_WB(" + Docentry + ") "
    End Function

    Public Function GetLastScale(ByVal Timbangan As String, ByVal shift As String, ByVal tanggal As String) As String Implements Queries.GetLastScale
        Return "SELECT IFNULL(MAX(T0.""U_SOL_END""),0) as ""MAXID"" FROM ""@SOL_RAW"" T0 where T0.""U_SOL_MESIN"" = '" + Timbangan + "' "
    End Function

    Public Function GetBagging(ByVal Mesin As String, ByVal shift As String, ByRef date1 As String) As String Implements Queries.GetBagging
        Return "SELECT T0.""DocEntry"" as ""DocEntry"" FROM ""@SOL_BAGGING_H"" T0 where T0.""U_SOL_SHIFT"" = '" & shift & "' and T0.""U_SOL_MESIN"" = '" & Mesin & "'"

    End Function

    Public Function GetRawSugarList(ByVal stat As String, ByVal date1 As String, ByVal date2 As String) As String Implements Queries.GetRawSugarList
        If stat = "All" Then
            stat = "T0.""U_SOL_STAT"""
        Else
            stat = "'" & stat & "'"
        End If
        Return "Select T0.""DocNum"" as ""NoTimbang"",ifnull(t1.""DocEntry"",0) as ""NoDraft"",ifnull(T2.""DocEntry"",0) as ""NoIssue"",t0.""U_SOL_QTY1"" AS ""Berat Timbangan"" , t0.""U_SOL_QTY2"" as ""Berat Konfirmasi"", t0.""U_SOL_DATE"" AS ""Tanggal"", 0 as ""Jam"",T0.""U_SOL_SHIFT"" AS ""Shift"", t0.""U_SOL_MESIN"" as ""Nomor Timbangan"" , t0.""U_SOL_OPERATOR"" AS ""Operator"" FROM " & _
                """@SOL_RAW"" T0 LEFT JOIN ODRF T1 ON T0.""DocEntry"" = T1.""U_SOL_RAW_NUMBER"" LEFT JOIN OIGE T2 ON T0.""DocEntry"" = T2.""U_SOL_RAW_NUMBER""  WHERE T0.""U_SOL_STAT"" = " & stat & " and T0.""U_SOL_DATE"" >= '" & date1 & "' AND T0.""U_SOL_DATE"" <= '" & date2 & "'"
    End Function

    Public Function GetShift() As String Implements Queries.GetShift
        Return "SELECT ""U_SOL_SHIFT"" AS ""Shift"",""U_SOL_START"" AS ""START"" ,""U_SOL_END"" as ""End"", ""U_SOL_TYPE"" as""Type"" FROM ""SOL_SHIFT"" "
    End Function

    Public Function GetSumRawSugar(ByVal mesin As String, ByVal date1 As String, ByVal shift As String) As String Implements Queries.GetBaggingConfirmList
        Return "SELECT SUM(""U_SOL_QTY2"") AS ""TOTALWEIGHT"" FROM ""@SOL_RAW"" WHERE ""U_SOL_MESIN"" = '" & mesin & "' AND ""U_SOL_SHIFT"" = '" & shift & "' AND ""U_SOL_DATE"" = '" & date1 & "'"
    End Function
    Public Function GetRawSugarDetail(ByVal shift As String, ByVal date1 As String, ByVal docentry As String) As String Implements Queries.GetRawSugarDetail
        Dim query As String = ""
        If docentry = "" Then
            query = query + "SELECT"
            query = query + " T0.""U_SOL_TIME"" AS ""Jam"""
            query = query + " ,T1.""Name"" as ""Timbangan"""
            query = query + " ,t0.""U_SOL_QTY1"" AS ""Qty1"""
            query = query + " ,t0.""U_SOL_QTY2"" AS ""Qty2"""
            query = query + " ,T2.""Name"" as ""Operator"""
            query = query + " FROM"
            query = query + " ""@SOL_RAW"" T0"
            query = query + " LEFT JOIN ""@SOL_MESIN"" T1 ON T0.""U_SOL_MESIN"" = T1.""Code"""
            query = query + " LEFT JOIN ""@SOL_OPERATOR"" T2 ON T0.""U_SOL_OPERATOR"" = T2.""Code"""
            query = query + " WHERE"
            query = query + " T0.""U_SOL_SHIFT"" = '" & shift & "' and T0.""U_SOL_DATE"" = '" & date1 & "' "
        Else
            query = query + "SELECT"
            query = query + " T1.""U_SOL_TIME"" AS ""Jam"""
            query = query + " ,T2.""Name"" as ""Timbangan"""
            query = query + " ,t1.""U_SOL_QTY1"" AS ""Qty1"""
            query = query + " ,t1.""U_SOL_QTY2"" AS ""Qty2"""
            query = query + " ,T3.""Name"" as ""Operator"""
            query = query + " FROM"
            query = query + " ""@SOL_RAW_D"" T0"
            query = query + " INNER JOIN ""@SOL_RAW"" T1 ON T0.""U_SOL_BASEENTRY"" = T1.""DocEntry"""
            query = query + " LEFT JOIN ""@SOL_MESIN"" T2 ON T1.""U_SOL_MESIN"" = T2.""Code"""
            query = query + " LEFT JOIN ""@SOL_OPERATOR"" T3 ON T1.""U_SOL_OPERATOR"" = T3.""Code"""
            query = query + " WHERE"
            query = query + " T0.""DocEntry"" = " & docentry
        End If
        Return query
    End Function


    Public Function GetMesin(ByVal mesin As String) As String Implements Queries.GetMesin
        Return "SELECT * FROM " & oSBOConnection.DatabaseName.ToString & ".""@SOL_MESIN"" "
    End Function


    Public Function GetBaggingConfirmList(ByVal mesin As String, ByVal date1 As String, ByVal shift As String) As String Implements Queries.GetSumRawSugar
        If mesin = "All" Then
            mesin = "T0.""Mesin"""
        Else
            mesin = "'" & mesin & "'"
        End If

        Return "SELECT T0.""DocEntry"" AS ""DocEntry"" , T0.""Shift"" as ""Shift"" ,T1.""Name"",T1.""U_SOL_TIME"" as ""Waktu"" , T0.""U_SOL_SUM"" AS ""Jumlah Karung"" , T0.""U_SOL_SUM"" AS ""Jumlah Karung Confirmed"" FROM ""@SOL_BAGGING"" T0 LEFT JOIN ""@SOL_MESIN"" T1 ON T0.""U_SOL_MESIN"" = T1.""Code"" WHERE T0.""U_SOL_SHIFT"" = '" & shift & "' AND T0.""U_SOL_DATE"" = '" & date1 & "' AND T0.""U_SOL_MESIN"" = " & mesin
    End Function

    Public Function GetBaggingConfirmRecap(ByVal mesin As String, ByVal date1 As String, ByVal shift As String) As String Implements Queries.GetBaggingConfirmRecap
        If mesin = "All" Then
            mesin = "T0.""Mesin"""
        Else
            mesin = "'" & mesin & "'"
        End If

        Return "SELECT T1.""Name"",SUM(T0.""U_SOL_SUM"") AS ""Jumlah Karung""  FROM ""@SOL_BAGGING"" T0 LEFT JOIN ""@SOL_MESIN"" T1 ON T0.""U_SOL_MESIN"" = T1.""Code"" WHERE T0.""U_SOL_SHIFT"" = '" & shift & "' AND T0.""U_SOL_DATE"" = '" & date1 & "' AND T0.""U_SOL_MESIN"" = " & mesin & " GROUP BY  T1.""Name"" ORDER BY T1.""Name"" ASC"
    End Function

    Public Function GetAllBagging(ByVal stat As String, ByVal date1 As String, ByVal date2 As String) As String Implements Queries.GetAllBagging
        Return "CALL SOL_GET_LISTBAGGINGHEADER('" & stat & "','" & date1 & "','" & date2 & "')"
    End Function



    'Public Function GetBaggingStatus(ByVal entry As String) As String Implements Queries.GetBaggingStatus
    '    Return "SELECT 1 FROM ""@SOL_BAGGING_H"" where ""DocEntry"" = " & entry & " and ""U_SOL_STAT"" = 'Confirm'"
    'End Function



    Public Function GetBaggingList(ByVal shift As String, ByVal date1 As String, ByVal docentry As String) As String Implements Queries.GetBaggingList
        Dim query As String = ""
        If docentry = "" Then
            query = query + "SELECT"
            query = query + " T0.""U_SOL_TIME"" AS ""Jam"""
            query = query + " ,T1.""Name"" as ""Mesin Bagging"""
            query = query + " ,ifnull(T0.""U_SOL_SUM"",0) as ""Jumlah Karung"""
            query = query + " ,ifnull(T0.""U_SOL_SUM2"",0) as ""Jumlah Karung Confirm"""
            query = query + " ,ifnull(T0.""U_SOL_WEIGHT"",0) as ""Weight"""
            query = query + " ,T2.""Name"" as ""Operator"""
            query = query + " ,T0.""DocEntry"" as ""DocEntry"""
            query = query + " ,T1.""U_SOL_PREFIX"" as ""Prefix"""
            query = query + " ,T0.""U_SOL_RECEIPT"" as ""TypeReceipt"""
            query = query + " ,T0.""U_SOL_ITEMCODE"" as ""ItemCode"""
            query = query + " ,T0.""U_SOL_ITEMNAME"" as ""ItemName"""
            query = query + " FROM"
            query = query + " ""@SOL_BAGGING"" T0"
            query = query + " LEFT JOIN ""@SOL_MESIN"" T1 ON T0.""U_SOL_MESIN"" = T1.""Code"""
            query = query + " LEFT JOIN ""@SOL_OPERATOR"" T2 ON T0.""U_SOL_OPERATOR"" = T2.""Code"""
            query = query + " WHERE"
            query = query + " T0.""U_SOL_SHIFT"" = '" & shift & "' and T0.""U_SOL_DATE"" = '" & date1 & "' "
        Else
            query = query + "SELECT"
            query = query + " T1.""U_SOL_TIME"" AS ""Jam"""
            query = query + " ,T2.""Name"" as ""Mesin Bagging"""
            query = query + " ,ifnull(T1.""U_SOL_SUM"",0) as ""Jumlah Karung"""
            query = query + " ,ifnull(T1.""U_SOL_SUM2"",0) as ""Jumlah Karung Confirm"""
            query = query + " ,ifnull(T1.""U_SOL_WEIGHT"",0) as ""Weight"""
            query = query + " ,T3.""Name"" as ""Operator"""
            query = query + " ,T1.""DocEntry"" as ""DocEntry"""
            query = query + " ,T2.""U_SOL_PREFIX"" as ""Prefix"""
            query = query + " ,T1.""U_SOL_RECEIPT"" as ""TypeReceipt"""
            query = query + " ,T1.""U_SOL_ITEMCODE"" as ""ItemCode"""
            query = query + " ,T1.""U_SOL_ITEMNAME"" as ""ItemName"""
            query = query + " FROM"
            query = query + " ""@SOL_BAGGING_D"" T0"
            query = query + " INNER JOIN ""@SOL_BAGGING"" T1 ON T0.""U_SOL_BASEENTRY"" = T1.""DocEntry"""
            query = query + " LEFT JOIN ""@SOL_MESIN"" T2 ON T1.""U_SOL_MESIN"" = T2.""Code"""
            query = query + " LEFT JOIN ""@SOL_OPERATOR"" T3 ON T1.""U_SOL_OPERATOR"" = T3.""Code"""
            query = query + " WHERE"
            query = query + " T0.""DocEntry"" = " & docentry
        End If
        Return query

    End Function

    Public Function GetWB(ByVal date1 As String, ByVal date2 As String, ByVal status As String, ByVal type As String, ByVal ColumnFilter As String, ByVal ValueFilter As String) As String Implements Queries.GetWB
        Dim test As String = "CALL " & oApplicationConfig.AddonDB.ToString & ".SOL_GetListWBNoPol('" & date1 & "','" & date2 & "','" & status & "','" & type & "','" & ColumnFilter & "','" & ValueFilter & "','" & oSBOConnection.DatabaseName.ToString & "')"
        Return "CALL " & oApplicationConfig.AddonDB.ToString & ".SOL_GetListWBNoPol('" & date1 & "','" & date2 & "','" & status & "','" & type & "','" & ColumnFilter & "','" & ValueFilter & "','" & oSBOConnection.DatabaseName.ToString & "')"
    End Function



    Public Function GetTransbySPT(ByVal nospt As String, ByVal status As String, ByVal type As String) As String Implements Queries.GetTransbySPT
        Return "CALL SOL_GET_TRANSAKSIBYSPT('" & nospt & "','" & status & "','" & type & "')"
    End Function
    Public Function GetTransbyNopol(ByVal nopol As String, ByVal status As String, ByVal type As String) As String Implements Queries.GetTransbyNopol
        Return "CALL SOL_GET_TRANSAKSIBYNOPOL('" & nopol & "','" & status & "','" & type & "')"
    End Function

    Public Function GetTransbyNopo(ByVal nopo As String, ByVal status As String, ByVal type As String) As String Implements Queries.GetTransbyNopo
        Return "CALL SOL_GET_TRANSAKSIBYSPT('" & nopo & "','" & status & "','" & type & "')"
    End Function

    Public Function GetTransDelivery(ByVal DocNum As String) As String Implements Queries.GetTransDelivery
        Return "SELECT 1 FROM ODLN T0 WHERE T0.""U_SOL_SPM"" = '" & DocNum & "' AND T0.""CANCELED"" = 'Y' "
    End Function
    Public Function GetTransDelivery2(ByVal DocNum As String) As String Implements Queries.GetTransDelivery2
        Return "SELECT 1 FROM ODLN T0 WHERE T0.""U_SOL_SPM"" = '" & DocNum & "' AND T0.""CANCELED"" = 'Y' "
    End Function
    Public Function UpdateStatusSPM(ByVal docent As Integer, ByVal status As String) As String Implements Queries.UpdateStatusSPM
        Return "UPDATE ""@SOL_WB"" SET ""U_SOL_STATUS"" = '" & status & "' WHERE ""DocEntry = " & docent
    End Function

    Public Function GetLastNumber(ByVal year As String, ByVal tipe As String) As String Implements Queries.GetLastNumber
        Return "SELECT ""DOCNUM"" FROM SOL_WB WHERE TO_NVARCHAR(""CreateDate"", 'YYYY') = '" & year & "' and ""U_SOL_TIPE"" = '" & tipe & "'"
    End Function

    Public Function GetComboBoxOperator(ByVal tipe As String) As String Implements Queries.GetComboBoxOperator
        Return "SELECT  ""Code"" as ""Code"", ""Name"" as ""Name"" FROM " + oSBOConnection.DatabaseName.ToString + ".""@SOL_OPERATOR"" WHERE ""U_SOL_TYPE"" = '" & tipe & "'  ORDER BY ""Name"" ASC"
    End Function
    Public Function GetComboBoxOperator_SOL_USER() As String Implements Queries.GetComboBoxOperator_SOL_USER
        Return "SELECT  ""UNAME"" as ""Name"" FROM " + oApplicationConfig.AddonDB.ToString + ".""SOL_USER"" ORDER BY ""UNAME"" ASC"
    End Function

    Public Function GetComboBoxKendaraan() As String Implements Queries.GetComboBoxKendaraan
        Return "SELECT  ""Code"" as ""Code"", ""Name"" as ""Name"" FROM " + oSBOConnection.DatabaseName.ToString + ".""@SOL_JENISKENDARAAN"" ORDER BY ""Name"" ASC"
    End Function

    Public Function GetComboBoxInsentifJarak() As String Implements Queries.GetComboBoxInsentifJarak
        Return "SELECT '0' as ""CODE"",'-' as ""NAME"",0 as ""INSENTIF"" FROM DUMMY UNION ALL SELECT ""CODE"",""NAME"", ""SOL_INSENTIF"" FROM " + oApplicationConfig.AddonDB.ToString + ".SOL_POSPANTAU WHERE ""SOL_ISEMPLACEMENT"" = 'N'"
    End Function


    Public Function GetComboBoxShift() As String Implements Queries.GetComboBoxShift
        Dim nana As String = "SELECT ""FldValue"" as ""Code"",""Descr"" as ""Name"" FROM " + oSBOConnection.DatabaseName.ToString + ".UFD1 WHERE ""TableID"" = 'OIGE' AND ""FieldID"" = '53' AND ""FldValue"" <> '0'"
        Return "SELECT ""FldValue"" as ""Code"",""Descr"" as ""Name"" FROM " + oSBOConnection.DatabaseName.ToString + ".UFD1 WHERE ""TableID"" = 'OIGE' AND ""FieldID"" = '53' AND ""FldValue"" <> '0'"
    End Function

    Public Function GetMBS() As String Implements Queries.GetMBS
        Return "SELECT  ""SOL_TIPLER"" as ""TIPLER"", ""SOL_TAP"" as ""TAP"", ""SOL_SPTA"" as ""SPTA"" FROM " + oSBOConnection.DatabaseName.ToString + ".""SOL_MASTER_MBS"" WHERE ""SOL_STATUS"" = 1"
    End Function

    Public Function GetListSPTA(ByVal tgl As String, ByVal flag As String, ByVal OrderBy As String) As String Implements Queries.GetListSPTA
        Return "CALL " & oApplicationConfig.AddonDB.ToString & ".SOL_GETLISTKONTRAK('" & tgl & "','" & flag & "', '" & OrderBy & "')"
    End Function

    Public Function GetListSPTAFILTER(ByVal tgl As String, ByVal flag As String, ByVal OrderBy As String, ByVal NamaPetani As String, ByVal RegisterID As String, ByVal NoKontrak As String, ByVal KDProvinsi As String, ByVal KDKabupaten As String, ByVal KDKecamatan As String, ByVal KDDesa As String, ByVal FilterColumn As String, ByVal FilterValue As String) As String Implements Queries.GetListSPTAFILTER
        Dim halo As String = "CALL " & oApplicationConfig.AddonDB.ToString & ".SOL_GETLISTKONTRAKFILTER('" & tgl & "','" & flag & "','" & OrderBy & "','" & NamaPetani & "','" & RegisterID & "','" & NoKontrak & "','" & KDProvinsi & "','" & KDKabupaten & "','" & KDKecamatan & "','" & KDDesa & "','" & FilterColumn & "','" & FilterValue & "')"
        Return "CALL " & oApplicationConfig.AddonDB.ToString & ".SOL_GETLISTKONTRAKFILTER('" & tgl & "','" & flag & "','" & OrderBy & "','" & NamaPetani & "','" & RegisterID & "','" & NoKontrak & "','" & KDProvinsi & "','" & KDKabupaten & "','" & KDKecamatan & "','" & KDDesa & "','" & FilterColumn & "','" & FilterValue & "')"

    End Function

#End Region

End Class
