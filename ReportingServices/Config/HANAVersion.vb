Public Class HANAVersion
    Implements Queries
    Private DatabaseName As String

    Public Sub New(ByVal DatabaseName As String)
        Me.DatabaseName = DatabaseName
    End Sub
    Public Function SeparatorSQL() As String Implements Queries.SeparatorSQL
        Return "SELECT ""DecSep"", ""ThousSep"" FROM ""OADM"""
    End Function

    Public Function CheckUDFSQL(ByVal tableName As String, ByVal fieldName As String) As String Implements Queries.CheckUDFSQL
        Return "SELECT ""TableID"" FROM ""CUFD"" WHERE ""TableID"" = '" + tableName.Trim + "' AND ""AliasID"" = '" + fieldName.Trim + "'"
    End Function

    Public Function CheckUDOSQL(ByVal tableName As String) As String Implements Queries.CheckUDOSQL
        Return "SELECT ""Code"" FROM ""OUDO"" WHERE ""Code"" = '" + tableName.Trim + "'"
    End Function

    Public Function CheckQueryCategorySQL(ByVal categoryName As String) As String Implements Queries.CheckQueryCategorySQL
        Return "SELECT ""CategoryId"" FROM ""OQCN"" WHERE ""CatName"" = '" + categoryName + "'"
    End Function

    Public Function CheckQuerySQL(ByVal categoryName As String, ByVal queryName As String) As String Implements Queries.CheckQuerySQL
        Return "SELECT ""IntrnalKey"" FROM ""OUQR"" WHERE ""QName"" = '" + queryName + "' AND ""QCategory"" IN (" + CheckQueryCategorySQL(categoryName) + ")"
    End Function

    Public Function CheckFMSSQL(ByVal formID As String, ByVal itemID As String) As String Implements Queries.CheckFMSSQL
        Return "SELECT ""IndexID"", ""QueryId"" FROM ""CSHS"" WHERE ""FormID"" = '" + formID + "' AND ""ItemID"" = '" + itemID + "'"
    End Function

    Public Function CheckFunctionExistsSQL(ByVal DBName As String, ByVal FunctionName As String) As String Implements Queries.CheckFunctionExistsSQL
        Return "SELECT TOP 1 1 FROM OBJECTS WHERE SCHEMA_NAME = '" + DBName + "' AND OBJECT_TYPE = 'FUNCTION' AND OBJECT_NAME = '" + FunctionName + "'"
    End Function

    Public Function CheckSPExistsSQL(ByVal DBName As String, ByVal SPName As String) As String Implements Queries.CheckSPExistsSQL
        Return "SELECT TOP 1 1 FROM OBJECTS WHERE SCHEMA_NAME = '" + DBName + "' AND OBJECT_TYPE = 'PROCEDURE' AND OBJECT_NAME = '" + SPName + "'"
    End Function


    Public Function GetSeries(ByVal SeriesName As String, ByVal ObjectCode As String) As String Implements Queries.GetSeries
        Return "SELECT  ""Series""  as ""Series"" FROM NNM1 WHERE ""SeriesName"" = '" + SeriesName + "' || (SELECT TO_VARCHAR(CURRENT_DATE,'YYYY') FROM DUMMY) AND ""ObjectCode"" = '" + ObjectCode + "'"
    End Function


    Public Function GetTransaksi(ByVal TipeJual As String, ByVal Date1 As String) As String Implements Queries.GetTransaksi

        Return " CALL " + oSBOConnection.DatabaseName.ToString + ".SOL_GET_TRANSAKSI('" & TipeJual & "','" & Date1 & "')"
    End Function

    Public Function GetListSPMSPB(ByVal DateFrom As String, ByVal DateTo As String, ByVal Status As String, ByVal flag As String) As String Implements Queries.GetListSPMSPB
        Return " CALL " + oSBOConnection.DatabaseName.ToString + ".SOL_GetListSPMSPB(' " & DateFrom & "','" & DateTo & "','" & Status & "','" & flag & "')"

    End Function


    Public Function GetCurrent() As String Implements Queries.GetCurrent
        Return "SELECT TO_NVARCHAR(CURRENT_DATE,'YYYYMMDD') AS ""CDate"",TO_NVARCHAR(CURRENT_TIME,'HH24MI') as ""Ctime"",current_date as ""CDate2"" FROM DUMMY"
    End Function


    Public Function GetToleransi() As String Implements Queries.GetToleransi
        Return "SELECT TOP 1 ""U_SOL_PERCENTAGE"" as ""Percentage"",""U_SOL_MIN"" AS ""Minimum"" , ""U_SOL_MAX"" as ""Maximum"",""U_SOL_PERCENTAGE_WB"" as ""Percentage2"" FROM ""@SOL_TOLERANSI"" "
    End Function

    Public Function GetForm(ByVal Docentry As Integer, ByVal type As String) As String Implements Queries.GetForm
        Return "CALL " + oApplicationConfig.AddonDB.ToString + ".SOL_FORM_SPMSPB(" + Docentry.ToString + ") "
    End Function


    Public Function GetFormTicket(ByVal Docentry As Integer, ByVal type As String) As String Implements Queries.GetFormTicket
        Return "CALL " + oSBOConnection.DatabaseName.ToString + ".SOL_FORM_TICKET(" + Docentry.ToString + ",'" & type & "') "
    End Function


    Public Function GetLastScale(ByVal Timbangan As String, ByVal shift As String, ByVal tanggal As String) As String Implements Queries.GetLastScale
        Return "SELECT IFNULL(MAX(T0.""U_SOL_END""),0) as ""MAXID"" FROM ""@SOL_RAW"" T0 where T0.""U_SOL_MESIN"" = '" + Timbangan + "' and t0.""U_SOL_DATE"" = '" & tanggal & "' and t0.""U_SOL_SHIFT"" = '" & shift & "' "
    End Function


    Public Function GetBagging(ByVal Mesin As String, ByVal shift As String, ByRef date1 As String) As String Implements Queries.GetBagging
        Return "SELECT T0.""DocEntry"" as ""DocEntry"" , T0.""U_SOL_WEIGHT"" AS ""Weight"" , T0.""U_SOL_SUM"" AS ""Bag"" , T0.""U_SOL_FAIL"" as ""Failed"" FROM ""@SOL_BAGGING_H"" T0 where T0.""U_SOL_SHIFT"" = '" & shift & "' and T0.""U_SOL_MESIN"" = '" & Mesin & "' AND T0.""U_SOL_DATE"" = '" & date1 & "'"
    End Function

    Public Function GetRawSugarList(ByVal stat As String, ByVal date1 As String, ByVal date2 As String) As String Implements Queries.GetRawSugarList
        Return "CALL " + oSBOConnection.DatabaseName.ToString + ".SOL_GET_LISTRAWSUGARHEADER('" & stat & "','" & date1 & "','" & date2 & "')"
    End Function


    Public Function GetMesin(ByVal Type As String) As String Implements Queries.GetMesin
        Return "SELECT * FROM " & oSBOConnection.DatabaseName.ToString & ".""@SOL_MESIN"" where ""U_SOL_TYPE"" = '" & Type & "'"
    End Function

    Public Function GetShift() As String Implements Queries.GetShift
        Return "SELECT ""U_SOL_SHIFT"" AS ""Shift"",""U_SOL_START"" AS ""Start"" ,""U_SOL_END"" as ""End"", ""U_SOL_TYPE"" as ""Type"" FROM " & oSBOConnection.DatabaseName.ToString & ".""@SOL_SHIFT"" "
    End Function


    Public Function GetShiftFilter(ByVal Type As String) As String Implements Queries.GetShiftFilter
        Return "SELECT ""U_SOL_SHIFT"" AS ""Shift"",""U_SOL_START"" AS ""Start"" ,""U_SOL_END"" as ""End"", ""U_SOL_TYPE"" as ""Type"" FROM " & oSBOConnection.DatabaseName.ToString & ".""@SOL_SHIFT"" WHERE ""U_SOL_TYPE"" = '" & Type.ToString & "' ORDER BY ""U_SOL_SHIFT"" ASC"
    End Function

    Public Function GetSumRawSugar(ByVal mesin As String, ByVal date1 As String, ByVal shift As String) As String Implements Queries.GetSumRawSugar
        Return "SELECT SUM(""U_SOL_QTY2"") AS ""TOTALWEIGHT"" FROM ""@SOL_RAW"" WHERE ""U_SOL_MESIN"" = '" & mesin & "' AND ""U_SOL_DATE"" = '" & date1 & "'"
    End Function



    Public Function GetRawSugarDetail(ByVal shift As String, ByVal date1 As String, ByVal docentry As String) As String Implements Queries.GetRawSugarDetail
        Dim query As String = ""
        query = "CALL " & oSBOConnection.DatabaseName.ToString & ".SOL_GET_LISTRAWSUGARDETAIL_DIAPI('" & shift & "','" & date1 & "'," & docentry & ")"
        Return query
    End Function


    Public Function GetBaggingConfirmList(ByVal mesin As String, ByVal date1 As String, ByVal shift As String) As String Implements Queries.GetBaggingConfirmList
        If mesin = "All" Then
            mesin = "T0.""U_SOL_MESIN"""
        Else
            mesin = "'" & mesin & "'"
        End If

        Return "SELECT T0.""DocEntry"" AS ""DocEntry"" , T0.""U_SOL_SHIFT"" as ""Shift"" ,T1.""Name"" as ""Mesin Bagging"",RIGHT('00' ||T0.""U_SOL_TIME"",2) || ':00'  as ""Waktu"" , T0.""U_SOL_SUM"" AS ""Jumlah Karung"" , T0.""U_SOL_SUM"" AS ""Jumlah Karung Confirmed"",'0' AS ""Jumlah Karung Cleaning"",'' as ""Comments"", IFNULL(T0.""U_SOL_SUMOVER"",0) AS ""Jumlah Karung Over"", IFNULL(T0.""U_SOL_SUMUNDER"",0) AS ""Jumlah Karung Under"", ""U_SOL_MAXCOUNT"" AS ""Counter"" FROM " + oSBOConnection.DatabaseName.ToString + ".""@SOL_BAGGING"" T0 LEFT JOIN " + oSBOConnection.DatabaseName.ToString + ".""@SOL_MESIN"" T1 ON T0.""U_SOL_MESIN"" = T1.""Code"" AND T1.""U_SOL_TYPE"" = '2'  WHERE T0.""U_SOL_STAT"" = 'Open' and T0.""U_SOL_SHIFT"" = '" & shift & "' AND T0.""U_SOL_DATE"" = '" & date1 & "' AND T0.""U_SOL_MESIN"" = " & mesin & " ORDER BY T1.""Name"" asc, t0.""U_SOL_TIME"" ASC"
    End Function

    Public Function GetRawSugarConfirmList(ByVal mesin As String, ByVal date1 As String, ByVal shift As String) As String Implements Queries.GetRawSugarConfirmList
        If mesin = "All" Then
            mesin = "T0.""U_SOL_MESIN"""
        Else
            mesin = "'" & mesin & "'"
        End If

        'Dim halo As String = "SELECT T0.""DocEntry"", T0.""U_SOL_SHIFT"" as ""Shift"",T1.""Name"" as ""Nama Mesin"",RIGHT('00' ||T0.""U_SOL_TIME"",2) || ':00'  as ""Waktu"" ,T0.""U_SOL_SUM"" as ""Jumlah Timbang"" , T0.""U_SOL_QTY1"" AS ""Berat"", T0.""U_SOL_QTY1"" AS ""Berat Confirmed"",' ' as ""Comments"" FROM ""RMI_IT"".""@SOL_RAW"" T0 LEFT JOIN ""RMI_IT"".""@SOL_MESIN"" T1 ON T0.""U_SOL_MESIN"" = T1.""Code"" AND T1.""U_SOL_TYPE"" = '1' WHERE T0.""U_SOL_STAT"" = 'Open' and T0.""U_SOL_SHIFT"" = '" & shift & "' AND T0.""U_SOL_DATE"" = '" & date1 & "' AND T0.""U_SOL_MESIN"" = " & mesin & " ORDER BY T1.""Name"" asc, T0.""U_SOL_TIME"" ASC"

        Dim nana As String = "SELECT T0.""DocEntry"", T0.""U_SOL_SHIFT"" as ""Shift"",T1.""Name"" as ""Nama Mesin"",RIGHT('00' ||T0.""U_SOL_TIME"",2) || ':00'  as ""Waktu"" ,T0.""U_SOL_SUM"" as ""Jumlah Timbang"" , T0.""U_SOL_QTY1"" AS ""Berat"", T0.""U_SOL_QTY1"" AS ""Berat Confirmed"",' ' as ""Comments"" FROM " + oSBOConnection.DatabaseName.ToString + ".""@SOL_RAW"" T0 LEFT JOIN " + oSBOConnection.DatabaseName.ToString + ".""@SOL_MESIN"" T1 ON T0.""U_SOL_MESIN"" = T1.""Code"" AND T1.""U_SOL_TYPE"" = '1' WHERE T0.""U_SOL_STAT"" = 'Open' and T0.""U_SOL_SHIFT"" = '" & shift & "' AND T0.""U_SOL_DATE"" = '" & date1 & "' AND T0.""U_SOL_MESIN"" = " & mesin & " ORDER BY T1.""Name"" asc, T0.""U_SOL_TIME"" ASC"
        Return "SELECT T0.""DocEntry"", T0.""U_SOL_SHIFT"" as ""Shift"",T1.""Name"" as ""Nama Mesin"",RIGHT('00' ||T0.""U_SOL_TIME"",2) || ':00'  as ""Waktu"" ,T0.""U_SOL_SUM"" as ""Jumlah Timbang"" , T0.""U_SOL_QTY1"" AS ""Berat"", T0.""U_SOL_QTY1"" AS ""Berat Confirmed"",' ' as ""Comments"" FROM " + oSBOConnection.DatabaseName.ToString + ".""@SOL_RAW"" T0 LEFT JOIN " + oSBOConnection.DatabaseName.ToString + ".""@SOL_MESIN"" T1 ON T0.""U_SOL_MESIN"" = T1.""Code"" AND T1.""U_SOL_TYPE"" = '1' WHERE T0.""U_SOL_STAT"" = 'Open' and T0.""U_SOL_SHIFT"" = '" & shift & "' AND T0.""U_SOL_DATE"" = '" & date1 & "' AND T0.""U_SOL_MESIN"" = " & mesin & " ORDER BY T1.""Name"" asc, T0.""U_SOL_TIME"" ASC"
        'Return "SELECT T0.""DocEntry"" AS ""DocEntry"" , T0.""U_SOL_SHIFT"" as ""Shift"" ,T1.""Name"" as ""Mesin Bagging"",RIGHT('00' ||T0.""U_SOL_TIME"",2) || ':00'  as ""Waktu"" , T0.""U_SOL_SUM"" AS ""Jumlah Karung"" , T0.""U_SOL_SUM"" AS ""Jumlah Karung Confirmed"",'0' AS ""Jumlah Karung Cleaning"",'' as ""Comments"", IFNULL(T0.""U_SOL_SUMOVER"",0) AS ""Jumlah Karung Over"", IFNULL(T0.""U_SOL_SUMUNDER"",0) AS ""Jumlah Karung Under"", ""U_SOL_MAXCOUNT"" AS ""Counter"" FROM " + oSBOConnection.DatabaseName.ToString + ".""@SOL_BAGGING"" T0 LEFT JOIN " + oSBOConnection.DatabaseName.ToString + ".""@SOL_MESIN"" T1 ON T0.""U_SOL_MESIN"" = T1.""Code"" AND T1.""U_SOL_TYPE"" = '2'  WHERE T0.""U_SOL_STAT"" = 'Open' and T0.""U_SOL_SHIFT"" = '" & shift & "' AND T0.""U_SOL_DATE"" = '" & date1 & "' AND T0.""U_SOL_MESIN"" = " & mesin & " ORDER BY T1.""Name"" asc, t0.""U_SOL_TIME"" ASC"
    End Function


    Public Function GetBaggingConfirmRecap(ByVal mesin As String, ByVal date1 As String, ByVal shift As String) As String Implements Queries.GetBaggingConfirmRecap
        If mesin = "All" Then
            mesin = "T0.""U_SOL_MESIN"""
        Else
            mesin = "'" & mesin & "'"
        End If
        Return "SELECT T1.""Name"",SUM(T0.""U_SOL_SUM"") AS ""Jumlah Karung""  FROM " + oSBOConnection.DatabaseName.ToString + ".""@SOL_BAGGING"" T0 LEFT JOIN " + oSBOConnection.DatabaseName.ToString + ".""@SOL_MESIN"" T1 ON T0.""U_SOL_MESIN"" = T1.""Code"" AND T1.""U_SOL_TYPE"" = '2' WHERE  T0.""U_SOL_STAT"" = 'Open' and  T0.""U_SOL_SHIFT"" = '" & shift & "' AND T0.""U_SOL_DATE"" = '" & date1 & "' AND T0.""U_SOL_MESIN"" = " & mesin & " GROUP BY  T1.""Name"" ORDER BY T1.""Name"" ASC"
    End Function

    Public Function GetAllBagging(ByVal stat As String, ByVal date1 As String, ByVal date2 As String) As String Implements Queries.GetAllBagging
        Return "CALL " + oSBOConnection.DatabaseName.ToString + ".SOL_GET_LISTBAGGINGHEADER('" & stat & "','" & date1 & "','" & date2 & "')"
    End Function



    'Public Function GetBaggingStatus(ByVal entry As String) As String Implements Queries.GetBaggingStatus
    '    Return "SELECT 1 FROM ""@SOL_BAGGING_H"" where ""DocEntry"" = " & entry & " and ""U_SOL_STAT"" = 'Confirm'"
    'End Function

    Public Function GetWB(ByVal date1 As String, ByVal date2 As String, ByVal status As String, ByVal type As String, ByVal ColumnFilter As String, ByVal ValueFilter As String) As String Implements Queries.GetWB
        Dim test As String = "CALL " & oApplicationConfig.AddonDB.ToString & ".SOL_GETLISTWB('" & date1 & "','" & date2 & "','" & status & "','" & type & "','" & ColumnFilter & "','" & ValueFilter & "','" & oSBOConnection.DatabaseName.ToString & "')"
        Return "CALL " & oApplicationConfig.AddonDB.ToString & ".SOL_GETLISTWB('" & date1 & "','" & date2 & "','" & status & "','" & type & "','" & ColumnFilter & "','" & ValueFilter & "','" & oSBOConnection.DatabaseName.ToString & "')"
    End Function

    Public Function GetBaggingList(ByVal shift As String, ByVal date1 As String, ByVal docentry As String) As String Implements Queries.GetBaggingList
        Dim query As String = ""
        query = "CALL " & oSBOConnection.DatabaseName.ToString & ".SOL_GET_LISTBAGGINGDETAIL_DIAPI('" & shift & "','" & date1 & "'," & docentry & ")"

        Return query
    End Function


    Public Function GetTransbySPT(ByVal nospt As String, ByVal status As String, ByVal type As String) As String Implements Queries.GetTransbySPT
        Return "CALL " & oApplicationConfig.AddonDB.ToString & ".SOL_GET_TRANSAKSIBYSPT('" & nospt & "','" & status & "','" & type & "','" & oSBOConnection.DatabaseName.ToString & "')"
    End Function

    Public Function GetTransbyNopol(ByVal nopol As String, ByVal status As String, ByVal type As String) As String Implements Queries.GetTransbyNopol
        Return "CALL " & oApplicationConfig.AddonDB.ToString & ".SOL_GET_TRANSAKSIBYNOPOL('" & nopol & "','" & status & "','" & type & "','" & oSBOConnection.DatabaseName.ToString & "')"
    End Function

    Public Function GetTransbyNopo(ByVal nopo As String, ByVal status As String, ByVal type As String) As String Implements Queries.GetTransbyNopo
        Return "CALL " & oApplicationConfig.AddonDB.ToString & ".SOL_GET_TRANSAKSIBYNOPO('" & nopo & "','" & status & "','" & type & "','" & oSBOConnection.DatabaseName.ToString & "')"
    End Function


    Public Function GetTransDelivery(ByVal DocNum As String) As String Implements Queries.GetTransDelivery
        Return "SELECT 1 FROM " & oDBConnection.DatabaseName & ".ODLN T0 WHERE IFNULL(T0.""U_SOL_SPM"",'') = '" & DocNum & "' AND T0.""CANCELED"" = 'Y' "
    End Function

    Public Function GetTransDelivery2(ByVal DocNum As String) As String Implements Queries.GetTransDelivery2
        Return "SELECT 1 FROM " & oDBConnection.DatabaseName & ".ODLN T0 WHERE IFNULL(T0.""U_SOL_SPM"",'') = '" & DocNum & "' "
    End Function

    Public Function UpdateStatusSPM(ByVal docent As Integer, ByVal status As String) As String Implements Queries.UpdateStatusSPM
        Return "UPDATE " & oApplicationConfig.AddonDB.ToString & ".""SOL_WB"" SET ""U_SOL_STATUS"" = '" & status & "' WHERE ""DocEntry"" = " & docent
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

    Public Function GetComboBoxShift() As String Implements Queries.GetComboBoxShift
        Dim nana As String = "SELECT ""FldValue"" as ""Code"",""Descr"" as ""Name"" FROM " + oSBOConnection.DatabaseName.ToString + ".UFD1 WHERE ""TableID"" = 'OIGE' AND ""FieldID"" = '53' AND ""FldValue"" <> '0'"
        Return "SELECT ""FldValue"" as ""Code"",""Descr"" as ""Name"" FROM " + oSBOConnection.DatabaseName.ToString + ".UFD1 WHERE ""TableID"" = 'OIGE' AND ""FieldID"" = '53' AND ""FldValue"" <> '0'"
    End Function

    Public Function GetComboBoxInsentifJarak() As String Implements Queries.GetComboBoxInsentifJarak
        Return "SELECT '0' as ""CODE"",'-' as ""NAME"",0 as ""INSENTIF"" FROM DUMMY UNION ALL SELECT ""CODE"",""NAME"", ""SOL_INSENTIF"" FROM " + oApplicationConfig.AddonDB.ToString + ".SOL_POSPANTAU WHERE ""SOL_ISEMPLACEMENT"" = 'N' AND""SOL_STATUS"" = '1'"
    End Function

    Public Function GetMBS() As String Implements Queries.GetMBS
        Return "SELECT  ""SOL_TIPLER"" as ""TIPLER"",""SOL_DESCRIPTION"" as ""DESCRIPTION"", ""SOL_TAP"" as ""TAP"", ""SOL_SPTA"" as ""SPTA"" FROM " + oApplicationConfig.AddonDB.ToString + ".""SOL_MASTER_MBS"" WHERE ""SOL_STATUS"" = 1"
    End Function


    Public Function GetListSPTA(ByVal tgl As String, ByVal flag As String, ByVal OrderBy As String) As String Implements Queries.GetListSPTA
        Dim halo As String = "CALL " & oApplicationConfig.AddonDB.ToString & ".SOL_GETLISTKONTRAK('" & tgl & "','" & flag & "','" & OrderBy & "')"
        Return "CALL " & oApplicationConfig.AddonDB.ToString & ".SOL_GETLISTKONTRAK('" & tgl & "','" & flag & "','" & OrderBy & "')"

    End Function

    Public Function GetListSPTAFILTER(ByVal tgl As String, ByVal flag As String, ByVal OrderBy As String, ByVal NamaPetani As String, ByVal RegisterID As String, ByVal NoKontrak As String, ByVal KDProvinsi As String, ByVal KDKabupaten As String, ByVal KDKecamatan As String, ByVal KDDesa As String, ByVal FilterColumn As String, ByVal FilterValue As String) As String Implements Queries.GetListSPTAFILTER
        Dim halo As String = "CALL " & oApplicationConfig.AddonDB.ToString & ".SOL_GETLISTKONTRAKFILTER('" & tgl & "','" & flag & "','" & OrderBy & "','" & NamaPetani & "','" & RegisterID & "','" & NoKontrak & "','" & KDProvinsi & "','" & KDKabupaten & "','" & KDKecamatan & "','" & KDDesa & "','" & FilterColumn & "','" & FilterValue & "')"
        Return "CALL " & oApplicationConfig.AddonDB.ToString & ".SOL_GETLISTKONTRAKFILTER('" & tgl & "','" & flag & "','" & OrderBy & "','" & NamaPetani & "','" & RegisterID & "','" & NoKontrak & "','" & KDProvinsi & "','" & KDKabupaten & "','" & KDKecamatan & "','" & KDDesa & "','" & FilterColumn & "','" & FilterValue & "')"

    End Function


End Class
