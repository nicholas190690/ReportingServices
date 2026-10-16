Public Interface Queries

    Function SeparatorSQL() As String
    Function CheckUDFSQL(ByVal tableName As String, ByVal fieldName As String) As String
    Function CheckUDOSQL(ByVal tableName As String) As String
    Function CheckQueryCategorySQL(ByVal categoryName As String) As String
    Function CheckQuerySQL(ByVal categoryName As String, ByVal queryName As String) As String
    Function CheckFMSSQL(ByVal formID As String, ByVal itemID As String) As String
    Function CheckFunctionExistsSQL(ByVal DBName As String, ByVal FunctionName As String) As String
    Function CheckSPExistsSQL(ByVal DBName As String, ByVal SPName As String) As String
    Function GetSeries(ByVal SeriesName As String, ByVal ObjectCode As String) As String
    Function GetTransaksi(ByVal TipeJual As String, ByVal Date1 As String) As String
    Function GetListSPMSPB(ByVal DateFrom As String, ByVal DateTo As String, ByVal Status As String, ByVal flag As String) As String
    Function GetCurrent() As String
    Function GetToleransi() As String
    Function GetShift() As String
    Function GetForm(ByVal Docentry As Integer, ByVal type As String) As String
    Function GetFormTicket(ByVal Docentry As Integer, ByVal type As String) As String
    Function GetLastScale(ByVal Timbangan As String, ByVal shift As String, ByVal tanggal As String) As String
    Function GetBagging(ByVal Mesin As String, ByVal shift As String, ByRef date1 As String) As String
    Function GetRawSugarList(ByVal stat As String, ByVal date1 As String, ByVal date2 As String) As String
    Function GetSumRawSugar(ByVal mesin As String, ByVal date1 As String, ByVal shift As String) As String
    Function GetRawSugarDetail(ByVal shift As String, ByVal date1 As String, ByVal docentry As String) As String
    Function GetMesin(ByVal Type As String) As String
    Function GetBaggingConfirmList(ByVal mesin As String, ByVal date1 As String, ByVal shift As String) As String
    Function GetBaggingConfirmRecap(ByVal mesin As String, ByVal date1 As String, ByVal shift As String) As String
    Function GetAllBagging(ByVal stat As String, ByVal date1 As String, ByVal date2 As String) As String
    Function GetBaggingList(ByVal shift As String, ByVal date1 As String, ByVal docentry As String) As String
    Function GetRawSugarConfirmList(ByVal mesin As String, ByVal date1 As String, ByVal shift As String) As String
    Function GetWB(ByVal date1 As String, ByVal date2 As String, ByVal status As String, ByVal type As String, ByVal ColumnFilter As String, ByVal ValueFilter As String) As String
    Function GetTransbySPT(ByVal nospt As String, ByVal status As String, ByVal type As String) As String
    Function GetTransbyNopol(ByVal nopol As String, ByVal status As String, ByVal type As String) As String
    Function GetTransbyNopo(ByVal nopo As String, ByVal status As String, ByVal type As String) As String
    Function UpdateStatusSPM(ByVal docent As Integer, ByVal status As String) As String
    Function GetTransDelivery(ByVal DocNum As String) As String
    Function GetTransDelivery2(ByVal DocNum As String) As String
    Function GetShiftFilter(ByVal Type As String) As String
    Function GetLastNumber(ByVal year As String, ByVal tipe As String) As String
    Function GetComboBoxOperator(ByVal tipe As String) As String
    Function GetComboBoxOperator_SOL_USER() As String
    Function GetComboBoxKendaraan() As String
    Function GetComboBoxShift() As String
    Function GetComboBoxInsentifJarak() As String
    Function GetMBS() As String
    Function GetListSPTA(ByVal tgl As String, ByVal flag As String, ByVal OrderBy As String) As String
    Function GetListSPTAFILTER(ByVal tgl As String, ByVal flag As String, ByVal OrderBy As String, ByVal NamaPetani As String, ByVal RegisterID As String, ByVal NoKontrak As String, ByVal KDProvinsi As String, ByVal KDKabupaten As String, ByVal KDKecamatan As String, ByVal KDDesa As String, ByVal FilterColumn As String, ByVal FilterValue As String) As String
End Interface
