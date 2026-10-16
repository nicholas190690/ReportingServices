Imports System.IO
Imports System.Globalization
Imports System.Net.Mail
Imports Microsoft.Office.Interop
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine


Public Class Service1
    Dim timer As New Timers.Timer

    Private Sub disconnect()
        If oDBConnection.Connected Then
            oDBConnection.Disconnect()
        End If
    End Sub

    Private Sub connect()
        If Not oDBConnection Is Nothing And Not oDBConnection.Connected Then
            oDBConnection.Connect()
        End If
        If Not oDBConnection.Connected Then
            Throw New Exception("Database Connection Is Not Connected, Please Check!!!")
        End If
    End Sub


    Public Shared Sub TraceService(ByVal content As String)
        Dim fs As New FileStream(oApplicationConfig.WorkFolder & "\Log\Log_" & DateTime.Now.ToString("yyyyMMdd") & ".txt", FileMode.OpenOrCreate, System.IO.FileAccess.Write)
        Dim sw As New StreamWriter(fs)

        sw.BaseStream.Seek(0, SeekOrigin.End)
        sw.WriteLine(content)
        sw.Flush()
        sw.Close()
    End Sub

    Protected Overrides Sub OnStart(ByVal args() As String)
        Try

            oConfigXMLParser.read()
            isInit = True
            If oConfigXMLParser.isConfigFileExists Then
                If oConfigXMLParser.read Then
                    isInit = False
                End If
            End If

            If Directory.Exists(oApplicationConfig.WorkFolder) = False Then
                Directory.CreateDirectory(oApplicationConfig.WorkFolder)
            End If

            If Directory.Exists(oApplicationConfig.WorkFolder & "\Log") = False Then
                Directory.CreateDirectory(oApplicationConfig.WorkFolder & "\Log")
            End If



            TraceService(DateTime.Now + " start service")
            timer = New Timers.Timer
            timer.Interval = 1000
            AddHandler timer.Elapsed, AddressOf MyTickHandler
            timer.Enabled = True
        Catch ex As Exception
            TraceService(ex.Message.ToString)
        End Try
    End Sub

    Protected Overrides Sub OnStop()
        timer.Enabled = False
        disconnect()
        TraceService(DateTime.Now + " stop service")
    End Sub

    Private Sub MyTickHandler(ByVal obj As Object, ByVal e As EventArgs)
        ' If (DateTime.Now.Second = 0) Then
        If (DateTime.Now.Second = 0 And DateTime.Now.Hour = 9 And DateTime.Now.Minute = 48) Then

            If BackgroundWorker1.IsBusy = False Then
                TraceService(DateTime.Now & " Start Proses  ")
                BackgroundWorker1.RunWorkerAsync()
            End If
        End If
    End Sub

    Private Function email(ByVal subject As String, ByVal body As String, ByVal emailTO As String, Optional ByVal emailCC As String = Nothing, Optional ByVal attachment As String() = Nothing) As Integer
        Try
            Dim SmtpServer As New SmtpClient()
            Dim mail As New MailMessage()


            SmtpServer.Credentials = New  _
            Net.NetworkCredential(My.Settings.emailfrom, My.Settings.emailpassword)
            SmtpServer.Port = My.Settings.port
            SmtpServer.Host = My.Settings.smtp
            mail = New MailMessage()
            mail.From = New MailAddress(My.Settings.emailfrom)
            If (attachment Is Nothing) = False Then
                For Each attacfile In attachment
                    Dim attach As New Net.Mail.Attachment(attacfile)
                    mail.Attachments.Add(attach)
                Next
            End If

            If (emailCC Is Nothing) = False Then
                mail.CC.Add(emailCC)
            End If
            mail.To.Add(emailTO)
            mail.Subject = subject

            mail.Body = body
            SmtpServer.Send(mail)
            TraceService(DateTime.Now & "  Email Notifikasi " & subject & " berhasil")
            Return 1
        Catch ex As Exception
            TraceService(DateTime.Now & "  Email Notifikasi " & subject & " gagal " & ex.Message.ToString)
            Return -1
        End Try
    End Function


    Private Sub ExportToExcel(ByVal dtTemp As DataTable, ByVal filepath As String)
        Try

            TraceService("Masuk export excel")
            Dim strFileName As String = filepath
            If System.IO.File.Exists(strFileName) Then

                System.IO.File.Delete(strFileName)

            End If

            TraceService("Selesai create excel")

            Dim _excel As New Excel.Application
            Dim wBook As Excel.Workbook
            Dim wSheet As Excel.Worksheet



            wBook = _excel.Workbooks.Add()
            wSheet = wBook.ActiveSheet()
            Dim dt As System.Data.DataTable = dtTemp
            Dim dc As System.Data.DataColumn
            Dim dr As System.Data.DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 0


            TraceService("buat kolom")
            For Each dc In dt.Columns
                colIndex = colIndex + 1
                wSheet.Cells(1, colIndex) = dc.ColumnName
                'If listdeny.Contains(dc.ColumnName.ToString) Then
                '    wSheet.Columns(colIndex).Hidden = True
                'End If
            Next

            For Each dr In dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns

                    colIndex = colIndex + 1
                    wSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                Next
            Next
            wSheet.Columns.AutoFit()
            colIndex = 0
            For Each dc In dt.Columns
                colIndex = colIndex + 1
                wSheet.Cells(1, colIndex) = dc.ColumnName
                'If listdeny.Contains(dc.ColumnName.ToString) Then
                '    wSheet.Columns(colIndex).Hidden = True
                'End If
            Next
            wBook.SaveAs(strFileName)
            ReleaseObject(wSheet)
            wBook.Close(False)
            ReleaseObject(wBook)
            _excel.Quit()
            ReleaseObject(_excel)
            GC.Collect()

            TraceService("File Export Successfully to " + filepath + "!")
        Catch ex As Exception
            TraceService(ex.Message.ToString)
        End Try
     
    End Sub


    Private Sub ExportToExcel2(ByVal dtTemp As DataTable, ByVal dtTemp2 As DataTable, ByVal filepath As String)
        Try

            TraceService("Masuk export excel")
            Dim strFileName As String = filepath
            If System.IO.File.Exists(strFileName) Then

                System.IO.File.Delete(strFileName)

            End If



            Dim _excel As New Excel.Application
            Dim wBook As Excel.Workbook
            Dim wSheet As Excel.Worksheet

            TraceService("Masuk 1 sheet")

            wBook = _excel.Workbooks.Add()
            wSheet = CType(wBook.Worksheets.Add(), Excel.Worksheet)

            wSheet.Name = "AvaragePrice"
            Dim dt As System.Data.DataTable = dtTemp2
            Dim dc As System.Data.DataColumn
            Dim dr As System.Data.DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 0
            For Each dc In dt.Columns
                colIndex = colIndex + 1
                wSheet.Cells(1, colIndex) = dc.ColumnName
            Next

            For Each dr In dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns

                    colIndex = colIndex + 1
                    wSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                Next
            Next
            wSheet.Columns.AutoFit()
            colIndex = 0
            For Each dc In dt.Columns
                colIndex = colIndex + 1
                wSheet.Cells(1, colIndex) = dc.ColumnName
            Next
            Dim range As Excel.Range
            range = wSheet.Range("B:G") ' Change A:B to the range you want
            range.NumberFormat = "#,##0.00"

            TraceService("Masuk 2 sheet")

            wSheet = CType(wBook.Worksheets.Add(), Excel.Worksheet)

            wSheet.Name = "AvaragePriceSubzone"
            dt = dtTemp

            colIndex = 0
            rowIndex = 0
            For Each dc In dt.Columns
                If dc.ColumnName <> "DATEPARAM" Then
                    colIndex = colIndex + 1
                    wSheet.Cells(1, colIndex) = dc.ColumnName
                End If

            Next

            For Each dr In dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns
                    If dc.ColumnName <> "DATEPARAM" Then
                        colIndex = colIndex + 1
                        wSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)

                    End If
                Next
            Next
            wSheet.Columns.AutoFit()
            colIndex = 0
            For Each dc In dt.Columns
                If dc.ColumnName <> "DATEPARAM" Then
                    colIndex = colIndex + 1
                    wSheet.Cells(1, colIndex) = dc.ColumnName
                End If

            Next

            range = wSheet.Range("B:W") ' Change A:B to the range you want
            range.NumberFormat = "#,##0.00"



            wBook.SaveAs(strFileName)
            ReleaseObject(wSheet)
            ReleaseObject(range)
            wBook.Close(False)
            ReleaseObject(wBook)
            _excel.Quit()
            ReleaseObject(_excel)
            GC.Collect()

            TraceService("File Export Successfully to " + filepath + "!")
        Catch ex As Exception
            TraceService(ex.Message.ToString)
        End Try

    End Sub


    Private Sub ExportToExcel3(ByVal dtTemp As DataTable, ByVal dtTemp2 As DataTable, ByVal dtTemp3 As DataTable, ByVal filepath As String)
        Dim _excel As New Excel.Application
        Dim wBook As Excel.Workbook
        Dim wSheet As Excel.Worksheet
        Dim range As Excel.Range
        Try

            TraceService("Masuk export excel")
            Dim strFileName As String = filepath
            If System.IO.File.Exists(strFileName) Then

                System.IO.File.Delete(strFileName)

            End If




            TraceService("Masuk 1 sheet")

            wBook = _excel.Workbooks.Add()
            wSheet = CType(wBook.Worksheets.Add(), Excel.Worksheet)

            wSheet.Name = "AvaragePrice"
            Dim dt As System.Data.DataTable = dtTemp2
            Dim dc As System.Data.DataColumn
            Dim dr As System.Data.DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 0
            For Each dc In dt.Columns
                colIndex = colIndex + 1
                wSheet.Cells(1, colIndex) = dc.ColumnName
            Next

            For Each dr In dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns

                    colIndex = colIndex + 1
                    wSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                Next
            Next
            wSheet.Columns.AutoFit()
            colIndex = 0
            For Each dc In dt.Columns
                colIndex = colIndex + 1
                wSheet.Cells(1, colIndex) = dc.ColumnName
            Next

            range = wSheet.Range("B:G") ' Change A:B to the range you want
            range.NumberFormat = "#,##0.00"

            TraceService("Masuk 2 sheet")

            wSheet = CType(wBook.Worksheets.Add(), Excel.Worksheet)

            wSheet.Name = "AvaragePriceSubzone"
            dt = dtTemp

            colIndex = 0
            rowIndex = 0
            For Each dc In dt.Columns
                If dc.ColumnName <> "DATEPARAM" Then
                    colIndex = colIndex + 1
                    wSheet.Cells(1, colIndex) = dc.ColumnName
                End If

            Next

            For Each dr In dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns
                    If dc.ColumnName <> "DATEPARAM" Then
                        colIndex = colIndex + 1
                        wSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)

                    End If
                Next
            Next
            wSheet.Columns.AutoFit()
            colIndex = 0
            For Each dc In dt.Columns
                If dc.ColumnName <> "DATEPARAM" Then
                    colIndex = colIndex + 1
                    wSheet.Cells(1, colIndex) = dc.ColumnName
                End If

            Next

            range = wSheet.Range("B:W") ' Change A:B to the range you want
            range.NumberFormat = "#,##0.00"



            TraceService("Masuk 3 sheet")

            wSheet = CType(wBook.Worksheets.Add(), Excel.Worksheet)

            wSheet.Name = "AvarageQtySubzone"
            dt = dtTemp3

            colIndex = 0
            rowIndex = 0
            For Each dc In dt.Columns
                If dc.ColumnName <> "DATEPARAM" Then
                    colIndex = colIndex + 1
                    wSheet.Cells(1, colIndex) = dc.ColumnName
                End If

            Next

            For Each dr In dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns
                    If dc.ColumnName <> "DATEPARAM" Then
                        colIndex = colIndex + 1
                        wSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)

                    End If
                Next
            Next
            wSheet.Columns.AutoFit()
            colIndex = 0
            For Each dc In dt.Columns
                If dc.ColumnName <> "DATEPARAM" Then
                    colIndex = colIndex + 1
                    wSheet.Cells(1, colIndex) = dc.ColumnName
                End If

            Next

            range = wSheet.Range("B:W") ' Change A:B to the range you want
            range.NumberFormat = "#,##0.00"



            TraceService("Save sheet")
            TraceService(strFileName)
            wBook.SaveAs(strFileName)


            TraceService("File Export Successfully to " + filepath + "!")
        Catch ex As Exception
            TraceService(ex.Message.ToString)
        Finally
            ReleaseObject(wSheet)
            ReleaseObject(range)
            wBook.Close(False)
            ReleaseObject(wBook)
            _excel.Quit()
            ReleaseObject(_excel)
            GC.Collect()
        End Try

    End Sub

    Private Sub ReleaseObject(ByVal o As Object)
        Try
            While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
            End While
        Catch
        Finally
            o = Nothing
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        TraceService(DateTime.Now & " masuk do work  ")
        Dim filepath As String = ""

        Dim dt_service As New DataTable
        Dim sql As String
        Dim dt_report As New DataTable
        Dim dt_param As New DataTable
        Dim dt_target As New DataTable
        Dim dt_targetzone As New DataTable
        Dim dt_targetzoneQTY As New DataTable

        Dim to_email As String = "martino@rejosomanisindo.com,aditya@rejosomanisindo.com,rizkapuji@rejosomanisindo.com,ageng.putro@rejosomanisindo.com,endahpuji@rejosomanisindo.com,debbypg@rejosomanisindo.com,rahmatjainuri@gmail.com,tridarmawan1985@gmail.com,ekasari@rejosomanisindo.com"
        Dim cc_email As String = "syahrulistad@rejosomanisindo.com,it@berkahmm.com"

        ' Dim to_email As String = "nicholas@berkahmm.com"
        ' Dim cc_email As String = "nicholas@berkahmm.com"
        Dim attachment(0) As String
        'Dim crReport As New ReportDocument
        'Dim crConnectionInfo As New ConnectionInfo
        'Dim CrTable As Table
        'Dim CrTables As Tables
        'Dim crtableLogoninfo As New TableLogOnInfo()
        Dim dt = DateTime.ParseExact("20190101", "yyyyMMdd", Nothing)
        TraceService(DateTime.Now & " Coba create folder  ")
        If Directory.Exists(oApplicationConfig.WorkFolder & "\Data") = False Then
            Directory.CreateDirectory(oApplicationConfig.WorkFolder & "\Data")
        End If
        Try
            TraceService(DateTime.Now & " Connect to database  ")
            connect()
            TraceService(DateTime.Now & " Call Procedure  ")
            ' sql = "select SERVICE_ID,SERVICE_NAME,EMAILTO,EMAILCC from " & oApplicationConfig.AddonDB & ".RMI_SERVICE where SERVICE_STAT = 'Y'"
            sql = "CALL " & oApplicationConfig.AddonDB & ".RMI_REPORT_KONTRAK4M "
            dt_service = oDBConnection.executeQuery(sql)


            If dt_service.Rows.Count > 0 Then
                ExportToExcel(dt_service, oApplicationConfig.WorkFolder & "\Data\" & "DataPetaniOverPer_" & Now.ToString("yyyyMMdd") & ".xlsx")
                attachment(0) = oApplicationConfig.WorkFolder & "\Data\" & "DataPetaniOverPer_" & Now.ToString("yyyyMMdd") & ".xlsx"
                email("[AUTOEMAIL] REMINDER PETANI OVER 3.5 MILIAR RUPIAH per tanggal " & Today.AddDays(-1).ToString("dd MMMM yyyy"), "Dear All, " & Environment.NewLine & "Berikut terlampir list petani yang sudah melebihi 3.5 Miliar Rupiah , dan status kontrak nya masih aktif." & Environment.NewLine & "Mohon untuk di tindak lanjuti. ", to_email, cc_email, attachment)
                ' email("[AUTOEMAIL] Proses Insentif Berat tanggal " & Today.AddDays(-1).ToString("dd MMMM yyyy"), "Dear All, " & Environment.NewLine & "Berhasil proses Insentif Berat sebanyak : " & (SuccessAddDocument / 2).ToString & " GRPO dan " & dt_setting.Rows.Count & " Petani untuk tanggal " & Today.AddDays(-1).ToString("dd MMMM yyyy"), "nicholas@berkahmm.com", "it@rejosomanisindo.com", attachment)
            End If
            TraceService(DateTime.Now & " End of Process Report Service kontrak over 4m")


            TraceService(DateTime.Now & " Start process data price daily")
            sql = "CALL " & oSBOConnection.SBODatabaseName & ".RMI_REPORTAVGTODATEZONE()"
            dt_target = oDBConnection.executeQuery(sql)

            sql = "CALL " & oSBOConnection.SBODatabaseName & ".RMI_REPORTAVGTODATE()"
            dt_targetzone = oDBConnection.executeQuery(sql)

            sql = "CALL " & oSBOConnection.SBODatabaseName & ".RMI_REPORTAVGTODATEZONEQTY()"
            dt_targetzoneQTY = oDBConnection.executeQuery(sql)

            If dt_target.Rows.Count > 0 And dt_targetzone.Rows.Count > 0 Then
                to_email = "apicharts@mitrphol.com,syahrulistad@rejosomanisindo.com"
                cc_email = "sofwan@rejosomanisindo.com,nicholas@rejosomanisindo.com"


                ExportToExcel3(dt_target, dt_targetzone, dt_targetzoneQTY, oApplicationConfig.WorkFolder & "\Data\" & "DataPriceDaily_" & Today.AddDays(-1).ToString("yyyyMMdd") & ".xlsx")
                attachment(0) = oApplicationConfig.WorkFolder & "\Data\" & "DataPriceDaily_" & Today.AddDays(-1).ToString("yyyyMMdd") & ".xlsx"
                email("[AUTOEMAIL] Data Cane Price Daily per  End Of " & Today.AddDays(-1).ToString("dd MMMM yyyy"), "Dear All, " & Environment.NewLine & "Please find attached file for daily report of sugar cane price", to_email, cc_email, attachment)
            End If

            TraceService(DateTime.Now & " End of Process Report data price daily")
            'TraceService(DateTime.Now & " Proses service sebanyak : " & dt_service.Rows.Count)
            'If dt_service.Rows.Count > 0 Then
            '    For i As Integer = 0 To dt_service.Rows.Count - 1

            '        sql = "SELECT T1.* FROM " & oApplicationConfig.AddonDB & ".RMI_SERVICE_D T0 INNER JOIN " & oApplicationConfig.AddonDB & ".RMI_REPORT_MASTER T1 ON T0.REPORT_ID = T1.REPORT_ID WHERE T0.SERVICE_ID = '" & dt_service.Rows(i).Item("SERVICE_ID") & "' "
            '        dt_report = oDBConnection.executeQuery(sql)
            '        TraceService(DateTime.Now & " Proses Report sebanyak : " & dt_report.Rows.Count)
            '        For j As Integer = 0 To dt_report.Rows.Count - 1
            '            Dim attachment(dt_report.Rows.Count - 1) As String
            '            sql = "SELECT T0.* FROM " & oApplicationConfig.AddonDB & ".RMI_REPORT_MASTER_PARAM T0  WHERE T0.REPORT_ID = '" & dt_report.Rows(j).Item("REPORT_ID") & "' "
            '            dt_param = oDBConnection.executeQuery(sql)
            '            TraceService(DateTime.Now & " Proses Param : " & dt_param.Rows.Count)
            '            If dt_report.Rows(j).Item("REPORT_TYPE").ToString = "1" Then
            '                TraceService(DateTime.Now & " Proses generate rpt")
            '                TraceService(DateTime.Now & " report  : " & dt_report.Rows(j).Item("REPORT_PATH").ToString)
            '                crReport.Load(dt_report.Rows(j).Item("REPORT_PATH").ToString)
            '                TraceService(DateTime.Now & " berhasil di load")
            '                For k As Integer = 0 To dt_param.Rows.Count - 1
            '                    If dt_param.Rows(k).Item("PARAMTYPE").ToString = "04" Then
            '                        crReport.SetParameterValue(k, DateTime.ParseExact(Now.AddDays(-1).ToString("yyyyMMdd"), "yyyyMMdd", Nothing))
            '                    End If
            '                Next
            '                crConnectionInfo.ServerName = oSBOConnection.ServerName & ":" & oSBOConnection.Port
            '                crConnectionInfo.UserID = oSBOConnection.DBUserName.ToString
            '                crConnectionInfo.Password = oSBOConnection.DBPassword.ToString

            '                crConnectionInfo.Type = ConnectionInfoType.Unknown
            '                CrTables = crReport.Database.Tables
            '                For Each CrTable In CrTables
            '                    crtableLogoninfo = CrTable.LogOnInfo
            '                    crtableLogoninfo.ConnectionInfo = crConnectionInfo
            '                    CrTable.ApplyLogOnInfo(crtableLogoninfo)
            '                Next
            '                Dim CrExportOptions As ExportOptions
            '                Dim CrDiskFileDestinationOptions As New  _
            '                DiskFileDestinationOptions()
            '                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
            '                CrDiskFileDestinationOptions.DiskFileName = _
            '                                            oApplicationConfig.WorkFolder & "\Data\" & dt_report.Rows(j).Item("REPORT_EXPORTNAME").ToString & "." & dt_report.Rows(j).Item("REPORT_FILETYPE").ToString
            '                CrExportOptions = crReport.ExportOptions
            '                With CrExportOptions
            '                    .ExportDestinationType = ExportDestinationType.DiskFile
            '                    .ExportFormatType = ExportFormatType.PortableDocFormat
            '                    .DestinationOptions = CrDiskFileDestinationOptions
            '                    .FormatOptions = CrFormatTypeOptions
            '                End With
            '                crReport.Export()

            '            Else

            '            End If
            '            attachment(j) = oApplicationConfig.WorkFolder & "\Data\" & dt_report.Rows(j).Item("REPORT_EXPORTNAME").ToString & "_" & Today.ToString("yyyymmdd") & "." & dt_report.Rows(j).Item("REPORT_FILETYPE").ToString
            '        Next
            '        TraceService(DateTime.Now & " Email notifikasi")

            '    Next
            'End If



            'ExportToExcel(dt_setting, "D:\asd.xlsx")
        Catch ex As Exception
            TraceService(DateTime.Now & "   " & ex.Message.ToString)
        Finally
            disconnect()
        End Try
    End Sub
End Class
