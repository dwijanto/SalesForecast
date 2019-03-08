Imports Microsoft.Office.Interop
Public Class APOTH
    Public Sub Generate(myForm As Object, e As APOTHEventArgs)

        Dim sqlstr As String = String.Empty '= "select * from sales.sfcmmf;"

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("3730 GSTH_{0:yyyyMMdd}.xlsx", e.startPeriod)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then


            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1


            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            'Dim NextYear = String.Format("{0}-{1:MM}-01", e.startPeriod.Year + 1, e.startPeriod)
            Dim NextYear = String.Format("{0:yyyy}-{0:MM}-01", e.startPeriod.AddMonths(13))
            'sqlstr = String.Format("select  to_char(txdate,'YYYYMM') as ""Calendar year/Month"",tx.cmmf as ""CMMF"",'M066SG'::text as ""Market"" , 6610 as  ""Log.Distrib.Centre"" ," &
            '         " mla as ""Main Local Account"",null::text as ""Forecast Group"",null::text as ""Sales Force"", null::text as ""Sales Organization""," &
            '         " null::text as ""Sales Group"",sum(salesforecast) as ""Commercial Qty""" &            '         " from sales.sfgrouptxsg tx left " &
            '         "join sales.sfgroupmlasg gm on gm.groupid =  tx.groupid " &
            '         " where  tx.txdate >= '{0:yyyy-MM}-01' and  tx.txdate < '{1}' group by ""Calendar year/Month"",""CMMF"",""Market"",""Log.Distrib.Centre"" ,mla", e.startPeriod, NextYear)
            'sqlstr = String.Format("select to_char(txdate,'YYYYMM') as ""Calendar year/Month"",cmmf::bigint as ""CMMF"",'M064TH'::text as  ""Market"",6410::integer as ""Log.Distrib.Centre"",mc.mla::text as ""Main Local Account""," &
            '                       " null::text as ""Forecast Group"", null::text as ""Sales Force"",null::text as ""Sales Organization"",null::text as ""Sales Group"",sum(tx.salesforecast) as ""Commercial Qty""," &
            '                       " 'PC'::text as ""Unit"" from  sales.sfgrouptxth tx " &
            '                       " left join sales.sfkamassignmentth ka on ka.id = tx.kamassignmentid" &
            '                       " left join sales.sfmlacardnameth mc on mc.id = ka.mlacardnameid" &
            '                       " where not tx.salesforecast isnull and tx.txdate >= '{0:yyyy-MM}-01' and  tx.txdate < '{1}'  " &
            '                       " group by txdate,cmmf,mla order by txdate,mla,cmmf", e.startPeriod, NextYear)
            'sqlstr = String.Format("select to_char(txdate,'YYYYMM') as ""Calendar year/Month"",tx.cmmf::bigint as ""CMMF"",'M064TH'::text as  ""Market"",6410::integer as ""Log.Distrib.Centre"",mc.mla::text as ""Main Local Account""," &
            '                       " null::text as ""Forecast Group"", null::text as ""Sales Force"",null::text as ""Sales Organization"",null::text as ""Sales Group"",sum(tx.salesforecast) as ""Commercial Qty""," &
            '                       " 'PC'::text as ""Unit"" from  sales.sfgrouptxth tx " &
            '                       " left join sales.sfkamassignmentth ka on ka.id = tx.kamassignmentid" &
            '                       " left join sales.sfmlacardnameth mc on mc.id = ka.mlacardnameid" &
            '                       " left join sales.sfcmmfth c on c.cmmf = tx.cmmf" &
            '                       " where not tx.salesforecast isnull and tx.txdate >= '{0:yyyy-MM}-01' and  tx.txdate < '{1}'  and not c.cmmf isnull" &
            '                       " group by txdate,tx.cmmf,mla order by txdate,mla,tx.cmmf", e.startPeriod, NextYear)
            'Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx", "A1")
            sqlstr = String.Format("select 'NR', to_char(txdate,'YYYYMM') as ""Calendar year/Month"",tx.cmmf::bigint as ""CMMF"",6410::integer as ""Log.Distrib.Centre"",mc.mla::text as ""Main Local Account""," &
                                  " null::text as ""Forecast Group"", null::text as ""Sales Force"",null::text as ""Sales Organization"",null::text as ""Sales Group"",sum(tx.salesforecast) as ""Commercial Qty""," &
                                  " 'PC'::text as ""Unit"" from  sales.sfgrouptxth tx " &
                                  " left join sales.sfkamassignmentth ka on ka.id = tx.kamassignmentid" &
                                  " left join sales.sfmlacardnameth mc on mc.id = ka.mlacardnameid" &
                                  " left join sales.sfcmmfth c on c.cmmf = tx.cmmf" &
                                  " where not tx.salesforecast isnull and tx.txdate >= '{0:yyyy-MM}-01' and  tx.txdate < '{1}'  and not c.cmmf isnull" &
                                  " group by txdate,tx.cmmf,mla order by txdate,mla,tx.cmmf", e.startPeriod, NextYear)

            Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplateTH.xltx", "A2", False)
            myreport.Run(myForm, e)

        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        ' Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        ' osheet.Cells.EntireColumn.AutoFit()
    End Sub

    Private Sub PivotTable()
        'Throw New NotImplementedException
    End Sub
End Class
Public Class APOTHEventArgs
    Inherits EventArgs

    Public Property startPeriod As Date

End Class