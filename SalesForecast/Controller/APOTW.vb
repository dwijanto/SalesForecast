Public Class APOTW
    Public Sub Generate(myForm As Object, e As APOTWEventArgs)

        Dim sqlstr As String = String.Empty '= "select * from sales.sfcmmf;"

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("3740 GSTW_{0:yyyyMMdd}.xlsx", e.startPeriod)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then


            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable


            'sqlstr = "select  to_char(txdate,'YYYYMM') as ""Calendar year/Month"",tx.cmmf as ""CMMF"",'M037HK'::text as ""Market"" ,case  " &
            '    " when sales.get_producttype(productlinegps,brand) = 'CKW - Lago' then 'C37H' when sales.get_producttype(productlinegps,brand) = 'CKW - Tefal' and  mla in ('W9000335','W9000341','W900') then 'C37H'" &
            '    " Else '3730' end as  ""Log.Distrib.Centre"" ,mla as ""Main Local Account"",null::text as ""Forecast Group"",null::text as ""Sales Force"",null::text as ""Sales Organization"",null::text as ""Sales Group""," &
            '    " salesforecast as ""Commercial Qty"" from sales.sfmlatxhk tx " &
            '    " left join sales.sfcmmf c on c.cmmf = tx.cmmf" &
            '    " union all" &
            '    " select  to_char(txdate,'YYYYMM') as ""Calendar year/Month"",tx.cmmf as ""CMMF"",'M037HK'::text as ""Market"" ,case when sales.get_producttype(productlinegps,brand) = 'CKW - Lago' then 'C37H' " &
            '    " when sales.get_producttype(productlinegps,brand) = 'CKW - Tefal' then 'C37H' Else '3730' end as  ""Log.Distrib.Centre"" ,null::text as ""Main Local Account"",groupname as ""Forecast Group"",null::text as ""Sales Force""," &
            '    " null::text as ""Sales Organization"",null::text as ""Sales Group"",salesforecast as ""Commercial Qty"" from sales.sfgrouptxhk tx left join sales.sfcmmf c on c.cmmf = tx.cmmf"
            'Dim NextYear = String.Format("{0}-{1:MM}-01", e.startPeriod.Year + 1, e.startPeriod)
            Dim NextYear = String.Format("{0:yyyy}-{0:MM}-01", e.startPeriod.AddMonths(13))
            'sqlstr = String.Format("select  to_char(txdate,'YYYYMM') as ""Calendar year/Month"",tx.cmmf as ""CMMF"",'M037TW'::text as ""Market"" ,case when tx.groupname = 'DS' then 'C37T'  Else '3740' end as  ""Log.Distrib.Centre"" ,null::text as ""Main Local Account"",groupname as ""Forecast Group"",null::text as ""Sales Force"", null::text as ""Sales Organization"",null::text as ""Sales Group"",salesforecast as ""Commercial Qty"",'PC'::text as ""Unit"" from sales.sfgrouptxtw tx left join sales.sfcmmf c on c.cmmf = tx.cmmf" &
            '                       " where tx.txdate >= '{0:yyyy-MM-01}' and tx.txdate < '{1}'", e.startPeriod, NextYear)
            'Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx", "A1")
            sqlstr = String.Format("select 'NR', to_char(txdate,'YYYYMM') as ""Calendar year/Month"",tx.cmmf as ""CMMF"",case when tx.groupname = 'DS' then 'C37T'  Else '3740' end as  ""Log.Distrib.Centre"" ,null::text as ""Main Local Account"",groupname as ""Forecast Group"",null::text as ""Sales Force"", null::text as ""Sales Organization"",null::text as ""Sales Group"",salesforecast as ""Commercial Qty"",'PC'::text as ""Unit"" from sales.sfgrouptxtw tx left join sales.sfcmmf c on c.cmmf = tx.cmmf" &
                                   " where tx.txdate >= '{0:yyyy-MM-01}' and tx.txdate < '{1}'", e.startPeriod, NextYear)
            'Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx", "A1")
            Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplateTW.xltx", "A2", False)
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
Public Class APOTWEventArgs
    Inherits EventArgs

    Public Property startPeriod As Date

End Class