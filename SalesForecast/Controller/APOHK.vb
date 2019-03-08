Imports Microsoft.Office.Interop

Public Class APOHK
    'Dim myPeriodRange As Dictionary(Of Integer, Date)
    'Dim myKAMAssignment As New KAMAssignmentController
    'Dim myKAMAssignmentList As List(Of KAMAssignmentModel)
    'Dim HKReportProperty1 As HKReportProperty = HKReportProperty.getInstance
    'Dim HKParamDT As DataTable

    Public Sub Generate(myForm As Object, e As APOHKEventArgs)

        Dim sqlstr As String = String.Empty '= "select * from sales.sfcmmf;"

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("3730 GSHK_{0:yyyyMMdd}.xlsx", e.startPeriod)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then


            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1


            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable


            'sqlstr = String.Format("select  to_char(txdate,'YYYYMM') as ""Calendar year/Month"",tx.cmmf as ""CMMF"",'M037HK'::text as ""Market"" ,case  " &
            '    " when sales.get_producttype(productlinegps,brand) = 'CKW - Lago' then 'C37H' when sales.get_producttype(productlinegps,brand) = 'CKW - Tefal' and  mla in ('W9000335','W9000341') then 'C37H'" &
            '    " Else '3730' end as  ""Log.Distrib.Centre"" ,mla as ""Main Local Account"",null::text as ""Forecast Group"",null::text as ""Sales Force"",null::text as ""Sales Organization"",null::text as ""Sales Group""," &
            '    " salesforecast as ""Commercial Qty"" from sales.sfmlatxhk tx " &
            '    " left join sales.sfcmmf c on c.cmmf = tx.cmmf where tx.txdate >= '{0:yyyy-MM}-01'" &
            '    " union all" &
            '    " select  to_char(txdate,'YYYYMM') as ""Calendar year/Month"",tx.cmmf as ""CMMF"",'M037HK'::text as ""Market"" ,case when sales.get_producttype(productlinegps,brand) = 'CKW - Lago' then 'C37H' " &
            '    " when sales.get_producttype(productlinegps,brand) = 'CKW - Lago' then 'C37H' Else '3730' end as  ""Log.Distrib.Centre"" ,null::text as ""Main Local Account"",groupname as ""Forecast Group"",null::text as ""Sales Force""," &
            '    " null::text as ""Sales Organization"",null::text as ""Sales Group"",salesforecast as ""Commercial Qty"" from sales.sfgrouptxhk tx left join sales.sfcmmf c on c.cmmf = tx.cmmf where tx.txdate >= '{1:yyyy-MM}-01'", e.startPeriod, e.startPeriod.AddMonths(6))
            'Dim NextYear = String.Format("{0}-{1:MM}-01", e.startPeriod.Year + 1, e.startPeriod)
            Dim NextYear = String.Format("{0:yyyy}-{0:MM}-01", e.startPeriod.AddMonths(13))
            'sqlstr = String.Format("select  to_char(txdate,'YYYYMM') as ""Calendar year/Month"",tx.cmmf as ""CMMF"",'M037HK'::text as ""Market"" ,case  " &
            '   " when sales.get_producttype(productlinegps,brand) = 'CKW - Lago' then 'C37H' when sales.get_producttype(productlinegps,brand) = 'CKW - Tefal' and  mla in ('W9000335','W9000341') then 'C37H'" &
            '   " Else '3730' end as  ""Log.Distrib.Centre"" ,mla as ""Main Local Account"",null::text as ""Forecast Group"",null::text as ""Sales Force"",null::text as ""Sales Organization"",null::text as ""Sales Group""," &
            '   " sum(salesforecast) as ""Commercial Qty"" from sales.sfcmmf c" &
            '   " left join sales.sfmlatxhk tx  on c.cmmf = tx.cmmf where tx.txdate >= '{0:yyyy-MM}-01' and tx.txdate < '{1:yyyy-MM}-01' group by ""Calendar year/Month"",""CMMF"",""Market"",""Log.Distrib.Centre"" ,mla " &
            '   " union all" &
            '   " select  to_char(txdate,'YYYYMM') as ""Calendar year/Month"",tx.cmmf as ""CMMF"",'M037HK'::text as ""Market"" ,case when sales.get_producttype(productlinegps,brand) = 'CKW - Lago' then 'C37H' " &
            '   " when sales.get_producttype(productlinegps,brand) = 'CKW - Lago' then 'C37H' Else '3730' end as  ""Log.Distrib.Centre"" ,null::text as ""Main Local Account"",groupname as ""Forecast Group"",null::text as ""Sales Force""," &
            '   " null::text as ""Sales Organization"",null::text as ""Sales Group"",sum(salesforecast) as ""Commercial Qty"" from sales.sfcmmf c left join sales.sfgrouptxhk tx on c.cmmf = tx.cmmf where tx.txdate >= '{1:yyyy-MM}-01' and tx.txdate < '{2}' group by ""Calendar year/Month"",""CMMF"",""Market"",""Log.Distrib.Centre"" ,groupname", e.startPeriod, e.startPeriod.AddMonths(6), NextYear)
            sqlstr = String.Format("(select 'NR',to_char(txdate,'YYYYMM') as ""Calendar year/Month"" ,tx.cmmf as ""CMMF"", case   when sales.get_producttype(productlinegps,brand) = 'CKW - Lago' then 'C37H' when sales.get_producttype(productlinegps,brand) = 'CKW - Tefal' and  tx.mla in ('W9000335','W9000341') then 'C37H' Else '3730' end as  ""Log.Distrib.Centre"", tx.mla as ""Main Local Account"",null,null,null,null,sum(salesforecast) as ""Commercial Qty"" ,'PC'" &
                                   " from sales.sfcmmf c inner join sales.sfmlatxhk tx on tx.cmmf = c.cmmf inner join sales.sfcmmfkamassignment cka on cka.id = tx.cmmfkamassignmentid inner join sales.sfkamassignment ka on ka.id = cka.kamassignmentid inner join sales.sfmlacardname mc on mc.id = ka.mlacardnameid  where tx.txdate >= '{0:yyyy-MM}-01' and tx.txdate < '{1:yyyy-MM}-01' " &
                                   " group by ""Calendar year/Month"",""CMMF"",""Log.Distrib.Centre"" ,tx.mla order by ""Calendar year/Month"",""CMMF"",""Main Local Account"")" &
                                   " union all " &
                                   " select 'NR',to_char(txdate,'YYYYMM') as ""Calendar year/Month"" ,tx.cmmf as ""CMMF"",case when sales.get_producttype(productlinegps,brand) = 'CKW - Lago' then 'C37H'  when sales.get_producttype(productlinegps,brand) = 'CKW - Lago' then 'C37H' Else '3730' end as  ""Log.Distrib.Centre"", 'W9000343' as ""Main Local Account"",null,null,null,null,sum(salesforecast) as ""Commercial Qty"" ,'PC' " &
                                   " from sales.sfcmmf c inner join sales.sfgrouptxhk tx on c.cmmf = tx.cmmf inner join sales._user u on u.id = tx.userid where tx.txdate >= '{1:yyyy-MM}-01' and tx.txdate < '{2}' group by ""Calendar year/Month"",""CMMF"",""Log.Distrib.Centre"" ,groupname order by ""Calendar year/Month"",""CMMF""", e.startPeriod, e.startPeriod.AddMonths(6), NextYear)

            Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplateHK.xltx", "A2", False)
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
Public Class APOHKEventArgs
    Inherits EventArgs

    Public Property startPeriod As Date

End Class