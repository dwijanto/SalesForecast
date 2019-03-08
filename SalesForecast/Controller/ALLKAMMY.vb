Imports Microsoft.Office.Interop
Public Class ALLKAMMY
    Dim myPeriodRange As Dictionary(Of Integer, Date)
    ' Dim myKAMAssignment As New KAMAssignmentController
    ' Dim myKAMAssignmentList As List(Of KAMAssignmentModel)
    Dim MSReportProperty1 As MSReportProperty = MSReportProperty.getInstance
    'Dim TBParamDetailController1 As New TBParamDetailController

    Public Sub Generate(myForm As Object, e As ALLKAMEventArgs)

        Dim sqlstr As String = String.Empty
        ' Dim exRate = TBParamDetailController1.Model.getCurrency(country.TW, "EX-Rate")
        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("SalesForecastMYALL{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 2

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            'sqlstr = String.Format("(select c.*,tx.txdate,tx.kam,tx.groupname,tx.salesforecast,nsp.nsp1,nsp.nsp1/1 as nsp2,tx.salesforecast * nsp.nsp1 as grosssalestw,tx.salesforecast * (nsp.nsp1/1) as grosssalesusd from sales.sfgrouptxtw tx" &
            '         " left join sales.sfcmmfnsptw nsp on nsp.cmmf = tx.cmmf" &
            '         " left join sales.sfcmmftw c on c.cmmf = tx.cmmf {0})", String.Format(" where tx.txdate >= '{0:yyyy-MM-}01' and tx.txdate <= '{1:yyyy-MM-}01'", e.startperiod, e.endperiod))
            sqlstr = String.Format("with cc as (select cmmf, case when productline like '%COOKWARE%' then 1 else 4 end as producttype from sales.sfcmmfms) select g.groupname,gm.mla,tx.kam,c.productline,c.familylvl2,tx.cmmf,c.description,c.discount,c.launch,p.frontmargin,p.ifrsrebate,c.rsp,(c.rsp * (1 - p.frontmargin)/1.06 ) as gross,(c.rsp * (1 - p.frontmargin)/1.06 ) * (1-p.ifrsrebate) as net,tx.txdate,tx.salesforecast,(c.rsp * (1 - p.frontmargin)/1.06 ) * (1-p.ifrsrebate) * tx.salesforecast as amount" &
                     " from sales.sfgrouptxms tx left join sales.sfcmmfms c on c.cmmf = tx.cmmf left join cc on cc.cmmf = tx.cmmf left join sales.sfmsparam p on p.kam = tx.kam and p.groupid = tx.groupid and p.producttype = cc.producttype" &
                     " left join sales.sfgroupmlams gm on gm.groupid = tx.groupid" &
                     " left join sales.sfgroup g on g.id = tx.groupid {0}", String.Format(" where tx.txdate >= '{0:yyyy-MM-}01' and tx.txdate <= '{1:yyyy-MM-}01'", e.startperiod, e.endperiod))

            Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\TWTemplate.xltx", MSReportProperty1.RowStartdData)
            myreport.Run(myForm, e)

        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        Dim DataStart As Integer = MSReportProperty1.ColumnStartData
        'osheet.Range("P:S").NumberFormat = "#,##0.00"
        osheet.Name = "RAWDATA"
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        Dim owb As Excel.Workbook = DirectCast(sender, Excel.Workbook)
        Dim oxl = owb.Parent

        Dim osheet As Excel.Worksheet
        owb.Worksheets(1).select()
        osheet = owb.Worksheets(1)

        owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "RAWDATA!ExternalData_1").CreatePivotTable(osheet.Name & "!R7C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.PivotTables("PivotTable1").Pivotfields("productline").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("familylvl2").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("cmmf").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("description").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("cmmf").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("salesforecast"), "Sum of salesforecast", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("amount"), "Sum of amount", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").Pivotfields("txdate").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").Pivotfields("txdate").numberformat = "mmm-yy"
        osheet.PivotTables("PivotTable1").Pivotfields("Sum of salesforecast").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
        osheet.PivotTables("PivotTable1").Pivotfields("Sum of amount").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
        osheet.Cells.EntireColumn.AutoFit()
    End Sub
End Class
