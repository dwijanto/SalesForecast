Imports Microsoft.Office.Interop
Public Class ALLKAMTW
    Dim myPeriodRange As Dictionary(Of Integer, Date)
    Dim myKAMAssignment As New KAMAssignmentController
    Dim myKAMAssignmentList As List(Of KAMAssignmentModel)
    Dim TWReportProperty1 As TWReportProperty = TWReportProperty.getInstance
    Dim TBParamDetailController1 As New TBParamDetailController

    Public Sub Generate(myForm As Object, e As ALLKAMEventArgs)

        Dim sqlstr As String = String.Empty
        Dim exRate = TBParamDetailController1.Model.getCurrency(country.TW, "EX-Rate")
        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("SalesForecastALL{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 2

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            sqlstr = String.Format("(select c.*,tx.txdate,tx.kam,tx.groupname,tx.salesforecast,nsp.nsp1,nsp.nsp1/{0} as nsp2,tx.salesforecast * nsp.nsp1 as grosssalestw,tx.salesforecast * (nsp.nsp1/{0}) as grosssalesusd from sales.sfgrouptxtw tx" &
                     " left join sales.sfcmmfnsptw nsp on nsp.cmmf = tx.cmmf" &
                     " left join sales.sfcmmftw c on c.cmmf = tx.cmmf {1})", exRate, String.Format(" where tx.txdate >= '{0:yyyy-MM-}01' and tx.txdate <= '{1:yyyy-MM-}01'", e.startperiod, e.endperiod))


            Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\TWTemplate.xltx", TWReportProperty1.RowStartdData)
            myreport.Run(myForm, e)

        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        Dim DataStart As Integer = TWReportProperty1.ColumnStartData
        osheet.Range("P:S").NumberFormat = "#,##0.00"

    End Sub

    Private Sub PivotTable()
        'Throw New NotImplementedException
    End Sub
End Class
