Imports Microsoft.Office.Interop
Public Class CMMFMYExport
    Public Sub Generate(myForm As Object, e As EventArgs)

        Dim sqlstr As String = String.Empty
        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("CMMFMY{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            sqlstr = String.Format("select cmmf,productline,familylvl2,description,discount,launch,rsp from sales.sfcmmfms order by cmmf")
            Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx", "A1")
            myreport.Run(myForm, e)

        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
    End Sub
End Class
