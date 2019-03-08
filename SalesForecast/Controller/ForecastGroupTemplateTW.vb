Imports System.Text
Imports Microsoft.Office.Interop

Public Class ForecastGroupTemplateTW
    Dim myPeriodRange As Dictionary(Of Integer, Date)
    Dim myKAMAGroup As New KamGroupTWController
    Dim myKAMAGroupList As List(Of KAMGroupTWModel)    
    Dim TWReportProperty1 As TWReportProperty = TWReportProperty.getInstance
    Dim KamTargetController As TWParamController

    Dim KamTargetList As List(Of TWParamModel)
    Private KamTargetDict As Dictionary(Of Date, Integer)
    Dim GrossSalesTWController1 As GrossSalesTWController
    Dim GrossSalesTargetTWController1 As GrossSalesTargetTWController

    Private GrossSalesDict As Dictionary(Of Date, Decimal)
    Private GrossSalesDictList As List(Of Object)
    Private GroupGrossSalesdict As Dictionary(Of Integer, Object)

    Private GrossSalesTargetDict As Dictionary(Of Date, Decimal)
    Private GrossSalesTargetDictList As List(Of Object)
    Private GroupGrossSalesTargetdict As Dictionary(Of Integer, Object)
    Private exRate As Decimal
    Dim fieldList As StringBuilder


    Public Sub Generate(myForm As Object, e As FGTemplateTWEventArgs)

        Dim sqlstr As String = String.Empty

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("ForecastGroupTW{0}{1:yyyyMMdd}.xlsx", e.userName, Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 2

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            myKAMAGroupList = New List(Of KAMGroupTWModel)
            myKAMAGroupList = myKAMAGroup.getAssignments(e.userName)

            KamTargetController = New TWParamController
            KamTargetList = KamTargetController.Model.PopulateKAMTarget(e.userName)
            KamTargetDict = New Dictionary(Of Date, Integer)
            For Each obj As TWParamModel In KamTargetList
                KamTargetDict.Add(obj.period, obj.targetgross)
            Next



            'Dim PeriodRange = New PeriodRange(e.startPeriod, 12)
            Dim PeriodRange = New PeriodRange(e.startPeriod, 13)
            myPeriodRange = PeriodRange.getPeriod
            fieldList = New StringBuilder
            Dim ColumnList As New StringBuilder
            Dim TxtFieldList As New StringBuilder
            Dim DateList As New List(Of Date)
            For i = 0 To myPeriodRange.Count - 1
                If fieldList.Length > 0 Then
                    fieldList.Append(",")
                    ColumnList.Append(",")
                    TxtFieldList.Append(",")
                End If
                fieldList.Append(String.Format("''{0:yyyy-MM-dd}''", myPeriodRange(i)))
                ColumnList.Append(String.Format("""{0:yyyyMM}"" bigint", myPeriodRange(i)))
                TxtFieldList.Append(String.Format("tx.""{0:yyyyMM}""", myPeriodRange(i)))
            Next

            Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\TWTemplate.xltx", TWReportProperty1.RowStartdData)
            myreport.QueryList = New List(Of QueryWorksheet)
            GroupGrossSalesdict = New Dictionary(Of Integer, Object)
            GroupGrossSalesTargetdict = New Dictionary(Of Integer, Object)
            exRate = e.ExRate
            For i = 0 To myKAMAGroupList.Count - 1
                GrossSalesTWController1 = New GrossSalesTWController
                GrossSalesTargetTWController1 = New GrossSalesTargetTWController
                Dim GrossSalesList = GrossSalesTWController1.Model.PopulateGrossSales(e.startPeriod, myKAMAGroupList(i).kam, myKAMAGroupList(i).groupname)
                Dim GrossSalesTargetList = GrossSalesTargetTWController1.Model.PopulateGrossSalesTarget(e.startPeriod, myKAMAGroupList(i).kam, myKAMAGroupList(i).groupname)
                GrossSalesDict = New Dictionary(Of Date, Decimal)
                GrossSalesTargetDict = New Dictionary(Of Date, Decimal)

                For Each obj As GrossSalesTWModel In GrossSalesList
                    GrossSalesDict.Add(CDate(String.Format("{0:yyyy-MM}-01", obj.period)), obj.amount)
                Next
                For Each obj As GrossSalesTargetTWModel In GrossSalesTargetList
                    GrossSalesTargetDict.Add(CDate(String.Format("{0:yyyy-MM}-01", obj.period)), obj.amount)
                Next
                GroupGrossSalesdict.Add(i, GrossSalesDict)
                GroupGrossSalesTargetdict.Add(i, GrossSalesTargetDict)
                Dim myQuery = New QueryWorksheet
                Dim username = e.userName
                If e.blanktemplate Then
                    sqlstr = String.Format("with tx as (select * from crosstab('select c.cmmf,null::date as txdate,null::integer as salesforecast from sales.sfcmmftw c " &
                                           " order by c.producttype,c.cmmf','select unnest(Array[{1}])::date') " &
                                       " as ct (cmmf bigint, {2})) ,nsp as (select cmmf,nsp1, nsp2 from sales.sfcmmfnsptw )" &
                                       " select  c.producttype,c.cmmf,c.localcmmf,c.chinesedesc,c.description,c.productrange,c.referenceno,c.launchdate,c.remarks,n.nsp1::numeric(13,0) as ""NSP(TW)"",n.nsp1 / {4} as ""NSP(USD)"" ,c.moq,{3} from sales.sfcmmftw c " &
                                       " left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf order by c.producttype,c.subproducttype,c.cmmf;", myPeriodRange(0), fieldList.ToString, ColumnList.ToString, TxtFieldList.ToString, exRate)


                Else
                    sqlstr = String.Format("with tx as (select * from crosstab('select c.cmmf,tx.txdate,tx.salesforecast from sales.sfgrouptxtw tx " &
                                           " left join sales.sfcmmftw c on c.cmmf = tx.cmmf where tx.kam = ''{0}'' and tx.groupname = ''{1}''" &
                                           " order by c.producttype,c.cmmf','select unnest(Array[{2}])::date') " &
                                       " as ct (cmmf bigint, {3})) ,nsp as (select cmmf,nsp1, nsp2 from sales.sfcmmfnsptw )" &
                                       " select  c.producttype,c.cmmf,c.localcmmf,c.chinesedesc,c.description,c.productrange,c.referenceno,c.launchdate,c.remarks,n.nsp1::numeric(13,0) as ""NSP(TW)"",n.nsp1 / {5} as ""NSP(USD)"" ,c.moq,{4} from sales.sfcmmftw c " &
                                       " left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf order by c.producttype,c.subproducttype,c.cmmf;", e.userName, myKAMAGroupList(i).groupname, fieldList.ToString, ColumnList.ToString, TxtFieldList.ToString, exRate)
                End If
                myQuery.Sqlstr = sqlstr
                myQuery.DataSheet = i + datasheet
                myQuery.SheetName = String.Format("TW-{0}-{1}", username, myKAMAGroupList(i).groupname) '"2016.09"
                myreport.QueryList.Add(myQuery)
            Next
            myreport.Run(myForm, e)
        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As ExportToExcelFileEventArgs)
        Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        Dim DataStart As Integer = TWReportProperty1.ColumnStartData
        osheet.Cells(2, 6).value = "Sales Deduction"
        osheet.Cells(2, 7).value = myKAMAGroupList(e.SheetNo - 2).sd
        osheet.Cells(2, 7).numberformat = "0%"

        osheet.Cells(2, 9).value = "Summary (Net Sales)"
        osheet.Cells(3, 9).value = "Summary (Gross Sales)"
        osheet.Cells(4, 9).value = "Gross sales Target(Y)"
        osheet.Cells(5, 9).value = "Gross sales (Y-1)"
        osheet.Cells(7, 9).value = "TEFAL"
        osheet.Cells(8, 9).value = "LAGOSTINA"
        osheet.Cells(9, 9).value = "RTM LP"
        osheet.Cells(10, 9).value = "Ex Rate"
        osheet.Cells(10, 10).value = exRate
        Dim ColumnStartData = TWReportProperty1.ColumnStartData
        'For i = 0 To 11
        For i = 0 To 12
            osheet.Cells(7, ColumnStartData + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R[5]C[{1}]:R[{0}]C[{1}]=""TEFAL""),R[5]C[{2}]:R[{0}]C[{2}],R[5]C:R[{0}]C)/1000", e.lastRow - 5, -(ColumnStartData - 1) - i, -3 - i)
            osheet.Cells(8, ColumnStartData + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R[4]C[{1}]:R[{0}]C[{1}]=""LAGOSTINA""),R[4]C[{2}]:R[{0}]C[{2}],R[4]C:R[{0}]C)/1000", e.lastRow - 6, -(ColumnStartData - 1) - i, -3 - i)
            osheet.Cells(9, ColumnStartData + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R[3]C[{1}]:R[{0}]C[{1}]=""RTM LP""),R[3]C[{2}]:R[{0}]C[{2}],R[3]C:R[{0}]C)/1000", e.lastRow - 7, -(ColumnStartData - 1) - i, -3 - i)
            osheet.Cells(2, ColumnStartData + i).FormulaR1C1 = String.Format("=SUM(R[5]C:R[7]C)")
            osheet.Cells(3, ColumnStartData + i).formulaR1C1 = String.Format("=R[-1]C/(1-R[-1]C[{0}])", -6 - i)
            Dim mydict As Dictionary(Of Date, Decimal) = GroupGrossSalesdict(e.SheetNo - 2)
            Dim mykey = CDate(String.Format("{0}-{1:MM}-01", myPeriodRange(i).Year - 1, myPeriodRange(i)))
            If (mydict.ContainsKey(mykey)) Then
                osheet.Cells(5, ColumnStartData + i).value = mydict(mykey)
            End If

            Dim mydictT As Dictionary(Of Date, Decimal) = GroupGrossSalesTargetdict(e.SheetNo - 2)
            Dim mykeyT = CDate(String.Format("{0:yyyy-MM}-01", myPeriodRange(i)))
            If (mydictT.ContainsKey(mykeyT)) Then
                osheet.Cells(4, ColumnStartData + i).value = mydictT(mykeyT)
            End If
        Next
        osheet.Columns("K:K").numberformat = "#,##0.00"
        osheet.Range("M2:Y9").Style = "Comma"

        osheet.Cells.EntireColumn.AutoFit()
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        Dim owb As Excel.Workbook = DirectCast(sender, Excel.Workbook)
        owb.Worksheets(1).select()

        Dim osheet As Excel.Worksheet = owb.Worksheets(1)

        osheet.Cells(2, 1).value = "Period"
        'For i = 0 To 11
        For i = 0 To 12
            osheet.Cells(2, 2 + i).value = String.Format("{0:yyyyMM}", myPeriodRange(i))
        Next

        osheet.Cells(3, 1) = "Summary Net Sales"
        'For i = 0 To 11
        For i = 0 To 12
            Dim mysb As New StringBuilder
            Dim mygssb As New StringBuilder
            Dim mygstsb As New StringBuilder
            For j = 0 To myKAMAGroupList.Count - 1
                If mysb.Length > 0 Then
                    mysb.Append("+")
                    mygssb.Append("+")
                    mygstsb.Append("+")
                End If
                mysb.Append(String.Format("'TW-{0}-{1}'!R[-1]C[{2}]", myKAMAGroupList(j).kam, myKAMAGroupList(j).groupname, 11))
                mygssb.Append(String.Format("'TW-{0}-{1}'!R[-2]C[{2}]", myKAMAGroupList(j).kam, myKAMAGroupList(j).groupname, 11))
                mygstsb.Append(String.Format("'TW-{0}-{1}'!R[-1]C[{2}]", myKAMAGroupList(j).kam, myKAMAGroupList(j).groupname, 11))
            Next
            osheet.Cells(3, 2 + i).formulaR1C1 = String.Format("={0}", mysb.ToString)
            osheet.Cells(4, 2 + i).formulaR1C1 = String.Format("={0}", mysb.ToString)
            osheet.Cells(5, 2 + i).formulaR1C1 = String.Format("={0}", mygstsb.ToString)
            osheet.Cells(7, 2 + i).formulaR1C1 = String.Format("={0}", mygssb.ToString)
        Next

        osheet.Cells(4, 1) = "Summary Gross Sales"

        osheet.Cells(5, 1) = "Summary Target Sales"
        osheet.Cells(6, 1) = "Summary Difference"
        osheet.Cells(7, 1) = "Gross sales(Y-1)"
        'osheet.Cells(8, 1) = "Difference (Y vs Y-1)"
        osheet.Cells(8, 1) = "Gross Sales Y vs Gross Sales(Y-1)"
        osheet.Cells(9, 1) = "Summary Net Sales (USD)"

        'For i = 0 To 11
        For i = 0 To 12
            Dim myresult As Integer = 0
            Try
                myresult = KamTargetDict(myPeriodRange(i))
            Catch ex As Exception

            End Try
            'osheet.Cells(5, 2 + i).value = "" & myresult.ToString
            osheet.Cells(6, 2 + i).FormulaR1C1 = String.Format("=R[-2]C-R[-1]C")
            osheet.Cells(8, 2 + i).FormulaR1C1 = String.Format("=R[-4]C-R[-1]C")
            'this cell need to divided by exrate // Don't forget the all KAM report for exrate
            osheet.Cells(9, 2 + i).FormulaR1C1 = String.Format("=R[-6]C/{0}", exRate)
        Next
        osheet.Range("B3:N9").Style = "Comma"
        osheet.Name = "TW-Summary"

        osheet.Cells.EntireColumn.AutoFit()
    End Sub
End Class
