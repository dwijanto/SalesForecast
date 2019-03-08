Imports System.Text
Imports Microsoft.Office.Interop
Public Class ForecastGroupTemplateTH
    Dim myPeriodRange As Dictionary(Of Integer, Date)

    Dim myKAMGroup As New KAMGroupTHController
    Dim myKAMGroupList As List(Of KAMGroupTHModel)

    Dim THReportProperty1 As THReportProperty = THReportProperty.getInstance

    'Dim KamParamController As MSParamController
    'Dim KamParamList As List(Of MSParamModel)
    'Dim KamBudgetNettController As MSBudgetController
    'Dim KamBudgetList As List(Of MSBudgetModel)

    'Dim SDABudgetDict As Dictionary(Of Date, Decimal)
    'Dim CWBudgetDict As Dictionary(Of Date, Decimal)
    'Private GroupSDABudgetDict As Dictionary(Of Integer, Object)
    'Private GroupCWBudgetDict As Dictionary(Of Integer, Object)

    'Private exRate As Decimal
    Dim fieldList As StringBuilder


    Public Sub Generate(myForm As Object, e As FGTemplateTHEventArgs)

        Dim sqlstr As String = String.Empty

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("ForecastGroupTH{0}{1:yyyyMMdd}.xlsx", e.userName, Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 2

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            myKAMGroupList = New List(Of KAMGroupTHModel)
            myKAMGroupList = myKAMGroup.getAssignments(e.userName)


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

            Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\TWTemplate.xltx", THReportProperty1.RowStartdData)
            myreport.QueryList = New List(Of QueryWorksheet)
            'GroupGrossSalesdict = New Dictionary(Of Integer, Object)
            'GroupGrossSalesTargetdict = New Dictionary(Of Integer, Object)
            'GroupSDABudgetDict = New Dictionary(Of Integer, Object)
            'GroupCWBudgetDict = New Dictionary(Of Integer, Object)
            'exRate = e.ExRate
            For i = 0 To myKAMGroupList.Count - 1
                'GrossSalesTWController1 = New GrossSalesTWController
                'GrossSalesTargetTWController1 = New GrossSalesTargetTWController
                'Dim GrossSalesList = GrossSalesTWController1.Model.PopulateGrossSales(e.startPeriod, myKAMGroupList(i).kam, myKAMGroupList(i).groupname)
                'Dim GrossSalesTargetList = GrossSalesTargetTWController1.Model.PopulateGrossSalesTarget(e.startPeriod, myKAMGroupList(i).kam, myKAMGroupList(i).groupname)
                'GrossSalesDict = New Dictionary(Of Date, Decimal)
                'GrossSalesTargetDict = New Dictionary(Of Date, Decimal)

                'Get SDA BudgetNet 
                'SDABudgetDict = New Dictionary(Of Date, Decimal)
                'CWBudgetDict = New Dictionary(Of Date, Decimal)
                'Dim MSBudgetController1 = New MSBudgetController
                'Dim SDABudgetList = MSBudgetController1.Model.PopulateBudgetList(e.startPeriod, myKAMGroupList(i).kam, myKAMGroupList(i).groupname, ProductTypeEnum.SDA)
                'Dim CWBudgetList = MSBudgetController1.Model.PopulateBudgetList(e.startPeriod, myKAMGroupList(i).kam, myKAMGroupList(i).groupname, ProductTypeEnum.CKW)

                'For Each obj As MSBudgetModel In SDABudgetList
                ' SDABudgetDict.Add(CDate(String.Format("{0:yyyy-MM}-01", obj.period)), obj.budgetnett)
                'Next

                'For Each obj As MSBudgetModel In CWBudgetList
                ' CWBudgetDict.Add(CDate(String.Format("{0:yyyy-MM}-01", obj.period)), obj.budgetnett)
                'Next

                'GroupSDABudgetDict.Add(i, SDABudgetDict)
                'GroupCWBudgetDict.Add(i, CWBudgetDict)

                'For Each obj As GrossSalesTWModel In GrossSalesList
                ' GrossSalesDict.Add(CDate(String.Format("{0:yyyy-MM}-01", obj.period)), obj.amount)
                'Next
                'For Each obj As GrossSalesTargetTWModel In GrossSalesTargetList
                '    GrossSalesTargetDict.Add(CDate(String.Format("{0:yyyy-MM}-01", obj.period)), obj.amount)
                'Next
                'GroupGrossSalesdict.Add(i, GrossSalesDict)
                'GroupGrossSalesTargetdict.Add(i, GrossSalesTargetDict)
                Dim myQuery = New QueryWorksheet
                Dim username = e.userName
                If e.blanktemplate Then
                    'sqlstr = String.Format("with tx as (select * from crosstab('select c.cmmf,null::date as txdate,null::integer as salesforecast from sales.sfcmmftw c " &
                    '                       " order by c.producttype,c.cmmf','select unnest(Array[{1}])::date') " &
                    '                   " as ct (cmmf bigint, {2})) ,nsp as (select cmmf,nsp1, nsp2 from sales.sfcmmfnsptw )" &
                    '                   " select  c.producttype,c.cmmf,c.localcmmf,c.chinesedesc,c.description,c.productrange,c.referenceno,c.launchdate,c.remarks,n.nsp1::numeric(13,0) as ""NSP(TW)"",n.nsp1 / {4} as ""NSP(USD)"" ,c.moq,{3} from sales.sfcmmftw c " &
                    '                   " left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf order by c.producttype,c.subproducttype,c.cmmf;", myPeriodRange(0), fieldList.ToString, ColumnList.ToString, TxtFieldList.ToString, exRate)


                Else
                    'sqlstr = String.Format("with c as (select cmmf,case when productline like '%COOKWARE%' then 4 else 1 end as producttype from sales.sfcmmfms), tx as (select * from crosstab('with c as (select cmmf,case when productline like ''%COOKWARE%'' then 4 else 1 end as producttype from sales.sfcmmfms)" &
                    '                       " select c.cmmf,tx.txdate,tx.salesforecast from sales.sfgrouptxms tx " &
                    '                       " left join  c on c.cmmf = tx.cmmf left join sales.sfgroup g on g.id = tx.groupid  where tx.kam = ''{0}'' and g.groupname = ''{1}'' and not c.cmmf isnull" &
                    '                       " order by c.cmmf','select unnest(Array[{2}])::date') " &
                    '                   " as ct (cmmf bigint, {3})) ,sd as (select p.groupid,p.producttype,kam,frontmargin,ifrsrebate from sales.sfmsparam p" &
                    '                   " left join sales.sfgroup g on g.id = p.groupid where g.groupname = '{1}' and  p.kam = '{0}') " &
                    '                   " select  sc.productline,sc.familylvl2,sc.cmmf,sc.description ,sc.discount,sc.launch,null as frontmargin,null as ifrsrebate, sc.rsp as nnsp,null as gross, sc.rsp as nnsp, {4} from  sales.sfcmmfms sc " &
                    '                   " left join c on c.cmmf = sc.cmmf left join tx on tx.cmmf = c.cmmf left join sd on sd.producttype = c.producttype order by sc.productline,sc.familylvl2,sc.cmmf;", e.userName, myKAMGroupList(i).groupname, fieldList.ToString, ColumnList.ToString, TxtFieldList.ToString, exRate)
                    'sqlstr = String.Format("with tx as (select * from crosstab('select tx.cmmf,tx.txdate,tx.salesforecast from sales.sfgrouptxth tx  where tx.kamassignmentid = {0}  order by tx.cmmf','select unnest(Array[{1}])::date') as ct (cmmf bigint, {2}))" &
                    '         " select c.cmmf,c.productline,c.familylv2,c.commercialcode,c.itemdescription,c.status,c.rsp,c.cogs,null as nnsp,{3} from sales.sfcmmfth c left join tx on tx.cmmf = c.cmmf order by productline,familylv2,commercialcode", myKAMGroupList(i).assigmentid, fieldList.ToString, ColumnList.ToString, TxtFieldList.ToString)
                    sqlstr = String.Format("with tx as (select * from crosstab('select tx.cmmf,tx.txdate,tx.salesforecast from sales.sfgrouptxth tx  where tx.kamassignmentid = {0}  order by tx.cmmf','select unnest(Array[{1}])::date') as ct (cmmf bigint, {2}))" &
                             " select c.cmmf,c.productline,c.familylv2,c.commercialcode,c.itemdescription,c.status,c.rsp,c.cogs,case when productline like '%COOKWARE' then  c.rsp * (1 - gpcw) / 1.07 * (1-ifrscw) else c.rsp * (1 - gpsda) / 1.07 * (1-ifrssda) " &
                             " end as nnsp,{3} from sales.sfcmmfth c left join sales.sfkamassignmentth ka on ka.id = {0}" &
                             " left join sales.sfmlacardnameth mc on mc.id = ka.mlacardnameid" &
                             " left join sales.sfmlagroupth mg on mg.mla = mc.mla left join tx on tx.cmmf = c.cmmf order by productline,familylv2,commercialcode", myKAMGroupList(i).assigmentid, fieldList.ToString, ColumnList.ToString, TxtFieldList.ToString)
                End If
                myQuery.Sqlstr = sqlstr
                myQuery.DataSheet = i + datasheet
                myQuery.SheetName = String.Format("TH-{0}-{1}", myKAMGroupList(i).assigmentid, myKAMGroupList(i).assignment)
                myreport.QueryList.Add(myQuery)
            Next
            myreport.Run(myForm, e)
        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As ExportToExcelFileEventArgs)
        Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        Dim DataStart As Integer = THReportProperty1.ColumnStartData
        osheet.Cells(1, 2).value = "Trade Discount from RSP"
        osheet.Cells(1, 3).value = "SDA"
        osheet.Cells(1, 4).value = "CW"
        osheet.Cells(2, 2).value = "GP%"
        osheet.Cells(3, 2).value = "IFRS%"

        osheet.Range("B1:B3").Font.Bold = True
        osheet.Range("C1:D1").Font.Bold = True
        osheet.Cells(2, 3).value = myKAMGroupList(e.SheetNo - 2).gpsda
        osheet.Cells(3, 3).value = myKAMGroupList(e.SheetNo - 2).ifrssda
        osheet.Cells(2, 4).value = myKAMGroupList(e.SheetNo - 2).gpcw
        osheet.Cells(3, 4).value = myKAMGroupList(e.SheetNo - 2).ifrscw
        osheet.Range("C2:D3").NumberFormat = "0.0%"

        osheet.Cells(1, 9).value = "SDA"
        osheet.Cells(2, 9).value = "CW"
        osheet.Cells(3, 9).value = "Total"
        osheet.Range("I1:I3").Font.Bold = True

        Dim ColumnStartData = THReportProperty1.ColumnStartData
        'For i = 0 To 11
        For i = 0 To 12
            osheet.Cells(1, ColumnStartData + i).FormulaR1C1 = String.Format("=SUMPRODUCT(SIGN(R[6]C[{1}]:R[{0}]C[{1}]<>""5 - COOKWARE""),R[6]C[{2}]:R[{0}]C[{2}],R[6]C:R[{0}]C)/1000", e.lastRow - 1, -(ColumnStartData - 2) - i, -1 - i)
            osheet.Cells(2, ColumnStartData + i).FormulaR1C1 = String.Format("=SUMPRODUCT(SIGN(R[5]C[{1}]:R[{0}]C[{1}]=""5 - COOKWARE""),R[5]C[{2}]:R[{0}]C[{2}],R[5]C:R[{0}]C)/1000", e.lastRow - 2, -(ColumnStartData - 2) - i, -1 - i)
            osheet.Cells(3, ColumnStartData + i).formulaR1C1 = String.Format("=R[-2]C[0]+R[-1]C[0]")
        Next
        osheet.Columns("G:I").numberformat = "#,##0.00"
        osheet.Columns("J:V").numberformat = "0"
        osheet.Range("J1:V5").Style = "Comma"
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

        osheet.Cells(3, 1) = "SDA"
        osheet.Cells(4, 1) = "CW"
        osheet.Cells(5, 1) = "Total"
        'For i = 0 To 11
        For i = 0 To 12
            '    Dim mysdabsb As New StringBuilder
            Dim mysdafsb As New StringBuilder
            '    Dim mycwbsb As New StringBuilder
            Dim mycwfsb As New StringBuilder

            For j = 0 To myKAMGroupList.Count - 1
                If mysdafsb.Length > 0 Then
                    mysdafsb.Append("+")
                    mycwfsb.Append("+")
                End If
                mysdafsb.Append(String.Format("'TH-{0}-{1}'!R[-2]C[{2}]", myKAMGroupList(j).assigmentid, myKAMGroupList(j).assignment, 8))
                mycwfsb.Append(String.Format("'TH-{0}-{1}'!R[-2]C[{2}]", myKAMGroupList(j).assigmentid, myKAMGroupList(j).assignment, 8))
            Next

            osheet.Cells(3, 2 + i).formulaR1C1 = String.Format("={0}", mysdafsb.ToString)
            osheet.Cells(4, 2 + i).formulaR1C1 = String.Format("={0}", mycwfsb.ToString)
            osheet.Cells(5, 2 + i).FormulaR1C1 = String.Format("=R[-2]C[0]+R[-1]C[0]")
        Next



        'For i = 0 To 11
        '    Dim myresult As Integer = 0
        '    Try
        '        'myresult = KamTargetDict(myPeriodRange(i))
        '    Catch ex As Exception

        '    End Try
        '    'osheet.Cells(5, 2 + i).value = "" & myresult.ToString
        '    'osheet.Cells(6, 2 + i).FormulaR1C1 = String.Format("=R[-2]C-R[-1]C")
        '    'osheet.Cells(8, 2 + i).FormulaR1C1 = String.Format("=R[-4]C-R[-1]C")
        '    'this cell need to divided by exrate // Don't forget the all KAM report for exrate
        '    'osheet.Cells(9, 2 + i).FormulaR1C1 = String.Format("=R[-6]C/{0}", exRate)
        'Next
        osheet.Range("B3:N5").Style = "Comma"
        osheet.Range("A2:A5").Font.Bold = True
        osheet.Range("A2:N2").Font.Bold = True
        osheet.Name = "TH-Summary"

        'osheet.Cells.EntireColumn.AutoFit()
    End Sub

End Class
