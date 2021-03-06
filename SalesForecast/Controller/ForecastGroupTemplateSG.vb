﻿Imports System.Text
Imports Microsoft.Office.Interop

Public Class ForecastGroupTemplateSG
    Dim myPeriodRange As Dictionary(Of Integer, Date)

    Dim myKAMGroup As New KAMGroupSGController
    Dim myKAMGroupList As List(Of KAMGroupSGModel)

    Dim SGReportProperty1 As SGReportProperty = SGReportProperty.getInstance

    'Dim KamParamController As MSParamController
    'Dim KamParamList As List(Of MSParamModel)
    'Dim KamBudgetNettController As MSBudgetController
    'Dim KamBudgetList As List(Of MSBudgetModel)

    'Private KamTargetDict As Dictionary(Of Date, Integer)
    'Dim GrossSalesTWController1 As GrossSalesTWController
    'Dim GrossSalesTargetTWController1 As GrossSalesTargetTWController

    'Private GrossSalesDict As Dictionary(Of Date, Decimal)
    'Private GrossSalesDictList As List(Of Object)
    'Private GroupGrossSalesdict As Dictionary(Of Integer, Object)

    'Private GrossSalesTargetDict As Dictionary(Of Date, Decimal)
    'Private GrossSalesTargetDictList As List(Of Object)
    'Private GroupGrossSalesTargetdict As Dictionary(Of Integer, Object)

    'Dim SDABudgetDict As Dictionary(Of Date, Decimal)
    'Dim CWBudgetDict As Dictionary(Of Date, Decimal)
    'Private GroupSDABudgetDict As Dictionary(Of Integer, Object)
    'Private GroupCWBudgetDict As Dictionary(Of Integer, Object)

    Private exRate As Decimal
    Dim fieldList As StringBuilder


    Public Sub Generate(myForm As Object, e As FGTemplateSGEventArgs)

        Dim sqlstr As String = String.Empty

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("ForecastGroupSG{0}{1:yyyyMMdd}.xlsx", e.userName, Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 2

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            myKAMGroupList = New List(Of KAMGroupSGModel)
            myKAMGroupList = myKAMGroup.getAssignments(e.userName)

            'KamTargetController = New TWParamController
            'KamTargetList = KamTargetController.Model.PopulateKAMTarget(e.userName)
            'KamTargetDict = New Dictionary(Of Date, Integer)
            'For Each obj As TWParamModel In KamTargetList
            '    KamTargetDict.Add(obj.period, obj.targetgross)
            'Next



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

            Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\TWTemplate.xltx", SGReportProperty1.RowStartdData)
            myreport.QueryList = New List(Of QueryWorksheet)
            'GroupGrossSalesdict = New Dictionary(Of Integer, Object)
            'GroupGrossSalesTargetdict = New Dictionary(Of Integer, Object)
            'GroupSDABudgetDict = New Dictionary(Of Integer, Object)
            'GroupCWBudgetDict = New Dictionary(Of Integer, Object)
            exRate = e.ExRate
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
                '    SDABudgetDict.Add(CDate(String.Format("{0:yyyy-MM}-01", obj.period)), obj.budgetnett)
                'Next

                'For Each obj As MSBudgetModel In CWBudgetList
                '    CWBudgetDict.Add(CDate(String.Format("{0:yyyy-MM}-01", obj.period)), obj.budgetnett)
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
                    '                   " select  sc.productline,sc.familylvl2,sc.cmmf,sc.description ,sc.discount,sc.launch,sd.frontmargin,sd.ifrsrebate, sc.rsp,sc.rsp * (1-sd.frontmargin) / 1.06 as gross,(sc.rsp * (1-sd.frontmargin) / 1.06) * (1 - sd.ifrsrebate) as net, {4} from  sales.sfcmmfms sc " &
                    '                   " left join c on c.cmmf = sc.cmmf left join tx on tx.cmmf = c.cmmf left join sd on sd.producttype = c.producttype order by sc.productline,sc.familylvl2,sc.cmmf;", e.userName, myKAMGroupList(i).groupname, fieldList.ToString, ColumnList.ToString, TxtFieldList.ToString, exRate)
                    sqlstr = String.Format("with c as (select cmmf,case when productline like '%COOKWARE%' then 4 else 1 end as producttype from sales.sfcmmfsg), " &
                                            " tx as (select * from crosstab('with c as (select cmmf,case when productline like ''%COOKWARE%'' then 4 else 1 end as producttype from sales.sfcmmfsg) " &
                                            " select c.cmmf,tx.txdate,tx.salesforecast from sales.sfgrouptxsg tx  left join  c on c.cmmf = tx.cmmf left join sales.sfgroup g on g.id = tx.groupid  where tx.kam = ''{0}'' and g.id = {1} and not c.cmmf isnull order by c.cmmf'," &
                                            "'select unnest(Array[{2}])::date')  as ct (cmmf bigint, {3})) " &
                                            " select  sc.productline,sc.familylvl2,sc.cmmf,sc.brand,sc.description ,sc.status,sc.pi2status," &
                                            " sc.rsp, rsp * (1-sd.sd) * (1 - sales.gettradediscountsg(sc.brand,sc.productline,td.sda,td.rowenta,td.cw)) as net,{4} from  sales.sfcmmfsg sc  left join c on c.cmmf = sc.cmmf left join tx on tx.cmmf = c.cmmf " &
                                            " left join sales.sfgroup g on g.id = {1} " &
                                            " left join sales.sfgroupsdsg sd on sd.groupid = g.id and sd.producttype = c.producttype " &
                                            " left join sales.sfgrouptradediscsg td on td.groupid = g.id order by sc.productline,sc.familylvl2,sc.cmmf;", e.userName, myKAMGroupList(i).groupid, fieldList.ToString, ColumnList.ToString, TxtFieldList.ToString)
                End If
                myQuery.Sqlstr = sqlstr
                myQuery.DataSheet = i + datasheet
                myQuery.SheetName = String.Format("SG-{0}-{1}", username, myKAMGroupList(i).groupname) '"2016.09"
                myreport.QueryList.Add(myQuery)
            Next
            myreport.Run(myForm, e)
        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As ExportToExcelFileEventArgs)
        Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        Dim DataStart As Integer = SGReportProperty1.ColumnStartData
        osheet.Cells(1, 6).value = "Trade Discount"
        osheet.Cells(2, 6).value = "SDA"
        osheet.Cells(3, 6).value = "Rowenta"
        osheet.Cells(4, 6).value = "CW"
        osheet.Cells(5, 6).value = "Sales Deduction"
        osheet.Cells(6, 6).value = "SDA"
        osheet.Cells(7, 6).value = "CW IFRS"
        osheet.Cells(2, 7).value = myKAMGroupList(e.SheetNo - 2).TDSDA
        osheet.Cells(2, 7).numberformat = "0%"
        osheet.Cells(3, 7).value = myKAMGroupList(e.SheetNo - 2).TDRowenta
        osheet.Cells(3, 7).numberformat = "0%"
        osheet.Cells(4, 7).value = myKAMGroupList(e.SheetNo - 2).TDCW
        osheet.Cells(4, 7).numberformat = "0%"
        osheet.Cells(6, 7).value = myKAMGroupList(e.SheetNo - 2).SDSDA
        osheet.Cells(6, 7).numberformat = "0%"
        osheet.Cells(7, 7).value = myKAMGroupList(e.SheetNo - 2).SDCW
        osheet.Cells(7, 7).numberformat = "0%"


        osheet.Cells(1, 9).value = "SDA"
        osheet.Cells(2, 9).value = "CW"
        osheet.Cells(3, 9).value = "Total"
        'osheet.Cells(4, 9).value = "CW Budget Nett"
        'osheet.Cells(5, 9).value = "CW Forecast Nett"
        'osheet.Cells(6, 9).value = "CW Forecast vs Budget(%)"
        'osheet.Cells(7, 9).value = "Combined Budget Nett"
        'osheet.Cells(8, 9).value = "Combined Forecast Nett"
        'osheet.Cells(9, 9).value = "Most Recent Forecast"
        'osheet.Cells(10, 10).value = exRate
        Dim ColumnStartData = SGReportProperty1.ColumnStartData
        'For i = 0 To 11
        For i = 0 To 12
            osheet.Cells(1, ColumnStartData + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R[10]C[{1}]:R[{0}]C[{1}]<>""5 - COOKWARE""),R[10]C[{2}]:R[{0}]C[{2}],R[10]C:R[{0}]C)/1000", e.lastRow - 1, -(ColumnStartData - 1) - i, -1 - i)
            osheet.Cells(2, ColumnStartData + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R[9]C[{1}]:R[{0}]C[{1}]=""5 - COOKWARE""),R[9]C[{2}]:R[{0}]C[{2}],R[9]C:R[{0}]C)/1000", e.lastRow - 2, -(ColumnStartData - 1) - i, -1 - i)
            osheet.Cells(3, ColumnStartData + i).formulaR1C1 = String.Format("=R[-2]C[0]+R[-1]C[0]")
            'osheet.Cells(3, ColumnStartData + i).FormulaR1C1 = String.Format("=IFERROR(R[-1]C[0]/R[-2]C[0],0)")
            'osheet.Cells(6, ColumnStartData + i).formulaR1C1 = String.Format("=IFERROR(R[-1]C[0]/R[-2]C[0],0)")
            'osheet.Cells(7, ColumnStartData + i).formulaR1C1 = String.Format("=R[-6]C[0]+R[-3]C[0]")
            'osheet.Cells(8, ColumnStartData + i).formulaR1C1 = String.Format("=R[-6]C[0]+R[-3]C[0]")
            'osheet.Cells(9, ColumnStartData + i).formulaR1C1 = String.Format("=IFERROR(R[-1]C[0]/R[-2]C[0],0)")
            'Dim mySDAdict As Dictionary(Of Date, Decimal) = GroupSDABudgetDict(e.SheetNo - 2)
            'Dim myCWdict As Dictionary(Of Date, Decimal) = GroupCWBudgetDict(e.SheetNo - 2)
            'Dim mykey = CDate(String.Format("{0}-{1:MM}-01", myPeriodRange(i).Year, myPeriodRange(i)))
            'If (mySDAdict.ContainsKey(mykey)) Then
            '    osheet.Cells(1, ColumnStartData + i).value = mySDAdict(mykey)
            'End If
            'If (myCWdict.ContainsKey(mykey)) Then
            '    osheet.Cells(4, ColumnStartData + i).value = myCWdict(mykey)
            'End If

            ''Dim mydictT As Dictionary(Of Date, Decimal) = GroupGrossSalesTargetdict(e.SheetNo - 2)
            ''Dim mykeyT = CDate(String.Format("{0:yyyy-MM}-01", myPeriodRange(i)))
            ''If (mydictT.ContainsKey(mykeyT)) Then
            ''    osheet.Cells(4, ColumnStartData + i).value = mydictT(mykeyT)
            ''End If
        Next
        osheet.Range("J1:V3").Style = "Comma"
        osheet.Columns("H:I").numberformat = "#,##0.00"
        'osheet.Columns("G:H").numberformat = "0.0%"
        'osheet.Range("L1:W9").Style = "Comma"
        'osheet.Range("L3:W3").NumberFormat = "0.00%"
        'osheet.Range("L6:W6").NumberFormat = "0.00%"
        'osheet.Range("L9:W9").NumberFormat = "0.00%"

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

        ' For i = 0 To 11
        For i = 0 To 12
            Dim mysdabsb As New StringBuilder
            Dim mysdafsb As New StringBuilder
            Dim mycwbsb As New StringBuilder
            Dim mycwfsb As New StringBuilder

            For j = 0 To myKAMGroupList.Count - 1
                If mysdabsb.Length > 0 Then
                    mysdabsb.Append("+")
                    mysdafsb.Append("+")
                    mycwbsb.Append("+")
                    mycwfsb.Append("+")
                End If
                mysdabsb.Append(String.Format("'SG-{0}-{1}'!R[-2]C[{2}]", myKAMGroupList(j).kam, myKAMGroupList(j).groupname, 8))
                mysdafsb.Append(String.Format("'SG-{0}-{1}'!R[-2]C[{2}]", myKAMGroupList(j).kam, myKAMGroupList(j).groupname, 8))
            Next
            osheet.Cells(3, 2 + i).formulaR1C1 = String.Format("={0}", mysdabsb.ToString)
            osheet.Cells(4, 2 + i).formulaR1C1 = String.Format("={0}", mysdafsb.ToString)
            osheet.Cells(5, 2 + i).FormulaR1C1 = String.Format("=R[-2]C[0]+R[-1]C[0]")

        Next



        'For i = 0 To 11
        For i = 0 To 12
            Dim myresult As Integer = 0
            Try
                'myresult = KamTargetDict(myPeriodRange(i))
            Catch ex As Exception

            End Try
            'osheet.Cells(5, 2 + i).value = "" & myresult.ToString
            'osheet.Cells(6, 2 + i).FormulaR1C1 = String.Format("=R[-2]C-R[-1]C")
            'osheet.Cells(8, 2 + i).FormulaR1C1 = String.Format("=R[-4]C-R[-1]C")
            'this cell need to divided by exrate // Don't forget the all KAM report for exrate
            'osheet.Cells(9, 2 + i).FormulaR1C1 = String.Format("=R[-6]C/{0}", exRate)
        Next
        osheet.Range("B3:N9").Style = "Comma"
        osheet.Name = "SG-Summary"

        osheet.Cells.EntireColumn.AutoFit()
        osheet.Columns("A:N").columnwidth = 15
    End Sub
End Class
