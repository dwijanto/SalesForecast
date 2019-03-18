Imports Microsoft.Office.Interop

Public Class ALLKAMHK
    Dim myPeriodRange As Dictionary(Of Integer, Date)
    Dim myKAMAssignment As New KAMAssignmentController
    Dim myKAMAssignmentList As List(Of KAMAssignmentModel)
    Dim HKReportProperty1 As HKReportProperty = HKReportProperty.getInstance

    Public Sub Generate(myForm As Object, e As ALLKAMEventArgs)
        Dim sqlstr As String = String.Empty

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("SalesForecastALL{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            'Dim datasheet As Integer = 2
            Dim datasheet As Integer = 4

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable
            Dim criteria1 As String = String.Format("where tx.txdate >= '{0:yyyy-MM-}01' and tx.txdate <= '{1:yyyy-MM-}01' and tx.salesforecast > 0 ", e.startperiod, e.startperiod.AddMonths(5))
            Dim criteria2 As String = String.Format("where tx.txdate >= '{0:yyyy-MM-}01' and tx.txdate <= '{1:yyyy-MM-}01' and tx.salesforecast > 0 ", e.startperiod.AddMonths(6), e.endperiod)

            'sqlstr = String.Format("(select c.*,sales.get_producttype(c.productlinegpsid,c.brand) as producttype,tx.txdate,tx.kam,tx.mla as entity,tx.salesforecast,nsp.nsp1,nsp.nsp2,tx.salesforecast * nsp.nsp1 as grosssalesusd,tx.salesforecast * nsp.nsp2 as grosssaleshkd from sales.sfmlatxhk tx" &
            '         " left join sales.sfcmmfnsp nsp on nsp.cmmf = tx.cmmf" &
            '         " left join sales.sfcmmf c on c.cmmf = tx.cmmf {0}) union all" &
            '         " (select c.*,sales.get_producttype(c.productlinegpsid,c.brand) as producttype,tx.txdate,null::character varying as kam,tx.groupname,tx.salesforecast,nsp.nsp1,nsp.nsp2,tx.salesforecast * nsp.nsp1 as grosssalesusd,tx.salesforecast * nsp.nsp2 as grosssaleshkd " &
            '         " from sales.sfgrouptxhk tx" &
            '         " left join sales.sfcmmfnsp nsp on nsp.cmmf = tx.cmmf" &
            '         " left join sales.sfcmmf c on c.cmmf = tx.cmmf {1} )", criteria1, criteria2)
            'sqlstr = String.Format("(select c.*,sales.get_producttype(c.productlinegpsid,c.brand) as producttype,tx.txdate,tx.kam,tx.mla as entity,mc.cardname as customer,tx.salesforecast,nsp.nsp1,nsp.nsp2,tx.salesforecast * nsp.nsp1 as grosssalesusd,tx.salesforecast * nsp.nsp2 as grosssaleshkd from sales.sfmlatxhk tx" &
            '       " left join sales.sfcmmfnsp nsp on nsp.cmmf = tx.cmmf" &
            '       " left join sales.sfcmmf c on c.cmmf = tx.cmmf " &
            '       " left join sales.sfcmmfkamassignment cka on cka.id = tx.cmmfkamassignmentid" &
            '       " left join sales.sfkamassignment ka on ka.id = cka.kamassignmentid" &
            '       " left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid" &
            '       " {0}) union all" &
            '       " (select c.*,sales.get_producttype(c.productlinegpsid,c.brand) as producttype,tx.txdate,null::character varying as kam,tx.groupname,null::character varying as customer,tx.salesforecast,nsp.nsp1,nsp.nsp2,tx.salesforecast * nsp.nsp1 as grosssalesusd,tx.salesforecast * nsp.nsp2 as grosssaleshkd " &
            '       " from sales.sfgrouptxhk tx" &
            '       " left join sales.sfcmmfnsp nsp on nsp.cmmf = tx.cmmf" &
            '       " left join sales.sfcmmf c on c.cmmf = tx.cmmf {1} )", criteria1, criteria2)
            'sqlstr = String.Format("(select c.cmmf,c.origin,c.brand,c.reference,c.description,c.productsegmentation,gps.productlinegpsname as sbu,f.familyname,f.familylv2,c.activedate,sales.get_producttype(c.productlinegpsid,c.brand) as producttype,tx.txdate,tx.kam,tx.mla as entity,mc.cardname as customer,tx.salesforecast,nsp.nsp1,nsp.nsp2,tx.salesforecast * nsp.nsp1 as netsalesusd,tx.salesforecast * nsp.nsp2 as netsaleshkd from sales.sfcmmf c" &
            '       " left join sales.sfmlatxhk tx on c.cmmf = tx.cmmf " &
            '       " left join sales.sfcmmfnsp nsp on nsp.cmmf = c.cmmf" &
            '       " left join sales.sfcmmfkamassignment cka on cka.id = tx.cmmfkamassignmentid" &
            '       " left join sales.sfkamassignment ka on ka.id = cka.kamassignmentid" &
            '       " left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid" &
            '       " left join sales.sfproductlinegps gps on gps.productlinegpsid = c.productlinegpsid" &
            '       " left join sales.sffamily f on f.familyid = c.familyid" &
            '       " {0}) union all" &
            '       " (select c.cmmf,c.origin,c.brand,c.reference,c.description,c.productsegmentation,gps.productlinegpsname as sbu,f.familyname,f.familylv2,c.activedate,sales.get_producttype(c.productlinegpsid,c.brand) as producttype,tx.txdate,null::character varying as kam,tx.groupname,null::character varying as customer,tx.salesforecast,nsp.nsp1,nsp.nsp2,tx.salesforecast * nsp.nsp1 as netsalesusd,tx.salesforecast * nsp.nsp2 as netsaleshkd " &
            '       " from sales.sfcmmf c" &
            '       " left join sales.sfgrouptxhk tx on c.cmmf = tx.cmmf " &
            '       " left join sales.sfcmmfnsp nsp on nsp.cmmf = c.cmmf" &
            '       " left join sales.sfproductlinegps gps on gps.productlinegpsid = c.productlinegpsid" &
            '       " left join sales.sffamily f on f.familyid = c.familyid {1} )", criteria1, criteria2)
            sqlstr = String.Format("(select c.cmmf,c.origin,c.brand,c.reference,c.description,c.productsegmentation,gps.productlinegpsname as sbu,to_char(f.familyid,'FM000') as familyid,f.familyname,f.familylv2,f.productlinedesc,c.activedate,sales.get_producttype(c.productlinegpsid,c.brand) as producttype,tx.txdate,ka.kam,tx.mla as entity,mc.cardname as customer,tx.salesforecast,nsp.nsp1,nsp.nsp2,tx.salesforecast * nsp.nsp1 as netsalesusd,tx.salesforecast * nsp.nsp2 as netsaleshkd,tx.cmmfkamassignmentid from sales.sfcmmf c" &
                   " left join sales.sfmlatxhk tx on c.cmmf = tx.cmmf " &
                   " left join sales.sfcmmfnsp nsp on nsp.cmmf = c.cmmf" &
                   " left join sales.sfcmmfkamassignment cka on cka.id = tx.cmmfkamassignmentid" &
                   " left join sales.sfkamassignment ka on ka.id = cka.kamassignmentid" &
                   " left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid" &
                   " left join sales.sfproductlinegps gps on gps.productlinegpsid = c.productlinegpsid" &
                   " left join sales.sffamily f on f.familyid = c.familyid" &
                   " {0}) union all" &
                   " (select c.cmmf,c.origin,c.brand,c.reference,c.description,c.productsegmentation,gps.productlinegpsname as sbu,to_char(f.familyid,'FM000'),f.familyname,f.familylv2,f.productlinedesc,c.activedate,sales.get_producttype(c.productlinegpsid,c.brand) as producttype,tx.txdate,u.username::character varying as kam,tx.groupname,null::character varying as customer,tx.salesforecast,nsp.nsp1,nsp.nsp2,tx.salesforecast * nsp.nsp1 as netsalesusd,tx.salesforecast * nsp.nsp2 as netsaleshkd ,null::bigint" &
                   " from sales.sfcmmf c" &
                   " left join sales.sfgrouptxhk tx on c.cmmf = tx.cmmf " &
                   " left join sales._user u on u.id = tx.userid" &
                   " left join sales.sfcmmfnsp nsp on nsp.cmmf = c.cmmf" &
                   " left join sales.sfproductlinegps gps on gps.productlinegpsid = c.productlinegpsid" &
                   " left join sales.sffamily f on f.familyid = c.familyid {1} )", criteria1, criteria2)
            'Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\HKMLATemplate.xltx", HKReportProperty1.RowStartdData)
            'Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\HKMLATemplate.xltx", "A14")
            Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\HKALLKAMTemplate01.xltx", "A14")
            myreport.Run(myForm, e)

        End If
    End Sub

    Public Sub GenerateTarget(myForm As Object, e As ALLKAMEventArgs)

        Dim sqlstr As String = String.Empty

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("SalesForecastALLTarget{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 2

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTableTarget
            Dim criteria1 As String = String.Format("where tx.txdate >= '{0:yyyy-MM-}01' and tx.txdate <= '{1:yyyy-MM-}01' ", e.startperiod, e.endperiod)
            Dim criteria2 As String = String.Format("where p.period >= '{0:yyyy-MM-}01' and p.period <= '{1:yyyy-MM-}01' ", e.startperiod, e.endperiod)
            sqlstr = String.Format("(select sales.get_producttypeid(c.productlinegpsid::integer,c.brand) as producttypeid," &
                     " upper(sales.get_producttype(c.productlinegpsid,c.brand)) as producttype,tx.*,n.nsp2 ,tx.salesforecast * n.nsp2 as netsales,p.sdpct,(tx.salesforecast * n.nsp2) / (1 - p.sdpct) as totalgross" &
                     " from sales.sfmlatxhk tx left join sales.sfcmmfnsp n on n.cmmf = tx.cmmf left join sales.sfcmmf c on c.cmmf = tx.cmmf" &
                     " left join sales.sfhkparam p on p.kam = tx.kam and p.period = tx.txdate and p.producttype = sales.get_producttypeid(c.productlinegpsid::integer,c.brand) {0})" &
                     " union all (select p.producttype,sales.get_producttypename(producttype)::text as producttype,null::bigint,null::integer,period,null::bigint,null::integer,null::bigint,null::character varying," &
                     " p.kam,null::numeric,null::numeric,null::numeric,p.targetgross * -1 as targetgross from sales.sfhkparam p {1})", criteria1, criteria2)
            Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\HKMLATemplate.xltx", HKReportProperty1.RowStartdData)
            myreport.Run(myForm, e)

        End If
    End Sub
    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        Dim DataStart As Integer = HKReportProperty1.ColumnStartData
        osheet.Columns("L:L").numberformat = "dd-MMM-yyyy"
        osheet.Name = "DATA"
    End Sub

    Private Sub PivotTable1(ByRef oWB As Excel.Workbook, ByRef e As EventArgs)
        Dim oSheet As Excel.Worksheet
        oWB.Worksheets(1).select()
        oSheet = oWB.Worksheets(1)
        'oWB.Names.Add("dbRange", RefersToR1C1:="=OFFSET('DATA'!R18C1,0,0,COUNTA('DATA'!C18),COUNTA('DATA'!R14))")
        oWB.Names.Add("dbRange", RefersToR1C1:="=OFFSET('DATA'!R18C1,0,0,COUNTA('DATA'!C1),COUNTA('DATA'!R18))")
        oWB.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "dbRange").CreatePivotTable(oSheet.Name & "!R14C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With oSheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        oSheet.PivotTables("PivotTable1").pivotfields("customer").orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").pivotfields("kam").orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").pivotfields("sbu").orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").pivotfields("producttype").orientation = Excel.XlPivotFieldOrientation.xlPageField
        'oSheet.PivotTables("PivotTable1").pivotfields("productdesc").orientation = Excel.XlPivotFieldOrientation.xlPageField

        oSheet.PivotTables("PivotTable1").pivotfields("txdate").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        oSheet.PivotTables("PivotTable1").pivotfields("txdate").numberformat = "MMM-yy"
        'oSheet.Range("A15").Group(True, True, Periods:={False, False, False, True, True, False, True})
        'oSheet.PivotTables("PivotTable1").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlHidden
        'oSheet.PivotTables("PivotTable1").pivotfields("txdate").orientation = Excel.XlPivotFieldOrientation.xlHidden
        'oSheet.PivotTables("PivotTable1").Pivotfields("Months").orientation = Excel.XlPivotFieldOrientation.xlColumnField

        oSheet.PivotTables("PivotTable1").Pivotfields("familyname").orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1").Pivotfields("familyname").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        oSheet.PivotTables("PivotTable1").Pivotfields("brand").orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1").Pivotfields("brand").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        oSheet.PivotTables("PivotTable1").Pivotfields("reference").orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1").Pivotfields("reference").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        oSheet.PivotTables("PivotTable1").Pivotfields("description").orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1").Pivotfields("description").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}



        'oSheet.PivotTables("PivotTable1").Pivotfields("kam").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'oSheet.PivotTables("PivotTable1").Pivotfields("kam").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'oSheet.PivotTables("PivotTable1").Pivotfields("brand").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'oSheet.PivotTables("PivotTable1").Pivotfields("brand").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'oSheet.PivotTables("PivotTable1").Pivotfields("productdesc").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'oSheet.PivotTables("PivotTable1").Pivotfields("productdesc").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'oSheet.PivotTables("PivotTable1").Pivotfields("familyname").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'oSheet.PivotTables("PivotTable1").Pivotfields("familyname").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        'oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("nsp2"), " Sum NSP2 ", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("salesforecast"), " Sum Sales Forecast ", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").PivotFields(" Sum Sales Forecast ").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
        'oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("grosssaleshkd"), " Sum Gross Sales (HKD) ", Excel.XlConsolidationFunction.xlSum)

        'For Each PV As Excel.PivotItem In oSheet.PivotTables("PivotTable1").PivotFields("kam").PivotItems
        '    If PV.Value = "(blank)" Then
        '        PV.Visible = False
        '    End If
        'Next
        oSheet.Cells.EntireColumn.AutoFit()
    End Sub
    Private Sub PivotTable(ByRef oWB As Excel.Workbook, ByRef e As EventArgs)
        oWB.Worksheets(1).select()
        Dim osheet As Excel.Worksheet
        osheet = oWB.Worksheets(1)
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        osheet = oWB.Worksheets(2)
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        osheet = oWB.Worksheets(3)
        osheet.PivotTables("PivotTable2").PivotCache.Refresh()
        'Dim oSheet As Excel.Worksheet
        'oWB.Worksheets(1).select()
        'oSheet = oWB.Worksheets(1)
        'oWB.Names.Add("dbRange", RefersToR1C1:="=OFFSET('DATA'!R14C1,0,0,COUNTA('DATA'!C18),COUNTA('DATA'!R14))")
        'oWB.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "dbRange").CreatePivotTable(oSheet.Name & "!R14C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        'With oSheet.PivotTables("PivotTable1")
        '    .ingriddropzones = True
        '    .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        '    .DisplayErrorString = True
        'End With

        'oSheet.PivotTables("PivotTable1").pivotfields("customer").orientation = Excel.XlPivotFieldOrientation.xlPageField
        'oSheet.PivotTables("PivotTable1").pivotfields("kam").orientation = Excel.XlPivotFieldOrientation.xlPageField
        'oSheet.PivotTables("PivotTable1").pivotfields("sbu").orientation = Excel.XlPivotFieldOrientation.xlPageField
        'oSheet.PivotTables("PivotTable1").pivotfields("producttype").orientation = Excel.XlPivotFieldOrientation.xlPageField
        ''oSheet.PivotTables("PivotTable1").pivotfields("productdesc").orientation = Excel.XlPivotFieldOrientation.xlPageField

        'oSheet.PivotTables("PivotTable1").pivotfields("txdate").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        'oSheet.PivotTables("PivotTable1").pivotfields("txdate").numberformat = "MMM-yy"
        ''oSheet.Range("A15").Group(True, True, Periods:={False, False, False, True, True, False, True})
        ''oSheet.PivotTables("PivotTable1").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlHidden
        ''oSheet.PivotTables("PivotTable1").pivotfields("txdate").orientation = Excel.XlPivotFieldOrientation.xlHidden
        ''oSheet.PivotTables("PivotTable1").Pivotfields("Months").orientation = Excel.XlPivotFieldOrientation.xlColumnField

        'oSheet.PivotTables("PivotTable1").Pivotfields("familyname").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'oSheet.PivotTables("PivotTable1").Pivotfields("familyname").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'oSheet.PivotTables("PivotTable1").Pivotfields("brand").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'oSheet.PivotTables("PivotTable1").Pivotfields("brand").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'oSheet.PivotTables("PivotTable1").Pivotfields("reference").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'oSheet.PivotTables("PivotTable1").Pivotfields("reference").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'oSheet.PivotTables("PivotTable1").Pivotfields("description").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'oSheet.PivotTables("PivotTable1").Pivotfields("description").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}



        ''oSheet.PivotTables("PivotTable1").Pivotfields("kam").orientation = Excel.XlPivotFieldOrientation.xlRowField
        ''oSheet.PivotTables("PivotTable1").Pivotfields("kam").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        ''oSheet.PivotTables("PivotTable1").Pivotfields("brand").orientation = Excel.XlPivotFieldOrientation.xlRowField
        ''oSheet.PivotTables("PivotTable1").Pivotfields("brand").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        ''oSheet.PivotTables("PivotTable1").Pivotfields("productdesc").orientation = Excel.XlPivotFieldOrientation.xlRowField
        ''oSheet.PivotTables("PivotTable1").Pivotfields("productdesc").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        ''oSheet.PivotTables("PivotTable1").Pivotfields("familyname").orientation = Excel.XlPivotFieldOrientation.xlRowField
        ''oSheet.PivotTables("PivotTable1").Pivotfields("familyname").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        ''oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("nsp2"), " Sum NSP2 ", Excel.XlConsolidationFunction.xlSum)
        'oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("salesforecast"), " Sum Sales Forecast ", Excel.XlConsolidationFunction.xlSum)
        'oSheet.PivotTables("PivotTable1").PivotFields(" Sum Sales Forecast ").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
        ''oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("grosssaleshkd"), " Sum Gross Sales (HKD) ", Excel.XlConsolidationFunction.xlSum)

        ''For Each PV As Excel.PivotItem In oSheet.PivotTables("PivotTable1").PivotFields("kam").PivotItems
        ''    If PV.Value = "(blank)" Then
        ''        PV.Visible = False
        ''    End If
        ''Next
        'oSheet.Cells.EntireColumn.AutoFit()
    End Sub
    Private Sub PivotTableTarget(ByRef oWB As Excel.Workbook, ByRef e As EventArgs)
        Dim oSheet As Excel.Worksheet
        oWB.Worksheets(1).select()
        oSheet = oWB.Worksheets(1)
        oWB.Names.Add("dbRange", RefersToR1C1:="=OFFSET('DATA'!R18C1,0,0,COUNTA('DATA'!C1),COUNTA('DATA'!R18))")
        oWB.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "dbRange").CreatePivotTable(oSheet.Name & "!R18C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With oSheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        oSheet.PivotTables("PivotTable1").pivotfields("producttype").orientation = Excel.XlPivotFieldOrientation.xlPageField
        
        oSheet.PivotTables("PivotTable1").pivotfields("txdate").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        oSheet.PivotTables("PivotTable1").pivotfields("txdate").numberformat = "MMM-yy"
        
        oSheet.PivotTables("PivotTable1").Pivotfields("kam").orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1").Pivotfields("kam").subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}


        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("totalgross"), " Sum Total Gross ", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").PivotFields(" Sum Total Gross ").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"

        oSheet.Cells.EntireColumn.AutoFit()
    End Sub
End Class
