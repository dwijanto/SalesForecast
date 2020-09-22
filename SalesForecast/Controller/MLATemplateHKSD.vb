Imports System.Text
Imports Microsoft.Office.Interop

Public Class MLATemplateHKSD
    Dim myPeriodRange As Dictionary(Of Integer, Date)
    Dim myKAMAssignment As New KAMAssignmentController
    Dim myKAMAssignmentList As List(Of KAMAssignmentModel)
    Dim HKReportProperty1 As HKReportProperty = HKReportProperty.getInstance
    Dim HKParamDT As DataTable
    Public Sub Generate(myForm As Object, e As MLATemplateHKEventArgs)

        Dim sqlstr As String = String.Empty '= "select * from sales.sfcmmf;"

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("MyFile{0}{1:yyyyMMdd}.xlsx", e.userName, Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            'get HKParam
            Dim HKParamController1 As New HKParamController
            HKParamController1.loaddata(String.Format(" where kam = '{0}'", e.userName))
            HKParamDT = HKParamController1.GetTable

            myKAMAssignmentList = New List(Of KAMAssignmentModel)
            myKAMAssignmentList = myKAMAssignment.getAssignments(e.userName)
            Dim fieldList As New StringBuilder
            Dim txfieldlist As New StringBuilder
            For i = 0 To myKAMAssignmentList.Count - 1
                If fieldList.Length > 0 Then
                    fieldList.Append(",")
                    txfieldlist.Append(",")
                End If
                If myKAMAssignmentList(i).cardname <> "" Then
                    fieldList.Append(String.Format("""{0}"" int", myKAMAssignmentList(i).cardname))
                    txfieldlist.Append(String.Format("tx.""{0}""", myKAMAssignmentList(i).cardname))
                End If

            Next

            If txfieldlist.Length = 0 Then
                MessageBox.Show("Sorry, no data for this KAM.")
                Exit Sub
            End If

            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1 'because hidden

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim PeriodRange As New PeriodRange(e.startPeriod, 6)
            myPeriodRange = PeriodRange.getPeriod
            Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\HKMLATemplate.xltx", HKReportProperty1.RowStartdData)
            myreport.QueryList = New List(Of QueryWorksheet)

            For i = 0 To 5
                Dim myQuery = New QueryWorksheet
                Dim username = e.userName
                If e.blanktemplate Then

                    sqlstr = String.Format("with tx as (select * from crosstab('" &
                           " with ka as (select distinct kam.username,mc.mla,mc.cardname,trim(mc.mla) || '' - '' || mc.cardname as assignment  from sales.sfkam kam  " &
                           " left join sales.sfkamassignment ka on ka.kam = kam.username  	left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid where kam.username = ''{0}'' " &
                           " order by mc.mla) , kam as (select 1 as id,row_number() over (order by ka.mla,ka.cardname) as recid,* from ka), " &
                           " cmmf as(select  1 as id,* from sales.sfcmmf cmmf order by cmmf), " &
                           " sa as (select tx.txdate,null::bigint as salesforecast,tx.cmmfkamassignmentid,cka.cmmf,ka.kam,mc.mla,mc.cardname from sales.sfmlatxhk tx " &
                           " left join sales.sfcmmfkamassignment cka on cka.id = tx.cmmfkamassignmentid" &
                           " left join sales.sfkamassignment ka on ka.id = cka.kamassignmentid left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid " &
                           " where tx.txdate = ''{1:yyyy-MM-dd}'' and ka.kam = ''{0}'' ),tx as(select cmmf.cmmf,kam.recid,sa.salesforecast from cmmf " &
                           " left join kam on kam.id = cmmf.id  left join sa on sa.cmmf = cmmf.cmmf and sa.kam = kam.username and sa.mla = kam.mla and sa.cardname = kam.cardname 	order by cmmf,kam.mla) select * from tx;','select m from generate_series(1,{2})m') as (cmmf bigint,{3})), " &
                           " nsp as (select cmmf, nsp1,nsp2 from sales.sfcmmfnsp)" &
                           " select sales.get_producttype(productlinegpsid,brand) as producttype ,c.brand,c.productlinegps,c.reference,to_char(c.familyid,'FM000') || ' - ' || f.familyname || ' - ' || f.familylv2 as familyname,c.cmmf,c.description ,n.nsp1 as ""NSP(USD)"",n.nsp2 as ""NSP(HKD)"" ,{4},null::text,null::text,c.launchingmonth,c.remarks " &
                           " from sales.sfcmmf c left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf left join sales.sffamily f on f.familyid = c.familyid  where to_char(activedate,'YYYYMM')::int <= to_char('{1:yyyy-MM-dd}'::date,'YYYYMM')::int  order by producttype,c.brand,c.reference ", username, myPeriodRange(i), myKAMAssignmentList.Count, fieldList.ToString, txfieldlist.ToString)
                Else

                    'sqlstr = String.Format("with tx as (select * from crosstab('" &
                    '        " with ka as (select distinct kam.username,mc.mla,mc.cardname,trim(mc.mla) || '' - '' || mc.cardname as assignment  from sales.sfkam kam  " &
                    '        " left join sales.sfkamassignment ka on ka.kam = kam.username  	left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid where kam.username = ''{0}'' " &
                    '        " order by mc.mla) , kam as (select 1 as id,row_number() over (order by ka.mla,ka.cardname) as recid,* from ka), " &
                    '        " cmmf as(select  1 as id,* from sales.sfcmmf cmmf order by cmmf), " &
                    '        " sa as (select tx.txdate,tx.salesforecast,tx.cmmfkamassignmentid,cka.cmmf,ka.kam,mc.mla,mc.cardname from sales.sfmlatxhk tx " &
                    '        " left join sales.sfcmmfkamassignment cka on cka.id = tx.cmmfkamassignmentid" &
                    '        " left join sales.sfkamassignment ka on ka.id = cka.kamassignmentid left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid " &
                    '        " where tx.txdate = ''{1:yyyy-MM-dd}'' and ka.kam = ''{0}'' ),tx as(select cmmf.cmmf,kam.recid,sa.salesforecast from cmmf " &
                    '        " left join kam on kam.id = cmmf.id  left join sa on sa.cmmf = cmmf.cmmf and sa.kam = kam.username and sa.mla = kam.mla and sa.cardname = kam.cardname 	order by cmmf,kam.mla) select * from tx;','select m from generate_series(1,{2})m') as (cmmf bigint,{3})), " &
                    '        " nsp as (select cmmf, nsp1,nsp2 from sales.sfcmmfnsp )" &
                    '        " select sales.get_producttype(productlinegpsid,brand) as producttype ,c.brand,c.productlinegps,c.reference,to_char(c.familyid,'FM000') || ' - ' || f.familyname || ' - ' || f.familylv2 as familyname,c.cmmf,c.description ,n.nsp1 as ""NSP(USD)"",n.nsp2 as ""NSP(HKD)"" ,{4},null::text,null::text,c.launchingmonth,c.remarks " &
                    '        " from sales.sfcmmf c left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf left join sales.sffamily f on f.familyid = c.familyid where to_char(activedate,'YYYYMM')::int <= to_char('{1:yyyy-MM-dd}'::date,'YYYYMM')::int  order by producttype,c.brand,c.reference  ", username, myPeriodRange(i), myKAMAssignmentList.Count, fieldList.ToString, txfieldlist.ToString)
                    sqlstr = String.Format("with tx as (select * from crosstab('" &
                            " with ka as (select distinct kam.username,mc.mla,mc.cardname,trim(mc.mla) || '' - '' || mc.cardname as assignment  from sales.sfkam kam  " &
                            " left join sales.sfkamassignment ka on ka.kam = kam.username  	left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid where kam.username = ''{0}'' " &
                            " order by mc.mla) , kam as (select 1 as id,row_number() over (order by ka.mla,ka.cardname) as recid,* from ka), " &
                            " cmmf as(select  1 as id,* from sales.sfcmmf cmmf order by cmmf), " &
                            " sa as (select tx.txdate,tx.salesforecast,tx.cmmfkamassignmentid,cka.cmmf,ka.kam,mc.mla,mc.cardname from sales.sfmlatxhk tx " &
                            " left join sales.sfcmmfkamassignment cka on cka.id = tx.cmmfkamassignmentid" &
                            " left join sales.sfkamassignment ka on ka.id = cka.kamassignmentid left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid " &
                            " where tx.txdate = ''{1:yyyy-MM-dd}'' and ka.kam = ''{0}'' ),tx as(select cmmf.cmmf,kam.recid,sa.salesforecast from cmmf " &
                            " left join kam on kam.id = cmmf.id  left join sa on sa.cmmf = cmmf.cmmf and sa.kam = kam.username and sa.mla = kam.mla and sa.cardname = kam.cardname 	order by cmmf,kam.mla) select * from tx;','select m from generate_series(1,{2})m') as (cmmf bigint,{3})) " &
                            " select sales.get_producttype(productlinegpsid,brand) as producttype ,c.brand,c.productlinegps,c.reference,to_char(c.familyid,'FM000') || ' - ' || f.familyname || ' - ' || f.familylv2 as familyname,c.cmmf,c.description ,n.nsp1 as ""NSP(Fortress)"",d.nsp1 as ""NSP(Default)"" ,{4},null::text,null::text,c.launchingmonth,c.remarks " &
                            " from sales.sfcmmf c " &
                            " left join tx on tx.cmmf = c.cmmf " &
                            " left join sales.sffamily f on f.familyid = c.familyid" &
                            " left join sales.sfmlansp n on n.cmmf = tx.cmmf and n.mla = 'W9000332' and n.period =  to_char('{1:yyyy-MM-dd}'::date,'yyyyMM')::integer" &
                            " left join sales.sfmlansp d on d.cmmf = tx.cmmf and d.mla ='Default' and d.period =  to_char('{1:yyyy-MM-dd}'::date,'yyyyMM')::integer" &
                            " where to_char(activedate,'YYYYMM')::int <= to_char('{1:yyyy-MM-dd}'::date,'YYYYMM')::int  order by producttype,c.brand,c.reference  ",
                            username, myPeriodRange(i), myKAMAssignmentList.Count, fieldList.ToString, txfieldlist.ToString)
                End If


                myQuery.Sqlstr = sqlstr
                myQuery.DataSheet = i + 1
                myQuery.SheetName = String.Format("HK-MLA-{0:yyyy.MM}", myPeriodRange(i)) '"2016.09"
                myreport.QueryList.Add(myQuery)
            Next
            myreport.Run(myForm, e)

        End If
    End Sub

    Private Sub FormattingReport1(ByRef sender As Object, ByRef e As ExportToExcelFileEventArgs)
       
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As ExportToExcelFileEventArgs)
        Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        Dim owb As Excel.Workbook = osheet.Parent
        Dim oXL As Excel.Application = owb.Parent
        Dim DataStart As Integer = HKReportProperty1.ColumnStartData
        Dim RowStartData As Integer = HKReportProperty1.RowStartDataI

        Dim lastcolumn = DataStart + myKAMAssignmentList.Count
        For i = 0 To myKAMAssignmentList.Count - 1

            'osheet.Cells(HKReportProperty1.RowStartDataI - 5, DataStart + i) = myKAMAssignmentList(i).sdasd
            'osheet.Cells(HKReportProperty1.RowStartDataI - 4, DataStart + i) = myKAMAssignmentList(i).tefalsd
            'osheet.Cells(HKReportProperty1.RowStartDataI - 3, DataStart + i) = myKAMAssignmentList(i).lagosd
            osheet.Cells(HKReportProperty1.RowStartDataI - 6, DataStart + i) = myKAMAssignmentList(i).sdasd
            osheet.Cells(HKReportProperty1.RowStartDataI - 5, DataStart + i) = myKAMAssignmentList(i).tefalsd
            osheet.Cells(HKReportProperty1.RowStartDataI - 4, DataStart + i) = myKAMAssignmentList(i).lagosd
            osheet.Cells(HKReportProperty1.RowStartDataI - 3, DataStart + i) = myKAMAssignmentList(i).wmfsd

            osheet.Cells(HKReportProperty1.RowStartDataI - 2, DataStart + i) = myKAMAssignmentList(i).username
            osheet.Cells(HKReportProperty1.RowStartDataI - 1, DataStart + i) = myKAMAssignmentList(i).mla


            'osheet.Cells(1, DataStart + i) = "Net"
            'osheet.Cells(1, DataStart + i).font.bold = True
            'osheet.Cells(1, DataStart + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            'osheet.Cells(1, DataStart + i).Interior.TintAndShade = 0.599993896298105
            'osheet.Cells(6, DataStart + i) = "Gross"
            'osheet.Cells(6, DataStart + i).font.bold = True
            'osheet.Cells(6, DataStart + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
            'osheet.Cells(6, DataStart + i).Interior.TintAndShade = 0.599993896298105

            'Gross SDA 
            osheet.Cells(7, DataStart + i).formulaR1C1 = String.Format("=R1C/(1-R13C)")
            osheet.Cells(7, DataStart + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
            osheet.Cells(7, DataStart + i).Interior.TintAndShade = 0.599993896298105
            'Gross SDA WMF
            osheet.Cells(8, DataStart + i).formulaR1C1 = String.Format("=R2C/(1-R13C)")
            osheet.Cells(8, DataStart + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
            osheet.Cells(8, DataStart + i).Interior.TintAndShade = 0.599993896298105
            'Gross CKW TEFAL
            osheet.Cells(9, DataStart + i).formulaR1C1 = String.Format("=R3C/(1-R14C)")
            osheet.Cells(9, DataStart + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
            osheet.Cells(9, DataStart + i).Interior.TintAndShade = 0.599993896298105
            'Gross CKW LAGO
            osheet.Cells(10, DataStart + i).formulaR1C1 = String.Format("=R4C/(1-R15C)")
            osheet.Cells(10, DataStart + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
            osheet.Cells(10, DataStart + i).Interior.TintAndShade = 0.599993896298105
            'Gross CKW WMF
            osheet.Cells(11, DataStart + i).formulaR1C1 = String.Format("=R5C/(1-R16C)")
            osheet.Cells(11, DataStart + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
            osheet.Cells(11, DataStart + i).Interior.TintAndShade = 0.599993896298105
            'Subtotal Gross
            osheet.Cells(12, DataStart + i).formulaR1C1 = "=sum(R[-5]C:R[-1]C)"
            osheet.Cells(12, DataStart + i).font.bold = True
            osheet.Cells(12, DataStart + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
            osheet.Cells(12, DataStart + i).Interior.TintAndShade = 0.399975585192419

            ''SubTotal Net
            'osheet.Cells(6, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(R{2}C{0}:R{1}C{0},R{2}C:R{1}C)", 9, e.lastRow, RowStartData + 1)
            'osheet.Cells(6, 10 + i).font.bold = True
            'osheet.Cells(6, 10 + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            'osheet.Cells(6, 10 + i).Interior.TintAndShade = 0.399975585192419

            ''Net SDA
            'osheet.Cells(1, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R{3}C{2}:R{0}C{2}=""SDA""),R{3}C{1}:R{0}C{1},R{3}C:R{0}C)", e.lastRow, 9, 1, RowStartData + 1)
            'osheet.Cells(1, 10 + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            'osheet.Cells(1, 10 + i).Interior.TintAndShade = 0.599993896298105
            ''Net SDA WMF
            'osheet.Cells(2, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R{3}C{2}:R{0}C{2}=""SDA - WMF""),R{3}C{1}:R{0}C{1},R{3}C:R{0}C)", e.lastRow, 9, 1, RowStartData + 1)
            'osheet.Cells(2, 10 + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            'osheet.Cells(2, 10 + i).Interior.TintAndShade = 0.599993896298105
            ''Net CKW TEFAL
            'osheet.Cells(3, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R{3}C{2}:R{0}C{2}=""CKW - Tefal""),R{3}C{1}:R{0}C{1},R{3}C:R{0}C)", e.lastRow, 9, 1, RowStartData + 1)
            'osheet.Cells(3, 10 + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            'osheet.Cells(3, 10 + i).Interior.TintAndShade = 0.599993896298105
            ''Net CKW LAGO
            'osheet.Cells(4, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R{3}C{2}:R{0}C{2}=""CKW - Lago""),R{3}C{1}:R{0}C{1},R{3}C:R{0}C)", e.lastRow, 9, 1, RowStartData + 1)
            'osheet.Cells(4, 10 + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            'osheet.Cells(4, 10 + i).Interior.TintAndShade = 0.599993896298105
            ''Net CKW WMF
            'osheet.Cells(5, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R{3}C{2}:R{0}C{2}=""CKW - WMF""),R{3}C{1}:R{0}C{1},R{3}C:R{0}C)", e.lastRow, 9, 1, RowStartData + 1)
            'osheet.Cells(5, 10 + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            'osheet.Cells(5, 10 + i).Interior.TintAndShade = 0.599993896298105

            'SubTotal Net
            osheet.Cells(6, 10 + i).FormulaR1C1 = String.Format("=sum(R[-5]C:R[-1]C)")
            osheet.Cells(6, 10 + i).font.bold = True
            osheet.Cells(6, 10 + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            osheet.Cells(6, 10 + i).Interior.TintAndShade = 0.399975585192419

            'Net SDA
            osheet.Cells(1, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R{3}C{2}:R{0}C{2}=""SDA""),if(R18C=""W9000332"",R{3}C{4}:R{0}C{4},R{3}C{1}:R{0}C{1}),R{3}C:R{0}C)", e.lastRow, 9, 1, RowStartData + 1, 8)
            osheet.Cells(1, 10 + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            osheet.Cells(1, 10 + i).Interior.TintAndShade = 0.599993896298105
            'Net SDA WMF
            osheet.Cells(2, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R{3}C{2}:R{0}C{2}=""SDA - WMF""),if(R18C=""W9000332"",R{3}C{4}:R{0}C{4},R{3}C{1}:R{0}C{1}),R{3}C:R{0}C)", e.lastRow, 9, 1, RowStartData + 1, 8)
            osheet.Cells(2, 10 + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            osheet.Cells(2, 10 + i).Interior.TintAndShade = 0.599993896298105
            'Net CKW TEFAL
            osheet.Cells(3, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R{3}C{2}:R{0}C{2}=""CKW - Tefal""),if(R18C=""W9000332"",R{3}C{4}:R{0}C{4},R{3}C{1}:R{0}C{1}),R{3}C:R{0}C)", e.lastRow, 9, 1, RowStartData + 1, 8)
            osheet.Cells(3, 10 + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            osheet.Cells(3, 10 + i).Interior.TintAndShade = 0.599993896298105
            'Net CKW LAGO
            osheet.Cells(4, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R{3}C{2}:R{0}C{2}=""CKW - Lago""),if(R18C=""W9000332"",R{3}C{4}:R{0}C{4},R{3}C{1}:R{0}C{1}),R{3}C:R{0}C)", e.lastRow, 9, 1, RowStartData + 1, 8)
            osheet.Cells(4, 10 + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            osheet.Cells(4, 10 + i).Interior.TintAndShade = 0.599993896298105
            'Net CKW WMF
            osheet.Cells(5, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R{3}C{2}:R{0}C{2}=""CKW - WMF""),if(R18C=""W9000332"",R{3}C{4}:R{0}C{4},R{3}C{1}:R{0}C{1}),R{3}C:R{0}C)", e.lastRow, 9, 1, RowStartData + 1, 8)
            osheet.Cells(5, 10 + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            osheet.Cells(5, 10 + i).Interior.TintAndShade = 0.599993896298105
        Next

        Dim tmp = osheet.Name.Split("-")
        Dim mydate = String.Format(tmp(2).Replace(".", "-") & "-1")
        
        osheet.Cells(1, 4) = "Net Sales"
        osheet.Cells(1, 4).font.bold = True

        osheet.Cells(7, 4) = "Gross Sales"
        osheet.Cells(7, 4).font.bold = True
        osheet.Cells(8, 2) = "Input by KAM"
        osheet.Cells(8, 2).font.bold = True
        osheet.Cells(10, 2) = "Difference"
        osheet.Cells(10, 2).font.bold = True

        osheet.Cells(1, 7) = "SDA"
        osheet.Cells(2, 7) = "SDA WMF"
        osheet.Cells(3, 7) = "CKW TEFAL"
        osheet.Cells(4, 7) = "CKW LAGO"
        osheet.Cells(5, 7) = "CKW WMF"

        osheet.Cells(7, 7) = "SDA"
        osheet.Cells(8, 7) = "SDA WMF"
        osheet.Cells(9, 7) = "CKW TEFAL"
        osheet.Cells(10, 7) = "CKW LAGO"
        osheet.Cells(11, 7) = "CKW WMF"

        osheet.Cells(13, 7) = "SDA"
        osheet.Cells(14, 7) = "CKW TEFAL"
        osheet.Cells(15, 7) = "CKW LAGO"
        osheet.Cells(16, 7) = "CKW WMF"

        osheet.Cells(1, lastcolumn + 1) = "SDA"
        osheet.Cells(2, lastcolumn + 1) = "SDA WMF"
        osheet.Cells(3, lastcolumn + 1) = "CKW TEFAL"
        osheet.Cells(4, lastcolumn + 1) = "CKW LAGO"
        osheet.Cells(5, lastcolumn + 1) = "CKW WMF"
        osheet.Cells(6, lastcolumn + 1).value = "Total"

        'Total NET Sales SDA as the end
        osheet.Cells(1, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count) 'Total SDA
        osheet.Cells(1, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
        osheet.Cells(1, lastcolumn).Interior.TintAndShade = 0.599993896298105
        osheet.Cells(2, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count) 'Total SDA WMF
        osheet.Cells(2, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
        osheet.Cells(2, lastcolumn).Interior.TintAndShade = 0.599993896298105

        osheet.Cells(3, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count) 'Total TEFAL
        osheet.Cells(3, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
        osheet.Cells(3, lastcolumn).Interior.TintAndShade = 0.599993896298105
        osheet.Cells(4, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count) 'Total Lago
        osheet.Cells(4, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
        osheet.Cells(4, lastcolumn).Interior.TintAndShade = 0.599993896298105

        osheet.Cells(5, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count) 'Total WMF
        osheet.Cells(5, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
        osheet.Cells(5, lastcolumn).Interior.TintAndShade = 0.599993896298105

        osheet.Cells(6, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count) 'Total All
        osheet.Cells(6, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
        osheet.Cells(6, lastcolumn).Interior.TintAndShade = 0.399975585192419
        osheet.Cells(6, lastcolumn).font.bold = True
        osheet.Cells(6, lastcolumn + 1).font.bold = True
        osheet.Cells(6, lastcolumn + 1).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
        osheet.Cells(6, lastcolumn + 1).Interior.TintAndShade = 0.399975585192419
        
        osheet.Cells(7, lastcolumn + 1).value = "SDA"
        osheet.Cells(8, lastcolumn + 1).value = "SDA WMF"
        osheet.Cells(9, lastcolumn + 1).value = "CKW TEFAL"
        osheet.Cells(10, lastcolumn + 1).value = "CKW LAGO"
        osheet.Cells(11, lastcolumn + 1).value = "CKW WMF"
        osheet.Cells(12, lastcolumn + 1).value = "Total"
        osheet.Cells(12, lastcolumn).font.bold = True
        osheet.Cells(12, lastcolumn + 1).font.bold = True
        osheet.Cells(12, lastcolumn + 1).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        osheet.Cells(12, lastcolumn + 1).Interior.TintAndShade = 0.399975585192419

        'Total Gross
        osheet.Cells(7, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count) 'Total Gross SDA
        osheet.Cells(7, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        osheet.Cells(7, lastcolumn).Interior.TintAndShade = 0.599993896298105

        osheet.Cells(8, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count) 'Total Gross SDA WMF
        osheet.Cells(8, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        osheet.Cells(8, lastcolumn).Interior.TintAndShade = 0.599993896298105
        osheet.Cells(9, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count) 'Total Gross CKW TEFAL
        osheet.Cells(9, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        osheet.Cells(9, lastcolumn).Interior.TintAndShade = 0.599993896298105
        osheet.Cells(10, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count) 'Total Gross CKW LAGO
        osheet.Cells(10, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        osheet.Cells(10, lastcolumn).Interior.TintAndShade = 0.599993896298105
        osheet.Cells(11, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count) 'Total Gross CKW WMF
        osheet.Cells(11, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        osheet.Cells(11, lastcolumn).Interior.TintAndShade = 0.599993896298105

        osheet.Cells(12, lastcolumn).formulaR1C1 = "=sum(R[-5]C:R[-1]C)"
        osheet.Cells(12, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        osheet.Cells(12, lastcolumn).Interior.TintAndShade = 0.399975585192419

        'osheet.Cells(1, lastcolumn).value = "Total Net"
        'osheet.Cells(1, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
        'osheet.Cells(1, lastcolumn).Interior.TintAndShade = 0.599993896298105
        'osheet.Cells(6, lastcolumn).value = "Total Gross"
        'osheet.Cells(6, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        'osheet.Cells(6, lastcolumn).Interior.TintAndShade = 0.599993896298105

        'Check PCT Net \ Gross
        osheet.Cells(15, lastcolumn).formulaR1C1 = "=1-(R[-9]C/R[-3]C)"
        osheet.Cells(15, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
        osheet.Cells(15, lastcolumn).Interior.TintAndShade = 0.599993896298105

        'Header at the last 
        osheet.Cells(HKReportProperty1.RowStartDataI, lastcolumn).value = "Total"
        osheet.Cells(8, 4).Interior.color = 65535

        osheet.Cells(10, 4).formulaR1C1 = String.Format("=IFERROR(R12c[{0}]-R8C,)", lastcolumn - 4)

        'osheet.Range("D2:D4").NumberFormat = "0.0%"
        osheet.Range("H14:I14").NumberFormat = "#,##0"
        osheet.Range("F2:F10").NumberFormat = "#,##0"
        osheet.Range("G2:G5").NumberFormat = "#,##0"
        osheet.Range("D8:D10").NumberFormat = "#,##0"

        osheet.Range(osheet.Cells(1, DataStart), osheet.Cells(12, lastcolumn)).NumberFormat = "#,##0"
        osheet.Range(osheet.Cells(13, DataStart), osheet.Cells(16, lastcolumn)).NumberFormat = "0.0%"
        'SDA
        'osheet.Cells(5, 10).FormulaR1C1 = String.Format("=SUMPRODUCT(R[9]C[-1]:R[{0}]C[-1],R[9]C:R[{0}]C)", e.lastRow)

        'Dim oRange As Excel.Range
        'oRange = osheet.Range(HKReportProperty1.RowStartdData)
        'oRange.Select()
        'osheet.Application.Selection.autofilter() 'should do twice because of the initial is fi
        'osheet.Application.Selection.autofilter()
        osheet.Cells.EntireColumn.AutoFit()
        osheet.Columns(lastcolumn).ColumnWidth = 15

        'osheet.Cells(10, 6).font.bold = True
        'osheet.Cells(10, 6).font.size = 14

        'osheet.Cells(10, 7).value = "Should Be ""0"""
        'osheet.Cells(10, 7).font.FontStyle = "Italic"
        'osheet.Cells(10, 7).font.bold = True
        'osheet.Cells(10, 7).font.size = 14
        'osheet.Columns(1).EntireColumn.Hidden = True
        osheet.Columns(3).EntireColumn.Hidden = True
        'osheet.Columns(5).EntireColumn.Hidden = True
        osheet.Columns(8).EntireColumn.Hidden = True
        osheet.Columns(9).EntireColumn.Hidden = True
        osheet.Rows("13:16").EntireRow.Hidden = True
        osheet.Columns(7).columnwidth = 40

        'SubTotal Row
        For i = (HKReportProperty1.RowStartDataI + 1) To e.lastRow
            osheet.Cells(i, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count)
        Next
        osheet.Range(osheet.Cells(HKReportProperty1.RowStartDataI + 1, DataStart), osheet.Cells(HKReportProperty1.RowStartDataI + e.lastRow, lastcolumn + 5)).Locked = False
        osheet.Range("D7").Locked = False
        osheet.Range(osheet.Cells(1, 1), osheet.Cells(10, lastcolumn)).FormulaHidden = True
        Dim Identitiy As UserController = User.getIdentity
        If Not Identitiy.isAdmin Then
            osheet.Protect(DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="salesforecast", AllowFiltering:=True)
            osheet.EnableSelection = Excel.XlEnableSelection.xlNoRestrictions

        End If
        osheet.Cells(HKReportProperty1.RowStartDataI + 1, DataStart).select()
        oXL.ActiveWindow.FreezePanes = True

    End Sub
    Private Sub FormattingReportOld(ByRef sender As Object, ByRef e As ExportToExcelFileEventArgs)
        Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        Dim owb As Excel.Workbook = osheet.Parent
        Dim oXL As Excel.Application = owb.Parent
        Dim DataStart As Integer = HKReportProperty1.ColumnStartData
        Dim lastcolumn = DataStart + myKAMAssignmentList.Count
        For i = 0 To myKAMAssignmentList.Count - 1

            osheet.Cells(HKReportProperty1.RowStartDataI - 6, DataStart + i) = myKAMAssignmentList(i).sdasd
            osheet.Cells(HKReportProperty1.RowStartDataI - 5, DataStart + i) = myKAMAssignmentList(i).tefalsd
            osheet.Cells(HKReportProperty1.RowStartDataI - 4, DataStart + i) = myKAMAssignmentList(i).lagosd

            osheet.Cells(HKReportProperty1.RowStartDataI - 2, DataStart + i) = myKAMAssignmentList(i).username
            osheet.Cells(HKReportProperty1.RowStartDataI - 1, DataStart + i) = myKAMAssignmentList(i).mla


            osheet.Cells(1, DataStart + i) = "Net"
            osheet.Cells(1, DataStart + i).font.bold = True
            osheet.Cells(1, DataStart + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            osheet.Cells(1, DataStart + i).Interior.TintAndShade = 0.599993896298105
            osheet.Cells(6, DataStart + i) = "Gross"
            osheet.Cells(6, DataStart + i).font.bold = True
            osheet.Cells(6, DataStart + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
            osheet.Cells(6, DataStart + i).Interior.TintAndShade = 0.599993896298105

            osheet.Cells(7, DataStart + i).formulaR1C1 = String.Format("=R[-5]C/(1-R[5]C)")
            osheet.Cells(7, DataStart + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
            osheet.Cells(7, DataStart + i).Interior.TintAndShade = 0.599993896298105
            osheet.Cells(8, DataStart + i).formulaR1C1 = String.Format("=R[-5]C/(1-R[5]C)")
            osheet.Cells(8, DataStart + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
            osheet.Cells(8, DataStart + i).Interior.TintAndShade = 0.599993896298105
            osheet.Cells(9, DataStart + i).formulaR1C1 = String.Format("=R[-5]C/(1-R[5]C)")
            osheet.Cells(9, DataStart + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
            osheet.Cells(9, DataStart + i).Interior.TintAndShade = 0.599993896298105
            osheet.Cells(10, DataStart + i).formulaR1C1 = "=sum(R[-3]C:R[-1]C)"
            osheet.Cells(10, DataStart + i).font.bold = True
            osheet.Cells(10, DataStart + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
            osheet.Cells(10, DataStart + i).Interior.TintAndShade = 0.399975585192419

            osheet.Cells(5, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(R[9]C[{0}]:R[{1}]C[{0}],R[9]C:R[{1}]C)", -1 - i, e.lastRow - 5)
            osheet.Cells(5, 10 + i).font.bold = True
            osheet.Cells(5, 10 + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            osheet.Cells(5, 10 + i).Interior.TintAndShade = 0.399975585192419


            osheet.Cells(2, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R[12]C[{2}]:R[{0}]C[{2}]=""SDA""),R[12]C[{1}]:R[{0}]C[{1}],R[12]C:R[{0}]C)", e.lastRow - 2, -1 - i, -9 - i)
            osheet.Cells(2, 10 + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            osheet.Cells(2, 10 + i).Interior.TintAndShade = 0.599993896298105
            osheet.Cells(3, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R[11]C[{2}]:R[{0}]C[{2}]=""CKW - Tefal""),R[11]C[{1}]:R[{0}]C[{1}],R[11]C:R[{0}]C)", e.lastRow - 3, -1 - i, -9 - i)
            osheet.Cells(3, 10 + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            osheet.Cells(3, 10 + i).Interior.TintAndShade = 0.599993896298105
            osheet.Cells(4, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R[10]C[{2}]:R[{0}]C[{2}]=""CKW - Lago""),R[10]C[{1}]:R[{0}]C[{1}],R[10]C:R[{0}]C)", e.lastRow - 4, -1 - i, -9 - i)
            osheet.Cells(4, 10 + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            osheet.Cells(4, 10 + i).Interior.TintAndShade = 0.599993896298105
        Next

        Dim tmp = osheet.Name.Split("-")
        Dim mydate = String.Format(tmp(2).Replace(".", "-") & "-1")
        'osheet.Cells(1, 4) = "SD%"
        'osheet.Cells(1, 6) = "Target (Gross)"
        'osheet.Cells(1, 7) = "Target (Net)"
        'For Each drv As DataRow In HKParamDT.Rows
        '    If drv("period") = mydate Then
        '        Select Case drv("producttype")
        '            Case 1
        '                osheet.Cells(2, 4) = drv("sdpct")
        '                osheet.Cells(2, 6) = drv("targetgross")
        '                osheet.Cells(2, 7) = drv("targetnet")
        '            Case 2
        '                osheet.Cells(3, 4) = drv("sdpct")
        '                osheet.Cells(3, 6) = drv("targetgross")
        '                osheet.Cells(3, 7) = drv("targetnet")
        '            Case 3
        '                osheet.Cells(4, 4) = drv("sdpct")
        '                osheet.Cells(4, 6) = drv("targetgross")
        '                osheet.Cells(4, 7) = drv("targetnet")
        '        End Select

        '    End If
        'Next
        'osheet.Cells(2, 2) = "SDA"
        'osheet.Cells(3, 2) = "CKW TEFAL"
        'osheet.Cells(4, 2) = "CKW LAGO"
        'osheet.Cells(5, 2).value = "Total"
        osheet.Cells(1, 4) = "Net Sales"
        osheet.Cells(1, 4).font.bold = True

        osheet.Cells(6, 4) = "Gross Sales"
        osheet.Cells(6, 4).font.bold = True
        osheet.Cells(7, 2) = "Input by KAM"
        osheet.Cells(7, 2).font.bold = True
        osheet.Cells(9, 2) = "Difference"
        osheet.Cells(9, 2).font.bold = True

        osheet.Cells(2, 7) = "SDA"
        osheet.Cells(3, 7) = "CKW TEFAL"
        osheet.Cells(4, 7) = "CKW LAGO"

        osheet.Cells(7, 7) = "SDA"
        osheet.Cells(8, 7) = "CKW TEFAL"
        osheet.Cells(9, 7) = "CKW LAGO"

        osheet.Cells(12, 7) = "SDA"
        osheet.Cells(13, 7) = "CKW TEFAL"
        osheet.Cells(14, 7) = "CKW LAGO"

        osheet.Cells(2, lastcolumn + 1) = "SDA"
        osheet.Cells(3, lastcolumn + 1) = "CKW TEFAL"
        osheet.Cells(4, lastcolumn + 1) = "CKW LAGO"
        osheet.Cells(5, lastcolumn + 1).value = "Total"


        'osheet.Cells(5, 6).formulaR1C1 = "=SUM(R[-3]C:R[-1]C)" 'Sum Target Gross
        'osheet.Cells(5, 7).formulaR1C1 = "=SUM(R[-3]C:R[-1]C)" 'Sum Target Net


        osheet.Cells(2, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count) 'Total SDA
        osheet.Cells(2, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
        osheet.Cells(2, lastcolumn).Interior.TintAndShade = 0.599993896298105

        osheet.Cells(3, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count) 'Total TEFAL
        osheet.Cells(3, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
        osheet.Cells(3, lastcolumn).Interior.TintAndShade = 0.599993896298105
        osheet.Cells(4, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count) 'Total Lago
        osheet.Cells(4, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
        osheet.Cells(4, lastcolumn).Interior.TintAndShade = 0.599993896298105
        osheet.Cells(5, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count) 'Total All
        osheet.Cells(5, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
        osheet.Cells(5, lastcolumn).Interior.TintAndShade = 0.399975585192419
        osheet.Cells(5, lastcolumn).font.bold = True
        osheet.Cells(5, lastcolumn + 1).font.bold = True
        osheet.Cells(5, lastcolumn + 1).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
        osheet.Cells(5, lastcolumn + 1).Interior.TintAndShade = 0.399975585192419
        'osheet.Cells(7, 5).value = "Gross"

        'osheet.Cells(7, 6).formulaR1C1 = String.Format("=RC[{0}]-R[-5]C", myKAMAssignmentList.Count + 4)
        'osheet.Cells(8, 6).formulaR1C1 = String.Format("=RC[{0}]-R[-5]C", myKAMAssignmentList.Count + 4)
        'osheet.Cells(9, 6).formulaR1C1 = String.Format("=RC[{0}]-R[-5]C", myKAMAssignmentList.Count + 4)
        'osheet.Cells(10, 6).formulaR1C1 = "=sum(R[-3]C:R[-1]C)"

        'osheet.Cells(7, 2).value = "SDA"
        'osheet.Cells(8, 2).value = "CKW TEFAL"
        'osheet.Cells(9, 2).value = "CKW LAGO"
        'osheet.Cells(10, 2).value = "Total"
        osheet.Cells(7, lastcolumn + 1).value = "SDA"
        osheet.Cells(8, lastcolumn + 1).value = "CKW TEFAL"
        osheet.Cells(9, lastcolumn + 1).value = "CKW LAGO"
        osheet.Cells(10, lastcolumn + 1).value = "Total"
        osheet.Cells(10, lastcolumn).font.bold = True
        osheet.Cells(10, lastcolumn + 1).font.bold = True
        osheet.Cells(10, lastcolumn + 1).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        osheet.Cells(10, lastcolumn + 1).Interior.TintAndShade = 0.399975585192419

        osheet.Cells(7, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count)
        osheet.Cells(7, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        osheet.Cells(7, lastcolumn).Interior.TintAndShade = 0.599993896298105
        osheet.Cells(8, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count)
        osheet.Cells(8, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        osheet.Cells(8, lastcolumn).Interior.TintAndShade = 0.599993896298105
        osheet.Cells(9, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count)
        osheet.Cells(9, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        osheet.Cells(9, lastcolumn).Interior.TintAndShade = 0.599993896298105
        osheet.Cells(10, lastcolumn).formulaR1C1 = "=sum(R[-3]C:R[-1]C)"
        osheet.Cells(10, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        osheet.Cells(10, lastcolumn).Interior.TintAndShade = 0.399975585192419

        osheet.Cells(1, lastcolumn).value = "Total Net"
        osheet.Cells(1, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
        osheet.Cells(1, lastcolumn).Interior.TintAndShade = 0.599993896298105
        osheet.Cells(6, lastcolumn).value = "Total Gross"
        osheet.Cells(6, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        osheet.Cells(6, lastcolumn).Interior.TintAndShade = 0.599993896298105

        osheet.Cells(14, lastcolumn).formulaR1C1 = "=1-(R[-9]C/R[-4]C)"
        osheet.Cells(14, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
        osheet.Cells(14, lastcolumn).Interior.TintAndShade = 0.599993896298105

        osheet.Cells(HKReportProperty1.RowStartDataI, lastcolumn).value = "Total"
        osheet.Cells(7, 4).Interior.color = 65535

        osheet.Cells(9, 4).formulaR1C1 = String.Format("=IFERROR(R[1]c[{0}]-R[-2]C,)", lastcolumn - 4)

        osheet.Range("D2:D4").NumberFormat = "0.0%"
        osheet.Range("H14:I14").NumberFormat = "#,##0"
        osheet.Range("F2:F10").NumberFormat = "#,##0"
        osheet.Range("G2:G5").NumberFormat = "#,##0"
        osheet.Range("D7:D9").NumberFormat = "#,##0"

        osheet.Range(osheet.Cells(2, DataStart), osheet.Cells(10, lastcolumn)).NumberFormat = "#,##0"
        osheet.Range(osheet.Cells(12, DataStart), osheet.Cells(14, lastcolumn)).NumberFormat = "0.0%"
        'SDA
        'osheet.Cells(5, 10).FormulaR1C1 = String.Format("=SUMPRODUCT(R[9]C[-1]:R[{0}]C[-1],R[9]C:R[{0}]C)", e.lastRow)

        'Dim oRange As Excel.Range
        'oRange = osheet.Range(HKReportProperty1.RowStartdData)
        'oRange.Select()
        'osheet.Application.Selection.autofilter() 'should do twice because of the initial is fi
        'osheet.Application.Selection.autofilter()
        osheet.Cells.EntireColumn.AutoFit()
        osheet.Columns(lastcolumn).ColumnWidth = 15

        'osheet.Cells(10, 6).font.bold = True
        'osheet.Cells(10, 6).font.size = 14

        'osheet.Cells(10, 7).value = "Should Be ""0"""
        'osheet.Cells(10, 7).font.FontStyle = "Italic"
        'osheet.Cells(10, 7).font.bold = True
        'osheet.Cells(10, 7).font.size = 14
        'osheet.Columns(1).EntireColumn.Hidden = True
        osheet.Columns(3).EntireColumn.Hidden = True
        osheet.Columns(5).EntireColumn.Hidden = True
        osheet.Columns(8).EntireColumn.Hidden = True
        osheet.Columns(9).EntireColumn.Hidden = True
        osheet.Rows("11:14").EntireRow.Hidden = True
        osheet.Columns(7).columnwidth = 40
        'SubTotal Row
        For i = (HKReportProperty1.RowStartDataI + 1) To e.lastRow
            osheet.Cells(i, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count)
        Next
        osheet.Range(osheet.Cells(HKReportProperty1.RowStartDataI + 1, DataStart), osheet.Cells(HKReportProperty1.RowStartDataI + e.lastRow, lastcolumn + 5)).Locked = False
        osheet.Range("D7").Locked = False
        osheet.Range(osheet.Cells(1, 1), osheet.Cells(10, lastcolumn)).FormulaHidden = True
        Dim Identitiy As UserController = User.getIdentity
        If Not Identitiy.isAdmin Then
            osheet.Protect(DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="salesforecast", AllowFiltering:=True)
            osheet.EnableSelection = Excel.XlEnableSelection.xlNoRestrictions

        End If
        osheet.Cells(HKReportProperty1.RowStartDataI + 1, DataStart).select()
        oXL.ActiveWindow.FreezePanes = True

    End Sub
    Private Sub PivotTable()
        'Throw New NotImplementedException
    End Sub
End Class
