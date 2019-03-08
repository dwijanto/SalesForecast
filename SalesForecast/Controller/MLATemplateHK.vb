Imports System.Text
Imports Microsoft.Office.Interop

Public Class MLATemplateHK
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
                    'sqlstr = String.Format("with tx as (select * from crosstab('" &
                    '         " with ka as (select distinct kam.username,mc.mla,mc.cardname,trim(mc.mla) || '' - '' || mc.cardname as assignment  from sales.sfkam kam  left join sales.sfkamassignment ka on ka.kam = kam.username " &
                    '         " left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid where kam.username = ''{0}'' order by mc.mla) ,kam as (select 1 as id,row_number() over (order by ka.mla,ka.cardname) as recid,* from ka),  cmmf as(select  1 as id,* from sales.sfcmmf cmmf order by cmmf),tx as(select cmmf.cmmf,kam.recid,null::integer from cmmf" &
                    '         " left join kam on kam.id = cmmf.id left join sales.sfmlatxhk tx on tx.cmmf = cmmf.cmmf and tx.kam = kam.username and tx.mla = kam.mla and tx.cardname = kam.cardname and tx.txdate = ''{1:yyyy-MM-dd}''" &
                    '         " order by cmmf,kam.mla) select * from tx;','select m from generate_series(1,{2})m') as (cmmf bigint,{3})), " &
                    '         " nsp as (select cmmf,max(nsp1) as nsp1,max(nsp2) as nsp2 from sales.sfcmmfmlansp group by cmmf)" &
                    '         " select c.cmmf,c.reference,c.productdesc ,sales.get_producttype(productlinegps,brand) as producttype,n.nsp1,n.nsp2 ,{4} from sales.sfcmmf c left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf ", username, myPeriodRange(i), myKAMAssignmentList.Count, fieldList.ToString, txfieldlist.ToString)
                    'sqlstr = String.Format("with tx as (select * from crosstab('" &
                    '       " with ka as (select distinct kam.username,mc.mla,mc.cardname,trim(mc.mla) || '' - '' || mc.cardname as assignment  from sales.sfkam kam  " &
                    '       " left join sales.sfkamassignment ka on ka.kam = kam.username  	left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid where kam.username = ''{0}'' " &
                    '       " order by mc.mla) , kam as (select 1 as id,row_number() over (order by ka.mla,ka.cardname) as recid,* from ka), " &
                    '       " cmmf as(select  1 as id,* from sales.sfcmmf cmmf order by cmmf), " &
                    '       " sa as (select tx.txdate,null::bigint as salesforecast,tx.cmmfkamassignmentid,cka.cmmf,ka.kam,mc.mla,mc.cardname from sales.sfmlatxhk tx " &
                    '       " left join sales.sfcmmfkamassignment cka on cka.id = tx.cmmfkamassignmentid" &
                    '       " left join sales.sfkamassignment ka on ka.id = cka.kamassignmentid left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid " &
                    '       " where tx.txdate = ''{1:yyyy-MM-dd}'' and ka.kam = ''{0}'' ),tx as(select cmmf.cmmf,kam.recid,sa.salesforecast from cmmf " &
                    '       " left join kam on kam.id = cmmf.id  left join sa on sa.cmmf = cmmf.cmmf and sa.kam = kam.username and sa.mla = kam.mla and sa.cardname = kam.cardname 	order by cmmf,kam.mla) select * from tx;','select m from generate_series(1,{2})m') as (cmmf bigint,{3})), " &
                    '       " nsp as (select cmmf,max(nsp1) as nsp1,max(nsp2) as nsp2 from sales.sfcmmfmlansp group by cmmf)" &
                    '       " select c.cmmf,c.reference,c.productdesc ,sales.get_producttype(productlinegps,brand) as producttype,n.nsp1,n.nsp2 ,{4} " &
                    '       " from sales.sfcmmf c left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf ", username, myPeriodRange(i), myKAMAssignmentList.Count, fieldList.ToString, txfieldlist.ToString)
                    'sqlstr = String.Format("with tx as (select * from crosstab('" &
                    '       " with ka as (select distinct kam.username,mc.mla,mc.cardname,trim(mc.mla) || '' - '' || mc.cardname as assignment  from sales.sfkam kam  " &
                    '       " left join sales.sfkamassignment ka on ka.kam = kam.username  	left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid where kam.username = ''{0}'' " &
                    '       " order by mc.mla) , kam as (select 1 as id,row_number() over (order by ka.mla,ka.cardname) as recid,* from ka), " &
                    '       " cmmf as(select  1 as id,* from sales.sfcmmf cmmf order by cmmf), " &
                    '       " sa as (select tx.txdate,null::bigint as salesforecast,tx.cmmfkamassignmentid,cka.cmmf,ka.kam,mc.mla,mc.cardname from sales.sfmlatxhk tx " &
                    '       " left join sales.sfcmmfkamassignment cka on cka.id = tx.cmmfkamassignmentid" &
                    '       " left join sales.sfkamassignment ka on ka.id = cka.kamassignmentid left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid " &
                    '       " where tx.txdate = ''{1:yyyy-MM-dd}'' and ka.kam = ''{0}'' ),tx as(select cmmf.cmmf,kam.recid,sa.salesforecast from cmmf " &
                    '       " left join kam on kam.id = cmmf.id  left join sa on sa.cmmf = cmmf.cmmf and sa.kam = kam.username and sa.mla = kam.mla and sa.cardname = kam.cardname 	order by cmmf,kam.mla) select * from tx;','select m from generate_series(1,{2})m') as (cmmf bigint,{3})), " &
                    '       " nsp as (select cmmf, nsp1,nsp2 from sales.sfcmmfnsp)" &
                    '       " select sales.get_producttype(productlinegpsid,brand) as producttype ,c.brand,c.productlinegps,c.reference,c.familyname,c.cmmf,c.description ,n.nsp1 as ""NSP(USD)"",n.nsp2 as ""NSP(HKD)"" ,{4} " &
                    '       " from sales.sfcmmf c left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf where date_part('Month',c.activedate) <= date_part('Month', '{1:yyyy-MM-dd}'::date) and date_part('Year',c.activedate) <= date_part('Year','{1:yyyy-MM-dd}'::date)  order by producttype,c.brand,c.reference ", username, myPeriodRange(i), myKAMAssignmentList.Count, fieldList.ToString, txfieldlist.ToString)
                    ''order by producttype,c.brand,c.productlinegps,c.familyname,c.reference
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
                           " select sales.get_producttype(productlinegpsid,brand) as producttype ,c.brand,c.productlinegps,c.reference,c.familyname,c.cmmf,c.description ,n.nsp1 as ""NSP(USD)"",n.nsp2 as ""NSP(HKD)"" ,{4} " &
                           " from sales.sfcmmf c left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf where to_char(activedate,'YYYYMM')::int <= to_char('{1:yyyy-MM-dd}'::date,'YYYYMM')::int  order by producttype,c.brand,c.reference ", username, myPeriodRange(i), myKAMAssignmentList.Count, fieldList.ToString, txfieldlist.ToString)
                Else
                    'sqlstr = String.Format("with tx as (select * from crosstab('" &
                    '         " with ka as (select distinct kam.username,mc.mla,mc.cardname,trim(mc.mla) || '' - '' || mc.cardname as assignment  from sales.sfkam kam  left join sales.sfkamassignment ka on ka.kam = kam.username " &
                    '         " left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid where kam.username = ''{0}'' order by mc.mla) ,kam as (select 1 as id,row_number() over (order by ka.mla,ka.cardname) as recid,* from ka),  cmmf as(select  1 as id,* from sales.sfcmmf cmmf order by cmmf),tx as(select cmmf.cmmf,kam.recid,tx.salesforecast from cmmf" &
                    '         " left join kam on kam.id = cmmf.id left join sales.sfmlatxhk tx on tx.cmmf = cmmf.cmmf and tx.kam = kam.username and tx.mla = kam.mla and tx.cardname = kam.cardname and tx.txdate = ''{1:yyyy-MM-dd}''" &
                    '         " order by cmmf,kam.mla) select * from tx;','select m from generate_series(1,{2})m') as (cmmf bigint,{3})), " &
                    '         " nsp as (select cmmf,max(nsp1) as nsp1,max(nsp2) as nsp2 from sales.sfcmmfmlansp group by cmmf)" &
                    '         " select c.cmmf,c.reference,c.productdesc ,sales.get_producttype(productlinegps,brand) as producttype,n.nsp1,n.nsp2 ,{4} from sales.sfcmmf c left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf ", username, myPeriodRange(i), myKAMAssignmentList.Count, fieldList.ToString, txfieldlist.ToString)
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
                    '        " nsp as (select cmmf,max(nsp1) as nsp1,max(nsp2) as nsp2 from sales.sfcmmfmlansp group by cmmf)" &
                    '        " select c.cmmf,c.reference,c.productdesc ,sales.get_producttype(productlinegps,brand) as producttype,n.nsp1,n.nsp2 ,{4} " &
                    '        " from sales.sfcmmf c left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf ", username, myPeriodRange(i), myKAMAssignmentList.Count, fieldList.ToString, txfieldlist.ToString)
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
                    '        " select sales.get_producttype(productlinegpsid,brand) as producttype ,c.brand,c.productlinegps,c.reference,c.familyname,c.cmmf,c.description ,n.nsp1 as ""NSP(USD)"",n.nsp2 as ""NSP(HKD)"" ,{4} " &
                    '        " from sales.sfcmmf c left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf where date_part('Month',c.activedate) <= date_part('Month', '{1:yyyy-MM-dd}'::date) and date_part('Year',c.activedate) <= date_part('Year','{1:yyyy-MM-dd}'::date)  order by producttype,c.brand,c.reference  ", username, myPeriodRange(i), myKAMAssignmentList.Count, fieldList.ToString, txfieldlist.ToString)
                    ''order by producttype,c.brand,c.productlinegps,c.familyname,c.reference
                    sqlstr = String.Format("with tx as (select * from crosstab('" &
                            " with ka as (select distinct kam.username,mc.mla,mc.cardname,trim(mc.mla) || '' - '' || mc.cardname as assignment  from sales.sfkam kam  " &
                            " left join sales.sfkamassignment ka on ka.kam = kam.username  	left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid where kam.username = ''{0}'' " &
                            " order by mc.mla) , kam as (select 1 as id,row_number() over (order by ka.mla,ka.cardname) as recid,* from ka), " &
                            " cmmf as(select  1 as id,* from sales.sfcmmf cmmf order by cmmf), " &
                            " sa as (select tx.txdate,tx.salesforecast,tx.cmmfkamassignmentid,cka.cmmf,ka.kam,mc.mla,mc.cardname from sales.sfmlatxhk tx " &
                            " left join sales.sfcmmfkamassignment cka on cka.id = tx.cmmfkamassignmentid" &
                            " left join sales.sfkamassignment ka on ka.id = cka.kamassignmentid left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid " &
                            " where tx.txdate = ''{1:yyyy-MM-dd}'' and ka.kam = ''{0}'' ),tx as(select cmmf.cmmf,kam.recid,sa.salesforecast from cmmf " &
                            " left join kam on kam.id = cmmf.id  left join sa on sa.cmmf = cmmf.cmmf and sa.kam = kam.username and sa.mla = kam.mla and sa.cardname = kam.cardname 	order by cmmf,kam.mla) select * from tx;','select m from generate_series(1,{2})m') as (cmmf bigint,{3})), " &
                            " nsp as (select cmmf, nsp1,nsp2 from sales.sfcmmfnsp )" &
                            " select sales.get_producttype(productlinegpsid,brand) as producttype ,c.brand,c.productlinegps,c.reference,c.familyname,c.cmmf,c.description ,n.nsp1 as ""NSP(USD)"",n.nsp2 as ""NSP(HKD)"" ,{4} " &
                            " from sales.sfcmmf c left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf where to_char(activedate,'YYYYMM')::int <= to_char('{1:yyyy-MM-dd}'::date,'YYYYMM')::int  order by producttype,c.brand,c.reference  ", username, myPeriodRange(i), myKAMAssignmentList.Count, fieldList.ToString, txfieldlist.ToString)
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
        Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        Dim DataStart As Integer = HKReportProperty1.ColumnStartData
        For i = 0 To myKAMAssignmentList.Count - 1
            osheet.Cells(HKReportProperty1.RowStartDataI - 2, DataStart + i) = myKAMAssignmentList(i).username
            osheet.Cells(HKReportProperty1.RowStartDataI - 1, DataStart + i) = myKAMAssignmentList(i).mla
            osheet.Cells(5, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(R[9]C[{0}]:R[{1}]C[{0}],R[9]C:R[{1}]C)", -1 - i, e.lastRow - 5)
            osheet.Cells(2, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R[12]C[{2}]:R[{0}]C[{2}]=""SDA""),R[12]C[{1}]:R[{0}]C[{1}],R[12]C:R[{0}]C)", e.lastRow - 2, -1 - i, -9 - i)
            osheet.Cells(3, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R[11]C[{2}]:R[{0}]C[{2}]=""CKW - Tefal""),R[11]C[{1}]:R[{0}]C[{1}],R[11]C:R[{0}]C)", e.lastRow - 3, -1 - i, -9 - i)
            osheet.Cells(4, 10 + i).FormulaR1C1 = String.Format("=SUMPRODUCT(--(R[10]C[{2}]:R[{0}]C[{2}]=""CKW - Lago""),R[10]C[{1}]:R[{0}]C[{1}],R[10]C:R[{0}]C)", e.lastRow - 4, -1 - i, -9 - i)
        Next
        Dim tmp = osheet.Name.Split("-")
        Dim mydate = String.Format(tmp(2).Replace(".", "-") & "-1")
        osheet.Cells(1, 4) = "SD%"
        osheet.Cells(1, 6) = "Target (Gross)"
        osheet.Cells(1, 7) = "Target (Net)"
        For Each drv As DataRow In HKParamDT.Rows
            If drv("period") = mydate Then
                Select Case drv("producttype")
                    Case 1
                        osheet.Cells(2, 4) = drv("sdpct")
                        osheet.Cells(2, 6) = drv("targetgross")
                    Case 2
                        osheet.Cells(3, 4) = drv("sdpct")
                        osheet.Cells(3, 6) = drv("targetgross")
                    Case 3
                        osheet.Cells(4, 4) = drv("sdpct")
                        osheet.Cells(4, 6) = drv("targetgross")
                End Select

            End If
        Next
        osheet.Cells(2, 5) = "SDA"
        osheet.Cells(3, 5) = "CKW TEFAL"
        osheet.Cells(4, 5) = "CKW LAGO"

        osheet.Cells(2, 7).formulaR1C1 = "=RC[-1]*(1-RC[-3])"
        osheet.Cells(3, 7).formulaR1C1 = "=RC[-1]*(1-RC[-3])"
        osheet.Cells(4, 7).formulaR1C1 = "=RC[-1]*(1-RC[-3])"

        osheet.Cells(5, 6).formulaR1C1 = "=SUM(R[-3]C:R[-1]C)" 'Sum Target Gross
        osheet.Cells(5, 7).formulaR1C1 = "=SUM(R[-3]C:R[-1]C)" 'Sum Target Net


        osheet.Cells(2, 9).formulaR1C1 = String.Format("=SUM(RC[1]:RC[{0}])", myKAMAssignmentList.Count) 'Total SDA
        osheet.Cells(3, 9).formulaR1C1 = String.Format("=SUM(RC[1]:RC[{0}])", myKAMAssignmentList.Count) 'Total TEFAL
        osheet.Cells(4, 9).formulaR1C1 = String.Format("=SUM(RC[1]:RC[{0}])", myKAMAssignmentList.Count) 'Total Lago
        osheet.Cells(5, 8).value = "Total"
        osheet.Cells(5, 9).formulaR1C1 = String.Format("=SUM(RC[1]:RC[{0}])", myKAMAssignmentList.Count) 'Total All

        osheet.Cells(7, 5).value = "Gross"
        osheet.Cells(7, 6).formulaR1C1 = "=RC[3]-R[-5]C"
        osheet.Cells(8, 6).formulaR1C1 = "=RC[3]-R[-5]C"
        osheet.Cells(9, 6).formulaR1C1 = "=RC[3]-R[-5]C"
        osheet.Cells(10, 6).formulaR1C1 = "=sum(R[-3]C:R[-1]C)"

        osheet.Cells(7, 8).value = "SDA"
        osheet.Cells(8, 8).value = "CKW TEFAL"
        osheet.Cells(9, 8).value = "CKW LAGO"
        osheet.Cells(10, 8).value = "Gross"

        osheet.Cells(7, 9).formulaR1C1 = "=R[-5]C/(1-R[-5]C[-5])"
        osheet.Cells(8, 9).formulaR1C1 = "=R[-5]C/(1-R[-5]C[-5])"
        osheet.Cells(9, 9).formulaR1C1 = "=R[-5]c/(1-R[-5]C[-5])"
        osheet.Cells(10, 9).formulaR1C1 = "=sum(R[-3]C:R[-1]C)"
        osheet.Range("D2:D4").NumberFormat = "0%"
        osheet.Range("H14:I14").NumberFormat = "#,##0.00"
        'SDA
        'osheet.Cells(5, 10).FormulaR1C1 = String.Format("=SUMPRODUCT(R[9]C[-1]:R[{0}]C[-1],R[9]C:R[{0}]C)", e.lastRow)

        'Dim oRange As Excel.Range
        'oRange = osheet.Range(HKReportProperty1.RowStartdData)
        'oRange.Select()
        'osheet.Application.Selection.autofilter() 'should do twice because of the initial is fi
        'osheet.Application.Selection.autofilter()
        osheet.Cells.EntireColumn.AutoFit()
    End Sub
    Private Sub FormattingReport(ByRef sender As Object, ByRef e As ExportToExcelFileEventArgs)
        Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        Dim owb As Excel.Workbook = osheet.Parent
        Dim oXL As Excel.Application = owb.Parent
        Dim DataStart As Integer = HKReportProperty1.ColumnStartData
        Dim lastcolumn = DataStart + myKAMAssignmentList.Count
        For i = 0 To myKAMAssignmentList.Count - 1
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

            osheet.Cells(7, DataStart + i).formulaR1C1 = String.Format("=R[-5]C/(1-R[-5]C[-{0}])", i + 6)
            osheet.Cells(7, DataStart + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
            osheet.Cells(7, DataStart + i).Interior.TintAndShade = 0.599993896298105
            osheet.Cells(8, DataStart + i).formulaR1C1 = String.Format("=R[-5]C/(1-R[-5]C[-{0}])", i + 6)
            osheet.Cells(8, DataStart + i).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
            osheet.Cells(8, DataStart + i).Interior.TintAndShade = 0.599993896298105
            osheet.Cells(9, DataStart + i).formulaR1C1 = String.Format("=R[-5]C/(1-R[-5]C[-{0}])", i + 6)
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
        osheet.Cells(1, 4) = "SD%"
        osheet.Cells(1, 6) = "Target (Gross)"
        osheet.Cells(1, 7) = "Target (Net)"
        For Each drv As DataRow In HKParamDT.Rows
            If drv("period") = mydate Then
                Select Case drv("producttype")
                    Case 1
                        osheet.Cells(2, 4) = drv("sdpct")
                        osheet.Cells(2, 6) = drv("targetgross")
                        osheet.Cells(2, 7) = drv("targetnet")
                    Case 2
                        osheet.Cells(3, 4) = drv("sdpct")
                        osheet.Cells(3, 6) = drv("targetgross")
                        osheet.Cells(3, 7) = drv("targetnet")
                    Case 3
                        osheet.Cells(4, 4) = drv("sdpct")
                        osheet.Cells(4, 6) = drv("targetgross")
                        osheet.Cells(4, 7) = drv("targetnet")
                End Select

            End If
        Next
        osheet.Cells(2, 2) = "SDA"
        osheet.Cells(3, 2) = "CKW TEFAL"
        osheet.Cells(4, 2) = "CKW LAGO"
        osheet.Cells(5, 2).value = "Total"
        osheet.Cells(2, lastcolumn + 1) = "SDA"
        osheet.Cells(3, lastcolumn + 1) = "CKW TEFAL"
        osheet.Cells(4, lastcolumn + 1) = "CKW LAGO"
        osheet.Cells(5, lastcolumn + 1).value = "Total"

        'osheet.Cells(2, 7).formulaR1C1 = "=RC[-1]*(1-RC[-3])"
        'osheet.Cells(3, 7).formulaR1C1 = "=RC[-1]*(1-RC[-3])"
        'osheet.Cells(4, 7).formulaR1C1 = "=RC[-1]*(1-RC[-3])"

        osheet.Cells(5, 6).formulaR1C1 = "=SUM(R[-3]C:R[-1]C)" 'Sum Target Gross
        osheet.Cells(5, 7).formulaR1C1 = "=SUM(R[-3]C:R[-1]C)" 'Sum Target Net


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

        osheet.Cells(7, 6).formulaR1C1 = String.Format("=RC[{0}]-R[-5]C", myKAMAssignmentList.Count + 4)
        osheet.Cells(8, 6).formulaR1C1 = String.Format("=RC[{0}]-R[-5]C", myKAMAssignmentList.Count + 4)
        osheet.Cells(9, 6).formulaR1C1 = String.Format("=RC[{0}]-R[-5]C", myKAMAssignmentList.Count + 4)
        osheet.Cells(10, 6).formulaR1C1 = "=sum(R[-3]C:R[-1]C)"

        osheet.Cells(7, 2).value = "SDA"
        osheet.Cells(8, 2).value = "CKW TEFAL"
        osheet.Cells(9, 2).value = "CKW LAGO"
        osheet.Cells(10, 2).value = "Total"
        osheet.Cells(7, lastcolumn + 1).value = "SDA"
        osheet.Cells(8, lastcolumn + 1).value = "CKW TEFAL"
        osheet.Cells(9, lastcolumn + 1).value = "CKW LAGO"
        osheet.Cells(10, lastcolumn + 1).value = "Total"
        osheet.Cells(10, lastcolumn).font.bold = True
        osheet.Cells(10, lastcolumn + 1).font.bold = True
        osheet.Cells(10, lastcolumn + 1).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        osheet.Cells(10, lastcolumn + 1).Interior.TintAndShade = 0.399975585192419

        osheet.Cells(7, lastcolumn).formulaR1C1 = String.Format("=R[-5]C/(1-R[-5]C[-{0}])", myKAMAssignmentList.Count + 6)
        osheet.Cells(7, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        osheet.Cells(7, lastcolumn).Interior.TintAndShade = 0.599993896298105
        osheet.Cells(8, lastcolumn).formulaR1C1 = String.Format("=R[-5]C/(1-R[-5]C[-{0}])", myKAMAssignmentList.Count + 6)
        osheet.Cells(8, lastcolumn).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        osheet.Cells(8, lastcolumn).Interior.TintAndShade = 0.599993896298105
        osheet.Cells(9, lastcolumn).formulaR1C1 = String.Format("=R[-5]C/(1-R[-5]C[-{0}])", myKAMAssignmentList.Count + 6)
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

        osheet.Cells(13, lastcolumn).value = "Total"
        osheet.Cells(10, 6).Interior.color = 65535
        osheet.Range("D2:D4").NumberFormat = "0.0%"
        osheet.Range("H14:I14").NumberFormat = "#,##0"
        osheet.Range("F2:F10").NumberFormat = "#,##0"
        osheet.Range("G2:G5").NumberFormat = "#,##0"

        osheet.Range(osheet.Cells(2, DataStart), osheet.Cells(10, lastcolumn)).NumberFormat = "#,##0"
        'SDA
        'osheet.Cells(5, 10).FormulaR1C1 = String.Format("=SUMPRODUCT(R[9]C[-1]:R[{0}]C[-1],R[9]C:R[{0}]C)", e.lastRow)

        'Dim oRange As Excel.Range
        'oRange = osheet.Range(HKReportProperty1.RowStartdData)
        'oRange.Select()
        'osheet.Application.Selection.autofilter() 'should do twice because of the initial is fi
        'osheet.Application.Selection.autofilter()
        osheet.Cells.EntireColumn.AutoFit()
        osheet.Columns(lastcolumn).ColumnWidth = 15
        osheet.Cells(10, 6).font.bold = True
        osheet.Cells(10, 6).font.size = 14

        osheet.Cells(10, 7).value = "Should Be ""0"""
        osheet.Cells(10, 7).font.FontStyle = "Italic"
        osheet.Cells(10, 7).font.bold = True
        osheet.Cells(10, 7).font.size = 14
        'osheet.Columns(1).EntireColumn.Hidden = True
        osheet.Columns(3).EntireColumn.Hidden = True
        osheet.Columns(5).EntireColumn.Hidden = True
        osheet.Columns(8).EntireColumn.Hidden = True
        osheet.Columns(9).EntireColumn.Hidden = True

        osheet.Columns(7).columnwidth = 40
        'SubTotal Row
        For i = (HKReportProperty1.RowStartDataI + 1) To e.lastRow
            osheet.Cells(i, lastcolumn).formulaR1C1 = String.Format("=SUM(RC[-{0}]:RC[-1])", myKAMAssignmentList.Count)
        Next
        osheet.Range(osheet.Cells(HKReportProperty1.RowStartDataI + 1, DataStart), osheet.Cells(HKReportProperty1.RowStartDataI + e.lastRow, lastcolumn + 5)).Locked = False
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
