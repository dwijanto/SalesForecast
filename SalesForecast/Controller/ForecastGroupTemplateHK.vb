Imports System.Text
Imports Microsoft.Office.Interop
Public Class ForecastGroupTemplateHK
    Dim myPeriodRange As Dictionary(Of Integer, Date)
    Dim myKAMAssignment As New KAMAssignmentController
    Dim myKAMAssignmentList As List(Of KAMAssignmentModel)
    Dim HKReportProperty1 As HKReportProperty = HKReportProperty.getInstance
    'Dim HKReportProperty1 As HKReportProperty = HKGroupReportProperty.getInstance
    Public Sub Generate(myForm As Object, e As ForecastGroupTemplateHKEventArgs)

        Dim sqlstr As String = String.Empty

        Dim mysaveform As New SaveFileDialog
        Dim Identitiy As UserController = User.getIdentity
        Dim username As String()
        'If Not Identitiy.isAdmin Then
        username = Identitiy.username.Split("\")
        username(1) = username(1) & "_"
        'Else
        'ReDim username(1)
        'End If
        mysaveform.FileName = String.Format("{0}ForecastGroupHK{1:yyyyMMdd}.xlsx", username(1), Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            'Dim PeriodRange As New PeriodRange(e.startperiod, 12)
            Dim PeriodRange As New PeriodRange(e.startperiod, 13)
            myPeriodRange = PeriodRange.getPeriod
            Dim fieldList As New StringBuilder
            Dim ColumnList As New StringBuilder
            Dim TxtFieldList As New StringBuilder
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


            If e.blankTemplate Then
                'sqlstr = String.Format("with tx as (select * from crosstab('with tx as(select cmmf,txdate,sum(salesforecast) as salesforecast from sales.sfmlatxhk tx " &
                '                        "  where txdate >= ''{0:yyyy-MM-dd}''" &
                '                        " group by cmmf,txdate),mla as(select c.cmmf,tx.txdate,tx.salesforecast from sales.sfcmmf c left join tx on tx.cmmf = c.cmmf " &
                '                        " order by 1,2),alltx as (select * from mla)" &
                '                        " select * from alltx order by 1,2','select unnest(Array[{1}])::date') " &
                '                        " as ct (cmmf bigint, {2})) ,nsp as (select cmmf,nsp1, nsp2 from sales.sfcmmfnsp )" &
                '                        " select c.cmmf,c.reference,c.description ,c.familyname,sales.get_producttype(productlinegps,brand) as producttype,n.nsp1,n.nsp2 ,{3} from sales.sfcmmf c " &
                '                        " left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf ", myPeriodRange(0), fieldList.ToString, ColumnList.ToString, TxtFieldList.ToString)
                sqlstr = String.Format("with tx as (select * from crosstab('with tx as(select cmmf,txdate,sum(salesforecast) as salesforecast from sales.sfmlatxhk tx " &
                                        "  where txdate >= ''{0:yyyy-MM-dd}''" &
                                        " group by cmmf,txdate),mla as(select c.cmmf,tx.txdate,tx.salesforecast from sales.sfcmmf c left join tx on tx.cmmf = c.cmmf " &
                                        " order by 1,2),alltx as (select * from mla)" &
                                        " select * from alltx order by 1,2','select unnest(Array[{1}])::date') " &
                                        " as ct (cmmf bigint, {2})) ,nsp as (select cmmf,nsp1, nsp2 from sales.sfcmmfnsp )" &
                                        " select c.cmmf,c.reference,c.description ,c.familyname,sales.get_producttype(productlinegps,brand) as producttype,n.nsp1,n.nsp2 ,{3} from sales.sfcmmf c " &
                                        " left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf ", myPeriodRange(0), fieldList.ToString, ColumnList.ToString, TxtFieldList.ToString)
            Else
                'sqlstr = String.Format("with tx as (select * from crosstab('with tx as(select cmmf,txdate,sum(salesforecast) as salesforecast from sales.sfmlatxhk tx " &
                '                       " where txdate >= ''{0:yyyy-MM-dd}''" &
                '                                      " group by cmmf,txdate),mla as(select c.cmmf,tx.txdate,tx.salesforecast from sales.sfcmmf c left join tx on tx.cmmf = c.cmmf " &
                '                                      " order by 1,2),alltx as (select * from mla union all select cmmf,txdate,salesforecast from sales.sfforecastgrouptxhk where txdate >= ''{1:yyyy-MM-dd}'')" &
                '                                      " select * from alltx order by 1,2','select unnest(Array[{2}])::date') " &
                '                                      " as ct (cmmf bigint, {3})),nsp as (select cmmf,nsp1,nsp2 from sales.sfcmmfnsp )" &
                '                                      " select c.cmmf,c.reference,c.description ,c.familyname,sales.get_producttype(productlinegps,brand) as producttype,n.nsp1,n.nsp2 ,{4} from sales.sfcmmf c " &
                '                                      " left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf ", myPeriodRange(0), myPeriodRange(6), fieldList.ToString, ColumnList.ToString, TxtFieldList.ToString)

                If Identitiy.isAdmin Then
                    sqlstr = String.Format("with tx as (select * from crosstab('with tx as(select cmmf,txdate,sum(salesforecast) as salesforecast from sales.sfmlatxhk tx " &
                                       " where txdate >= ''{0:yyyy-MM-dd}''" &
                                                      " group by cmmf,txdate),mla as(select c.cmmf,tx.txdate,tx.salesforecast from sales.sfcmmf c left join tx on tx.cmmf = c.cmmf " &
                                                      " order by 1,2),alltx as (select * from mla union all select cmmf,txdate,salesforecast from sales.sfgrouptxhk where txdate >= ''{1:yyyy-MM-dd}'')" &
                                                      " select * from alltx order by 1,2','select unnest(Array[{2}])::date') " &
                                                      " as ct (cmmf bigint, {3})),nsp as (select cmmf,nsp1,nsp2 from sales.sfcmmfnsp )" &
                                                      " select c.cmmf,c.reference,c.description ,c.familyname,sales.get_producttype(productlinegps,brand) as producttype,n.nsp1,n.nsp2 ,{4} from sales.sfcmmf c " &
                                                      " left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf ", myPeriodRange(0), myPeriodRange(6), fieldList.ToString, ColumnList.ToString, TxtFieldList.ToString)


                Else
                    'Selected UserId
                    sqlstr = String.Format("with tx as (select * from crosstab('with tx as(select cmmf,txdate,sum(salesforecast) as salesforecast from sales.sfmlatxhk tx " &
                                       " where txdate >= ''{0:yyyy-MM-dd}''" &
                                                      " group by cmmf,txdate),mla as(select c.cmmf,tx.txdate,tx.salesforecast from sales.sfcmmf c left join tx on tx.cmmf = c.cmmf " &
                                                      " order by 1,2),alltx as (select * from mla union all select cmmf,txdate,salesforecast from sales.sfgrouptxhk where txdate >= ''{1:yyyy-MM-dd}'' and userid = {5})" &
                                                      " select * from alltx order by 1,2','select unnest(Array[{2}])::date') " &
                                                      " as ct (cmmf bigint, {3})),nsp as (select cmmf,nsp1,nsp2 from sales.sfcmmfnsp )" &
                                                      " select c.cmmf,c.reference,c.description ,c.familyname,sales.get_producttype(productlinegps,brand) as producttype,n.nsp1,n.nsp2 ,{4} from sales.sfcmmf c " &
                                                      " left join tx on tx.cmmf = c.cmmf left join nsp n on n.cmmf = c.cmmf ", myPeriodRange(0), myPeriodRange(6), fieldList.ToString, ColumnList.ToString, TxtFieldList.ToString, Identitiy.getId)


                End If

            End If

            Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\HKMLATemplate.xltx", HKReportProperty1.FGRowStartData)
            myreport.Run(myForm, e)

        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        Dim DataStart As Integer = HKReportProperty1.ColumnStartData
       
        'Dim oRange As Excel.Range
        'oRange = osheet.Range(HKReportProperty1.RowStartdData)
        'oRange.Select()
        'osheet.Application.Selection.autofilter() 'should do twice because of the initial is fi
        'osheet.Application.Selection.autofilter()
        'osheet.Name = String.Format("{0:yyyy.MM}", myPeriodRange(0))
        osheet.Name = String.Format("HK-FG-TOUS")
        osheet.Cells.EntireColumn.AutoFit()
    End Sub

    Private Sub PivotTable()
        'Throw New NotImplementedException
    End Sub
End Class
