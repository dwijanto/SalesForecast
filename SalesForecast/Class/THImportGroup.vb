﻿Imports System.Text

Public Class THImportGroup
    Implements HKImport

    Dim DATASB As StringBuilder

    Dim THReportProperty1 As THReportProperty = THReportProperty.getInstance
    Dim PostgresqlDBAdapter1 As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance

    Public Property ErrorMsg As String Implements HKImport.ErrorMsg
    Dim Period As Date

    Public Sub New(ByVal Filename As String, ByVal myperiod As Date)
        Period = myperiod
    End Sub


    Public Function Run(ByVal myForm As Object, ByVal myDoc As List(Of DocCSV)) As Boolean Implements HKImport.Run
        Dim myret As Boolean = False
        DATASB = New StringBuilder
        Dim myList As New List(Of THData)

        'KAMid 
        Dim kamDict As New Dictionary(Of Integer, String)
        Dim mlaDict As New Dictionary(Of Integer, String)
        Dim cardDict As New Dictionary(Of Integer, String)
        Dim myPeriodDict As New Dictionary(Of Integer, Date)
        Dim GroupController1 As New GroupController
        Dim mygroupdict As Dictionary(Of String, Integer) = GroupController1.Model.GetGroupDict(LocationEnum.Thailand)
        Dim kamlist As New StringBuilder
        For i = 0 To myDoc.Count - 1
            'import
            If kamlist.Length > 0 Then
                kamlist.Append(",")
            End If
            kamlist.Append(myDoc(i).KAMAssignmentID)
            Dim myrecord() As String
            Dim filename = String.Format("{0}\{1}.csv", myDoc(i).folder, myDoc(i).Name)
            Using objTFParser = New FileIO.TextFieldParser(filename)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0
                    myForm.ProgressReport(1, "Read Data")
                    Do Until .EndOfData
                        myrecord = .ReadFields
                        If count = 3 Then
                            'Get list of period                           
                            If myPeriodDict.Count = 0 Then
                                populatedict(myPeriodDict, myrecord)
                            End If
                        ElseIf count > 3 Then
                            'data
                            myList.Add(New THData With {.myrecord = myrecord,
                                                        .KAMAssignmentId = myDoc(i).KAMAssignmentID})
                        End If
                        count += 1
                    Loop
                End With
            End Using
        Next
        'buid Row
        For i = 0 To myList.Count - 1
            'Dim myperiod As Date = myList(i).Period
            Dim mycol As Integer
            'For j = 0 To 11
            For j = 0 To 12
                mycol = (THReportProperty1.ColumnStartData - 1) + j
                If myList(i).myrecord(mycol) <> "" Then

                    DATASB.Append(String.Format("'{0:yyyy-MM-dd}'", myPeriodDict(j)) & vbTab &
                                  myList(i).KAMAssignmentId & vbTab & myList(i).myrecord(0) & vbTab & myList(i).myrecord(mycol) & vbCrLf)
                End If
            Next
        Next

        If DATASB.Length > 0 Then
            'clean data for KAM
            'Dim NextYear = String.Format("{0}-{1:MM}-01", Period.Year + 1, Period)
            Dim NextYear = String.Format("{0:yyyy}-{0:MM}-01", Period.AddMonths(13))
            Dim sqlstr1 = String.Format("delete from sales.sfgrouptxth tx  where kamassignmentid  in ({0}) and tx.txdate >= '{1:yyyy-MM}-01' and tx.txdate < '{2}'", kamlist, Period, NextYear)
            PostgresqlDBAdapter1.ExecuteNonQuery(sqlstr1)

            'copy
            Dim sqlstr As String = "copy sales.sfgrouptxth(txdate,kamassignmentid,cmmf,salesforecast) from stdin with null as 'Null';"
            ErrorMsg = PostgresqlDBAdapter1.copy(sqlstr, DATASB.ToString, myret)
            If myret Then
                myForm.ProgressReport(1, "Done.")
            End If
        Else
            ErrorMsg = "Nothing to import."
        End If

        Return myret
    End Function

    Private Sub populatedict(mydict As Dictionary(Of Integer, Date), myrecord As String())
        'For j = 0 To 11
        For j = 0 To 12
            Dim tmpdate As Date = CDate(String.Format("{0}-{1}-1", myrecord(THReportProperty1.ColumnStartData - 1 + j).ToString.Substring(0, 4), myrecord(THReportProperty1.ColumnStartData - 1 + j).ToString.Substring(4, 2)))
            mydict.Add(j, tmpdate)
        Next
    End Sub
End Class
