Imports System.Text

Public Class HKImportMLA
    Implements HKImport

    Dim KAMController1 As KAMController
    Dim MLACardController1 As MLACardController
    Dim KAMAssignmentController1 As KAMAssignmentController
    Dim CMMFKAMAssignmentController1 As CMMFKAMAssignmentController

    Dim KAMDT As DataTable
    Dim MLACardDT As DataTable
    Dim KAMAssignmentDT As DataTable
    Dim CMMFKAMAssignmentDT As DataTable
    Dim KAMSB As StringBuilder
    Dim MLACardSB As StringBuilder
    Dim KAMAssignmentSB As StringBuilder
    Dim CMMFKAMAssignmentSB As StringBuilder

    Dim DATASB As StringBuilder
    'Dim ReportProperty1 = HKReportProperty.getInstance
    Dim HKReportProperty1 As HKReportProperty = HKReportProperty.getInstance
    Dim PostgresqlDBAdapter1 As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance
    'Dim KAMAssignmentSeq As Long
    'Dim MLACARDSeq As Long
    'Dim CMMFKAMAssignmentSeq As Long
    Dim myPeriod As Date
    Public Property ErrorMsg As String Implements HKImport.ErrorMsg

    Public Sub New(ByVal Filename As String, ByVal myPeriod As Date)
        Me.myPeriod = myPeriod
        KAMController1 = New KAMController
        KAMController1.loaddata()
        MLACardController1 = New MLACardController
        MLACardController1.loaddata()
        KAMAssignmentController1 = New KAMAssignmentController
        KAMAssignmentController1.loaddata()
        CMMFKAMAssignmentController1 = New CMMFKAMAssignmentController
        CMMFKAMAssignmentController1.loaddata()

        KAMSB = New StringBuilder
        MLACardSB = New StringBuilder
        KAMAssignmentSB = New StringBuilder
        CMMFKAMAssignmentSB = New StringBuilder
        DATASB = New StringBuilder

        KAMDT = KAMController1.GetTable
        MLACardDT = MLACardController1.GetTable
        KAMAssignmentDT = KAMAssignmentController1.GetTable
        CMMFKAMAssignmentDT = CMMFKAMAssignmentController1.GetTable



    End Sub


    Public Function Run(ByVal myForm As Object, ByVal myDoc As List(Of DocCSV)) As Boolean Implements HKImport.Run
        Dim myret As Boolean = False
        DATASB = New StringBuilder
        Dim myList As New List(Of HKData)

        'KAMid 
        Dim kamDict As New Dictionary(Of Integer, String)
        Dim mlaDict As New Dictionary(Of Integer, String)
        Dim cardDict As New Dictionary(Of Integer, String)
        For i = 0 To myDoc.Count - 1
            'import
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
                        If count = 16 Then '10 Then
                            'KAM                           
                            If kamDict.Count = 0 Then
                                populatedict(kamDict, myrecord)
                            End If
                        ElseIf count = 17 Then '11 Then
                            If mlaDict.Count = 0 Then
                                populatedict(mlaDict, myrecord)
                            End If
                            'MLA
                        ElseIf count = 18 Then '12 Then
                            'cardname
                            If cardDict.Count = 0 Then
                                populatedict(cardDict, myrecord)
                            End If
                        ElseIf count > 18 Then '12 Then
                            'data
                            myList.Add(New HKData With {.Period = myDoc(i).Period, .myrecord = myrecord})
                        End If
                        count += 1
                    Loop
                End With
            End Using
        Next
        'buid Row
        myForm.ProgressReport(1, "Build Data...")
        For i = 0 To myList.Count - 1
            Dim myperiod As Date = myList(i).Period
            Dim mycol As Integer
            'For j = 0 To myList(i).myrecord.Count - 1 ' - HKReportProperty1.ColumnStartData
            'Dim mylastrecord = myList(i).myrecord.Count - 6
            'Dim mylastrecord = myList(i).myrecord.Count - 6
            'For j = 0 To (myList(i).myrecord.Count - 1) - (HKReportProperty1.ColumnStartData - 1) ' - HKReportProperty1.ColumnStartData
            'For j = 0 To mylastrecord - (HKReportProperty1.ColumnStartData) ' - HKReportProperty1.ColumnStartData

            For j = 0 To kamDict.Count - 1 ' - HKReportProperty1.ColumnStartData
                'If mycol < (myList(i).myrecord.Count - 1) - (HKReportProperty1.ColumnStartData - 1) Then
                mycol = (HKReportProperty1.ColumnStartData - 1) + j
                If myList(i).myrecord(mycol) <> "" Then
                    'get mlacard
                    Dim mykey(1) As Object
                    mykey(0) = mlaDict(j)
                    mykey(1) = cardDict(j)
                    Dim mlacardid As Integer
                    Dim result = MLACardDT.Rows.Find(mykey)
                    If IsNothing(result) Then
                        'mlacardid = 1
                        'insert newrecord
                        Dim mlacardmodel1 As New MLACardModel
                        mlacardmodel1.mla = mykey(0)
                        mlacardmodel1.cardname = mykey(1)
                        mlacardid = MLACardController1.Model.Add(mlacardmodel1)



                        Dim dr As DataRow = MLACardDT.NewRow
                        dr.Item("id") = mlacardid
                        dr.Item("mla") = mykey(0)
                        dr.Item("cardname") = mykey(1)
                        MLACardDT.Rows.Add(dr)

                        'add new StringBuilder
                    Else
                        mlacardid = result.Item("id")
                    End If

                    'get kam Assignment
                    Dim mykey2(1) As Object
                    mykey2(0) = mlacardid
                    mykey2(1) = kamDict(j)
                    Dim kamassignmentid As Integer
                    result = KAMAssignmentDT.Rows.Find(mykey2)
                    If IsNothing(result) Then
                        'mlacardid = 1
                        'insert newrecord
                        'add new StringBuilder
                        Dim KAMAssignmentModel1 As New KAMAssignmentModel
                        KAMAssignmentModel1.mlacardnameid = mlacardid
                        KAMAssignmentModel1.kam = kamDict(j)


                        kamassignmentid = KAMAssignmentController1.Model.Add(KAMAssignmentModel1)
                        Dim dr As DataRow = KAMAssignmentDT.NewRow
                        dr.Item("id") = kamassignmentid
                        dr.Item("mlacardnameid") = mlacardid
                        dr.Item("kam") = kamDict(j)

                        KAMAssignmentDT.Rows.Add(dr)
                    Else
                        kamassignmentid = result.Item("id")
                    End If

                    'get cmmfkamassignment
                    Dim mykey3(1) As Object
                    mykey3(0) = kamassignmentid
                    mykey3(1) = myList(i).myrecord(HKReportProperty1.ColumnStartKey - 1)
                    Dim Cmmfkamassignmentid As Integer
                    result = CMMFKAMAssignmentDT.Rows.Find(mykey3)
                    If IsNothing(result) Then
                        'mlacardid = 1
                        'insert newrecord
                        'add new StringBuilder
                        Dim cmmfkamassignmentmodel1 As New CMMFKAMAssignmentModel
                        cmmfkamassignmentmodel1.cmmf = mykey3(1)
                        cmmfkamassignmentmodel1.kamassignmentid = kamassignmentid
                        Cmmfkamassignmentid = CMMFKAMAssignmentController1.Model.Add(cmmfkamassignmentmodel1)

                        Dim dr As DataRow = CMMFKAMAssignmentDT.NewRow
                        dr.Item("id") = Cmmfkamassignmentid
                        dr.Item("kamassignmentid") = kamassignmentid
                        dr.Item("cmmf") = mykey3(1)
                        CMMFKAMAssignmentDT.Rows.Add(dr)
                    Else
                        Cmmfkamassignmentid = result.Item("id")
                    End If
                    'DATASB.Append(String.Format("'{0:yyyy-MM-dd}'", myperiod) & vbTab &
                    '              Cmmfkamassignmentid & vbTab & myList(i).myrecord(mycol) & vbCrLf)
                    'DATASB.Append(String.Format("'{0:yyyy-MM-dd}'", myperiod) & vbTab &
                    '              Cmmfkamassignmentid & vbTab & mykey3(1) & vbTab & mykey(0) & vbTab & mykey2(1) & vbTab & myList(i).myrecord(mycol) & vbCrLf)
                    DATASB.Append(String.Format("'{0:yyyy-MM-dd}'", myperiod) & vbTab &
                                  Cmmfkamassignmentid & vbTab & mykey3(1) & vbTab & mykey(0) & vbTab & mykey2(1) & vbTab & mykey(1) & vbTab & myList(i).myrecord(mycol) & vbCrLf)
                End If
                ' End If


            Next
        Next

        If DATASB.Length > 0 Then
            myForm.ProgressReport(1, "Preparing Data...")
            Dim NextYear = String.Format("{0}-{1:MM}-01", myPeriod.Year + 1, myPeriod)
            'clean data for KAM
            'Dim sqlstr1 = String.Format("delete from sales.sfmlatxhk tx where id in(select tx.id from  sales.sfmlatxhk tx" &
            '                       " left join sales.sfcmmfkamassignment ck on ck.id = tx.cmmfkamassignmentid" &
            '                       " left join sales.sfkamassignment ks on ks.id = ck.kamassignmentid" &
            '                       " left join sales.sfmlacardname mc on mc.id = ks.mlacardnameid where kam  = '{0}')", kamDict(0))
            Dim sqlstr1 = String.Format("delete from sales.sfmlatxhk tx where kam  = '{0}' and tx.txdate >= '{1:yyyy-MM}-01' and tx.txdate < '{2}'", kamDict(0), myPeriod, NextYear)
            PostgresqlDBAdapter1.ExecuteNonQuery(sqlstr1)

            'copy
            myForm.ProgressReport(1, "Copy Data...")
            Dim sqlstr As String = "copy sales.sfmlatxhk(txdate,cmmfkamassignmentid,cmmf,mla,kam,customer,salesforecast) from stdin with null as 'Null';"
            ErrorMsg = PostgresqlDBAdapter1.copy(sqlstr, DATASB.ToString, myret)
            If myret Then
                myForm.ProgressReport(1, "Done.")
            End If
        Else
            ErrorMsg = "Nothing to import."
        End If

        Return myret
    End Function

    Private Sub populatedict(mydict As Dictionary(Of Integer, String), myrecord As String())
        For j = 0 To (myrecord.Length - 6) - HKReportProperty1.ColumnStartData
            mydict.Add(j, myrecord(HKReportProperty1.ColumnStartData - 1 + j))
        Next
    End Sub


End Class

Public Class HKData
    Public Property Period As String
    Public Property myrecord As String()
    Public Property Group As String
End Class
Public Class TWData
    Public Property Period As String
    Public Property myrecord As String()
    Public Property Group As String
    Public Property KAM As String
End Class
Public Class MSData
    Public Property Period As String
    Public Property myrecord As String()
    Public Property GroupID As Integer
    Public Property KAM As String
End Class

Public Class SGData
    Public Property Period As String
    Public Property myrecord As String()
    Public Property GroupID As Integer
    Public Property KAM As String
End Class
Public Class THData
    Public Property Period As String
    Public Property myrecord As String()
    Public Property KAMAssignmentId As Integer    
End Class
Interface HKImport
    Property ErrorMsg As String
    Function Run(ByVal myForm As Object, ByVal Doc As List(Of DocCSV)) As Boolean
End Interface
