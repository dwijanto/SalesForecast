Imports System.Threading
Imports System.Text
Imports Npgsql

Public Class ImportMlaNspMonthly
    Private myForm As Object
    Private FileNameFullPath As String
    Private myThread As New Thread(AddressOf doWork)
    Private hasError As Boolean = False
    Dim sbError As StringBuilder
    Dim PostgresqlDBAdapter1 As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance
    Public Sub New(ByVal Parent As Object, ByVal FileNameFullPath As String)
        Me.myForm = Parent
        Me.FileNameFullPath = FileNameFullPath
    End Sub

    Public Sub Start()
        If Not myThread.IsAlive Then
            myThread = New Thread(AddressOf doWork)
            myThread.SetApartmentState(ApartmentState.STA)
            myThread.Start()
        Else
            myForm.ProgressReport(1, "Please wait until the current process is finished.")
        End If
    End Sub



    Sub doWork()
        myForm.ProgressReport(6, "Start")
        Dim sb As New StringBuilder
        sbError = New StringBuilder
        Dim myList As New List(Of String())
        Dim myrecord() As String
        Dim sw As New Stopwatch
        sw.Start()
        Try
            Using objTFParser = New FileIO.TextFieldParser(FileNameFullPath)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(",")
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0
                    myForm.ProgressReport(1, "Read Data")

                    Do Until .EndOfData
                        myrecord = .ReadFields
                        myList.Add(myrecord)
                    Loop

                    For i = 2 To myList.Count - 1
                        For j = 1 To myList(i).Length - 1
                            If myList(i)(j).Length > 0 Then
                                Dim mydata As MlaNspModel = New MlaNspModel With {.cmmf = myList(i)(0),
                                                                                .mla = myList(0)(j),
                                                                              .period = myList(1)(j),
                                                                             .nsp = myList(i)(j)}

                                sb.Append(mydata.cmmf & vbTab &
                                          mydata.mla & vbTab &
                                          mydata.period & vbTab &
                                          mydata.nsp & vbCrLf)
                            End If

                        Next
                    Next

                    If sb.Length > 0 Then
                        Dim sqlstr1 = String.Format("delete from sales.sfmlansp;select setval('sales.sfmlansp_id_seq',1,false)")
                        'PostgresqlDBAdapter1.ExecuteNonQuery(sqlstr1)
                        'copy
                        Dim sqlstr As String = String.Format("{0};copy sales.sfmlansp(cmmf,mla,period,nsp1) from stdin with null as 'Null';", sqlstr1)
                        Dim myret As Boolean
                        Dim ErrMessage = PostgresqlDBAdapter1.copy(sqlstr, sb.ToString, myret)
                        sw.Stop()
                        If myret Then
                            myForm.ProgressReport(1, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
                        Else
                            myForm.ProgressReport(1, ErrMessage)

                        End If
                    Else
                        myForm.ProgressReport(1, "Nothing to import.")
                    End If

                End With
            End Using
        Catch ex As Exception
            myForm.ProgressReport(1, ex.Message)
        Finally
            sw.Stop()
            myForm.ProgressReport(5, "Start")
            myForm.ProgressReport(5, "Stop")
        End Try


    End Sub
End Class

Public Class MlaNspModel
    Public Property cmmf
    Public Property mla
    Public Property period
    Public Property nsp
    Public Property ErrMessage As String = String.Empty
    Dim PostgresqlDBAdapter1 As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance
    'Dim PeriodMlaList As List(Of String)
    Dim _FieldNameAndType As String = String.Empty
    Dim _FieldName As String = String.Empty
    Dim _MlaList As List(Of String) = New List(Of String)

    Public ReadOnly Property FieldNameAndType As String
        Get
            If _FieldNameAndType.Length = 0 Then
                GetPeriodMla()
            End If
            Return _FieldNameAndType
        End Get
    End Property
    Public ReadOnly Property FieldName As String
        Get
            If _FieldName.Length = 0 Then
                GetPeriodMla()
            End If
            Return _FieldName
        End Get
    End Property

    Public ReadOnly Property MlaList As List(Of String)
        Get
            If _MlaList.Count = 0 Then
                GetMLAList()
            End If
            Return _MlaList
        End Get
    End Property


    Public Function GetPeriodMla() As Boolean
        Dim myret As Boolean = False
        Dim DS As New DataSet
        Try
            Dim dataadapter As NpgsqlDataAdapter = PostgresqlDBAdapter1.getDbDataAdapter
            Using conn As Object = PostgresqlDBAdapter1.getConnection
                conn.Open()
                Dim sqlstr = String.Format("select period::text || '_' || mla::text as periodmla from sales.sfmlansp" &
                                           " group by mla,period order by mla desc,period")
                dataadapter.SelectCommand = PostgresqlDBAdapter1.getCommandObject(sqlstr, conn)
                dataadapter.SelectCommand.CommandType = CommandType.Text
                dataadapter.Fill(DS)
                Dim FieldNameAndTypeSB As New StringBuilder
                Dim FieldNameSB As New StringBuilder
                For Each dr As DataRow In DS.Tables(0).Rows
                    If FieldNameAndTypeSB.Length > 0 Then
                        FieldNameAndTypeSB.Append(",")
                        FieldNameSB.Append(",")
                    End If
                    FieldNameAndTypeSB.Append(String.Format("""{0}"" Numeric", dr.Item(0)))
                    FieldNameSB.Append(String.Format("""{0}""", dr.Item(0)))
                Next
                _FieldNameAndType = FieldNameAndTypeSB.ToString
                _FieldName = FieldNameSB.ToString
                myret = True
            End Using
        Catch ex As Exception
            ErrMessage = ex.Message
        End Try
        Return myret
    End Function

    Public Function GetMLAList() As List(Of String)
        Dim DS As New DataSet
        Try
            Dim dataadapter As NpgsqlDataAdapter = PostgresqlDBAdapter1.getDbDataAdapter
            Using conn As Object = PostgresqlDBAdapter1.getConnection
                conn.Open()
                Dim sqlstr = String.Format("select distinct mla from sales.sfmlansp order by mla;")
                dataadapter.SelectCommand = PostgresqlDBAdapter1.getCommandObject(sqlstr, conn)
                dataadapter.SelectCommand.CommandType = CommandType.Text
                dataadapter.Fill(DS)
                Dim FieldNameAndTypeSB As New StringBuilder
                Dim FieldNameSB As New StringBuilder
                For Each dr As DataRow In DS.Tables(0).Rows
                    _MlaList.Add(dr.Item(0))
                Next
            End Using
        Catch ex As Exception
            ErrMessage = ex.Message
        End Try
        Return _MlaList
    End Function

End Class
