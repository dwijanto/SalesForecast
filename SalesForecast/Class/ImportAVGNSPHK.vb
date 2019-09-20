Imports System.Threading
Imports System.Text

Public Class ImportAVGNSPHK
    Private myForm As FormCMMF
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
                        If count > 0 Then '10 Then
                            myList.Add(myrecord)
                        End If
                        count += 1
                    Loop

                    For Each d In myList
                        ValidData(d)
                        If d.Length = 3 Then
                            sb.Append(String.Format("{0}{3}{1}{3}{2}{4}", d(0), d(1), d(2), vbTab, vbCrLf))
                        Else
                            sbError.Append(d)
                        End If

                    Next

                    If hasError Then
                        myForm.ProgressReport(1, "Error found. Please check.")
                        myForm.ProgressReport(7, sbError.ToString)
                    Else
                        If sb.Length > 0 Then
                            Dim sqlstr1 = String.Format("delete from sales.sfcmmfnsp;")
                            'PostgresqlDBAdapter1.ExecuteNonQuery(sqlstr1)
                            'copy
                            Dim sqlstr As String = String.Format("{0};copy sales.sfcmmfnsp(cmmf,nsp1,nsp2) from stdin with null as 'Null';", sqlstr1)
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

    Private Sub ValidData(d As String())
        If Not IsNumeric(d(0)) Or Not IsNumeric(d(1)) Or Not IsNumeric(d(2)) Then
            sbError.Append(String.Format("{0},{1},{2}{3}", d(0), d(1), d(2), vbCrLf))
            hasError = True
        End If
    End Sub

End Class
