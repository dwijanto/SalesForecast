Imports System.Text

Public Class ImportKAMTarget
    Enum ProductType
        SDA = 1
        TEFAL = 2
        LAGOSTINA = 3
    End Enum
    Private myForm As Object
    Private Filename As String
    Public ErrMessage As String
    Private SB As StringBuilder
    Private myController As New HKParamController
    Public Sub New(myForm As Object, Filename As String)
        Me.myForm = myForm
        Me.Filename = Filename
    End Sub
    Public Function Run() As Boolean
        myController.LoadDataImport()
        Dim myret As Boolean = False
        Dim myrecord() As String
        Dim myList As New List(Of String())
        Using objTFParser = New FileIO.TextFieldParser(Filename)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0
                myForm.ProgressReport(1, "Read Data..")
                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 1 Then
                        If myrecord(0) <> 0 Then
                            myList.Add(myrecord)
                        End If
                    End If
                    count += 1
                Loop
            End With
        End Using

        myForm.ProgressReport(1, "Build Records..")
        For i = 0 To myList.Count - 1
            If Not buildRecord(myList(i)) Then
                Return False
            End If
        Next

        myForm.ProgressReport(1, "Save Records..")
        myret = myController.save()        
        Return myret
    End Function

    Private Function buildRecord(myrecord As String()) As Boolean
        Dim myret As Boolean = False
        'Find Record, if not avail then create
        'Do Not create record when value Target/Net is 0

        'SDA
        If myrecord(0) = "201708" Then
            Debug.Print("in debug")
        End If

        Try
            Dim myKey(2) As Object
            myKey(0) = myrecord(0)
            myKey(1) = myrecord(1)
            myKey(2) = ProductType.SDA
            Dim myresult = myController.DS.Tables(0).Rows.Find(myKey)
            If Not IsNothing(myresult) Then
                myresult.Item("sdpct") = myrecord(7)
                myresult.Item("targetgross") = myrecord(4)
                myresult.Item("targetnet") = myrecord(10)
            Else
                If myrecord(4) <> 0 Then
                    Dim dr As DataRow = myController.DS.Tables(0).NewRow
                    dr.Item("period") = CDate(String.Format("{0}-{1}-1", myrecord(0).Substring(0, 4), myrecord(0).Substring(4, 2)))
                    dr.Item("myperiod") = myrecord(0)
                    dr.Item("kam") = myKey(1)
                    dr.Item("producttype") = myKey(2)
                    dr.Item("sdpct") = myrecord(7)
                    dr.Item("targetgross") = myrecord(4)
                    dr.Item("targetnet") = myrecord(10)
                    myController.DS.Tables(0).Rows.Add(dr)
                End If
            End If

            'Tefal
            Dim myKey1(2) As Object
            myKey1(0) = myrecord(0)
            myKey1(1) = myrecord(1)
            myKey1(2) = ProductType.TEFAL
            myresult = myController.DS.Tables(0).Rows.Find(myKey1)
            If Not IsNothing(myresult) Then
                myresult.Item("sdpct") = myrecord(5)
                myresult.Item("targetgross") = myrecord(2)
                myresult.Item("targetnet") = myrecord(8)
            Else
                If myrecord(2) <> 0 Then
                    Dim dr As DataRow = myController.DS.Tables(0).NewRow
                    dr.Item("period") = CDate(String.Format("{0}-{1}-1", myrecord(0).Substring(0, 4), myrecord(0).Substring(4, 2)))
                    dr.Item("myperiod") = myrecord(0)
                    dr.Item("kam") = myKey1(1)
                    dr.Item("producttype") = myKey1(2)
                    dr.Item("sdpct") = myrecord(5)
                    dr.Item("targetgross") = myrecord(2)
                    dr.Item("targetnet") = myrecord(8)
                    myController.DS.Tables(0).Rows.Add(dr)
                End If
            End If

            'Lagostina
            Dim myKey3(2) As Object
            myKey3(0) = myrecord(0)
            myKey3(1) = myrecord(1)
            myKey3(2) = ProductType.LAGOSTINA
            myresult = myController.DS.Tables(0).Rows.Find(myKey3)
            If Not IsNothing(myresult) Then
                myresult.Item("sdpct") = myrecord(6)
                myresult.Item("targetgross") = myrecord(3)
                myresult.Item("targetnet") = myrecord(9)
            Else
                If myrecord(3) <> 0 Then
                    Dim dr As DataRow = myController.DS.Tables(0).NewRow
                    dr.Item("period") = CDate(String.Format("{0}-{1}-1", myrecord(0).Substring(0, 4), myrecord(0).Substring(4, 2)))
                    dr.Item("myperiod") = myrecord(0)
                    dr.Item("kam") = myKey3(1)
                    dr.Item("producttype") = myKey3(2)
                    dr.Item("sdpct") = myrecord(6)
                    dr.Item("targetgross") = myrecord(3)
                    dr.Item("targetnet") = myrecord(9)
                    myController.DS.Tables(0).Rows.Add(dr)
                End If
            End If
            myret = True
        Catch ex As Exception
            ErrMessage = ex.Message
        End Try


        Return myret


    End Function

End Class
