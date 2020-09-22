Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Public Class ImportHKRawData
    Public errorMsg As String
    Dim filename As String
    Private myForm As Object
    Dim SB As StringBuilder
    Dim DS As DataSet
    Dim myadapter As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance
    Dim fullname As String

    Public Sub New(ByVal obj As Object, ByVal filename As String)
        Me.filename = filename
        Me.myForm = obj
    End Sub

    Public Function validatefile() As Boolean
        'Open Excel File
        Dim myret As Boolean = False
        If openExcelFile() Then
            myret = True
        Else

        End If

        Return myret
    End Function


    Public Function doImportFile(ByVal startperiod As Date, ByVal endperiod As Date)
        'get MasterDATA
        'ReadTextFile
        'Build DATA
        'Copy DATA

        Dim myret As Boolean = False
        SB = New StringBuilder
        DS = New DataSet
        Dim TxSB = New StringBuilder
        SB.Append("select id,groupname from sales.sfgroup where location = 1;") '1 is HK
        If myadapter.GetDataset(SB.ToString, DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("groupname")
            DS.Tables(0).PrimaryKey = pk
        Else
            Err.Raise(514, Description:="Error Get Dataset.")
        End If
        Dim myrecord() As String
        Dim myList As New List(Of String())
        Using objTFParser = New FileIO.TextFieldParser(fullname)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0
                myForm.ProgressReport(1, "Read Data")

                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then

                        myList.Add(myrecord)
                    End If
                    count += 1
                Loop
            End With
        End Using

        'Build DATA
        For i = 0 To myList.Count - 1
            'find groupid
            'Dim mykey(0) As Object

            'Dim groupname = myList(i)(0)
            'mykey(0) = groupname
            'Dim groupid As Integer
            'Dim result As Object
            'result = DS.Tables(0).Rows.Find(mykey)
            'If Not IsNothing(result) Then
            '    groupid = result.item("id")
            'Else
            '    Err.Raise(515, Description:=String.Format("Group : ""{0}"", is not available in master Group.", myList(i)(0)))
            '    Return False
            'End If
            'kam,groupid,cmmf,txdate,salesforecast
            If IsNumeric(myList(i)(17)) Then
                'Check valid dAte
                Dim txdate = CDate(myList(i)(13))
                If (txdate >= startperiod Or txdate <= endperiod) And myList(i)(16) <> "" Then
                    Dim mydata As sfmlatxhkModel = New sfmlatxhkModel With {.CMMF = myList(i)(0),
                                                                            .MLA = myList(i)(15),
                                                                            .KAM = myList(i)(14),
                                                                            .customer = myList(i)(16),
                                                                            .txdate = CDate(myList(i)(13)),
                                                                            .salesforecast = myList(i)(17),
                                                                            .cmmfkamassignmentid = myList(i)(22)}
                    TxSB.Append(
                        mydata.CMMF & vbTab &
                        mydata.MLA & vbTab &
                        mydata.KAM & vbTab &
                        mydata.customer & vbTab &
                        String.Format("'{0:yyyy-MM-dd}'", CDate(mydata.txdate)) & vbTab &
                        mydata.salesforecast & vbTab &
                        mydata.cmmfkamassignmentid & vbCrLf
                        )
                End If

            End If
        Next

        If TxSB.Length > 0 Then
            'clean data for based on selected Period
            Dim sqlstr1 = String.Format("delete from sales.sfmlatxhk tx where tx.txdate >= '{0:yyyy-MM-}01' and tx.txdate <= '{1:yyyy-MM-}01'", startperiod, endperiod)
            ' myadapter.ExecuteNonQuery(sqlstr1)

            'copy
            Dim sqlstr As String = String.Format("{0};copy sales.sfmlatxhk(cmmf,mla,kam,customer,txdate,salesforecast,cmmfkamassignmentid) from stdin with null as 'Null';", sqlstr1)
            errorMsg = myadapter.copy(sqlstr, TxSB.ToString, myret)
            If myret Then
                myForm.ProgressReport(1, "Done.")
            End If
        End If
        Return myret
    End Function


    Private Function openExcelFile() As Boolean
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Dim myret As Boolean = False
        Try
            'Create Object Excel 
            'ProgressReport(1, "Preparing Data...")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            oXl.Visible = False
            oXl.DisplayAlerts = False
            oWb = oXl.Workbooks.Open(filename)

            'Check FileType
            For i = 1 To oWb.Worksheets.Count
                oWb.Worksheets(i).select()
                oSheet = oWb.Worksheets(i)
                errorMsg = "Please check your worksheet name. No ""DATA"". "
                If oSheet.Name = "DATA" Then
                    'save to textfile
                    errorMsg = ""
                    Dim myFolder = Path.GetDirectoryName(filename)
                    oSheet.Columns("O:O").NumberFormat = "mm/dd/yyyy"
                    fullname = String.Format("{0}\{1}.csv", myFolder, Path.GetFileNameWithoutExtension(filename))
                    oWb.SaveAs(Filename:=fullname, FileFormat:=Excel.XlFileFormat.xlUnicodeText, CreateBackup:=False)
                    myret = True
                    Exit For
                End If
            Next

        Catch ex As Exception
            errorMsg = ex.Message
        Finally
            oXl.Quit()
            ExportToExcelFile.releaseComObject(oSheet)
            ExportToExcelFile.releaseComObject(oWb)
            ExportToExcelFile.releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'to make sure excel is no longer in memory
                ExportToExcelFile.EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try

        End Try
        Return myret
    End Function
End Class

Public Class sfmlatxhkModel
    Public Property CMMF As String
    Public Property MLA As String
    Public Property KAM As String
    Public Property customer As String
    Public Property txdate As String
    Public Property salesforecast As String
    Public Property cmmfkamassignmentid As String

End Class
