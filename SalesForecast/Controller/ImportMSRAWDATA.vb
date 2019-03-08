Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text

Public Class ImportMSRAWDATA
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
        SB.Append("select id,groupname from sales.sfgroup where location = 3;") '3 is malaysia
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
            Dim mykey(0) As Object

            Dim groupname = myList(i)(0)
            mykey(0) = groupname
            Dim groupid As Integer
            Dim result As Object
            result = DS.Tables(0).Rows.Find(mykey)
            If Not IsNothing(result) Then
                groupid = result.item("id")
            Else
                Err.Raise(515, Description:=String.Format("Group : ""{0}"", is not available in master Group.", myList(i)(0)))
                Return False
            End If
            'kam,groupid,cmmf,txdate,salesforecast
            If IsNumeric(myList(i)(15)) Then
                'Check valid dAte
                Dim txdate = CDate(myList(i)(14))
                If txdate >= startperiod Or txdate <= endperiod Then
                    TxSB.Append(myList(i)(2) & vbTab &
                        groupid & vbTab &
                        myList(i)(5) & vbTab &
                        String.Format("'{0:yyyy-MM-dd}'", txdate) & vbTab &
                        myList(i)(15) & vbCrLf
                        )
                End If
                
            End If
        Next

        If TxSB.Length > 0 Then
            'clean data for based on selected Period
            Dim sqlstr1 = String.Format("delete from sales.sfgrouptxms tx where tx.txdate >= '{0:yyyy-MM-}01' and tx.txdate <= '{1:yyyy-MM-}01'", startperiod, endperiod)
            myadapter.ExecuteNonQuery(sqlstr1)

            'copy
            Dim sqlstr As String = "copy sales.sfgrouptxms(kam,groupid,cmmf,txdate,salesforecast) from stdin with null as 'Null';"
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
                errorMsg = "Please check your worksheet name. No ""RAWDATA"". "
                If oSheet.Name = "RAWDATA" Then
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
