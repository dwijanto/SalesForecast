Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Runtime.InteropServices

Enum TaskTypeEnum
    HKMLA = 1
    HKFG = 2
    TWFG = 3
    MSFG = 4
    SGFG = 5
    THFG = 6
End Enum
Public Class ImportSalesForecastHK
  
    Private FileName As String
    Public Property ErrorMsg As String
    Dim Task As Object
    Dim TaskType As TaskTypeEnum
    Dim MyDoc As List(Of DocCSV)
    Dim myFolder As String
    Dim MyForm As Object
    Dim MyPeriod As Date
    Public Sub New(ByVal myForm As Object, ByVal filename As String)
        Me.MyForm = myForm
        Me.FileName = filename
    End Sub

    Public Function ValidateFile() As Boolean
        'Open Excel File
        Dim myret As Boolean = False
        If openExcelFile() Then           
            myret = True
        Else

        End If

        Return myret
    End Function

    Public Function DoImportFile(ByVal myPeriod As Date) As Boolean
        Me.MyPeriod = myPeriod.Date
        Dim myret As Boolean
        Select Case TaskType
            Case 1
                Task = New HKImportMLA(FileName, myPeriod)
            Case 2
                Task = New HKImportGroup(FileName, myPeriod)
            Case 3
                Task = New TWImportGroup(FileName, myPeriod)
            Case 4
                Task = New MSImportGroup(FileName, myPeriod)
            Case 5
                Task = New SGImportGroup(FileName, myPeriod)
            Case 6
                Task = New THImportGroup(FileName, myPeriod)
        End Select
        Try
            myret = Task.run(MyForm, MyDoc)
            If Not myret Then
                ErrorMsg = Task.ErrorMsg
            End If
        Catch ex As Exception
            ErrorMsg = ex.Message
        End Try

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
            oWb = oXl.Workbooks.Open(FileName)

            'Check FileType
            oWb.Worksheets(1).select()
            oSheet = oWb.Worksheets(1)

            Dim buff = oSheet.Name.Split("-")
            If buff(0) = "HK" Then
                If buff(1) = "MLA" Then
                    TaskType = TaskTypeEnum.HKMLA
                ElseIf buff(1) = "FG" Then
                    TaskType = TaskTypeEnum.HKFG
                Else
                    Throw New System.Exception("File is not valid.")
                End If
            ElseIf buff(0) = "TW" Then
                If buff(1) = "Summary" Then
                    TaskType = TaskTypeEnum.TWFG
                End If
            ElseIf buff(0) = "MY" Then
                If buff(1) = "Summary" Then
                    TaskType = TaskTypeEnum.MSFG
                End If
            ElseIf buff(0) = "SG" Then
                If buff(1) = "Summary" Then
                    TaskType = TaskTypeEnum.SGFG
                End If
            ElseIf buff(0) = "TH" Then
                If buff(1) = "Summary" Then
                    TaskType = TaskTypeEnum.THFG
                End If
            Else
                Throw New System.Exception("File is not valid.")
            End If
            myFolder = Path.GetDirectoryName(FileName) '& "\" & Path.GetFileNameWithoutExtension(FileName) & ".csv"
            MyDoc = New List(Of DocCSV)
            Select Case TaskType
                Case TaskTypeEnum.HKMLA
                    PrepareMyDoc("MLA", MyDoc, oWb)                    
                Case TaskTypeEnum.HKFG
                    PrepareFGDoc("HKFG", MyDoc, oWb)
                Case TaskTypeEnum.TWFG
                    PrepareTWMyDoc("TWFG", MyDoc, oWb)
                Case TaskTypeEnum.MSFG
                    PrepareMSMyDoc("MYFG", MyDoc, oWb)
                Case TaskTypeEnum.SGFG
                    PrepareSGMyDoc("SGFG", MyDoc, oWb)
                Case TaskTypeEnum.THFG
                    PrepareTHMyDoc("THFG", MyDoc, oWb)
            End Select
            myret = True
        Catch ex As Exception
            ErrorMsg = ex.Message
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

    Private Sub PrepareMyDoc(DocType As String, MyDoc As List(Of DocCSV), oWb As Excel.Workbook)
        Dim ReportProperty1 As HKReportProperty = HKReportProperty.getInstance

        For i = 1 To oWb.Worksheets.Count '- 1
            oWb.Worksheets(i).select()
            Dim oSheet = oWb.Worksheets(i)
            Dim buff = oSheet.Name.Split("-")
            Dim kam = oSheet.cells(ReportProperty1.RowStartDataI - 2, ReportProperty1.ColumnStartData).value
            If buff.length = 3 Then
                Dim Period = CDate(String.Format("{0}-1", buff(2).Replace(".", "-")))
                Dim FileName = String.Format("{0}{1:yyyyMMdd}", DocType, Period)
                MyDoc.Add(New DocCSV With {.KAM = kam, .Period = Period,
                                           .Name = FileName,
                                           .folder = myFolder})
                Dim fullname As String = String.Format("{0}\{1}.csv", myFolder, FileName)
                oWb.SaveAs(Filename:=fullname, FileFormat:=Excel.XlFileFormat.xlUnicodeText, CreateBackup:=False)
            End If
        Next
    End Sub

    Private Sub PrepareFGDoc(DocType As String, MyDoc As List(Of DocCSV), oWb As Excel.Workbook)
        Dim ReportProperty1 As HKReportProperty = HKReportProperty.getInstance

        For i = 1 To oWb.Worksheets.Count - 1
            oWb.Worksheets(i).select()
            Dim oSheet = oWb.Worksheets(i)
            Dim buff = oSheet.Name.Split("-")

            If buff.length = 3 Then
                Dim FG = buff(2)              
                Dim FileName = String.Format("{0}{1}", DocType, FG)
                MyDoc.Add(New DocCSV With {.KAM = FG,
                                           .Name = FileName,
                                           .folder = myFolder})
                Dim fullname As String = String.Format("{0}\{1}.csv", myFolder, FileName)
                oWb.SaveAs(Filename:=fullname, FileFormat:=Excel.XlFileFormat.xlUnicodeText, CreateBackup:=False)
            End If
        Next
    End Sub

    Private Sub PrepareTWMyDoc(DocType As String, MyDoc As List(Of DocCSV), oWb As Excel.Workbook)
        Dim ReportProperty1 As TWReportProperty = TWReportProperty.getInstance

        For i = 2 To oWb.Worksheets.Count - 1
            oWb.Worksheets(i).select()
            Dim oSheet = oWb.Worksheets(i)
            Dim buff = oSheet.Name.Split("-")

            If buff.length = 3 Then
                Dim FG = buff(2)
                Dim KAM = buff(1)
                Dim FileName = String.Format("{0}{1}", DocType, FG)
                MyDoc.Add(New DocCSV With {.KAM = KAM,
                                            .FG = FG,
                                           .Name = FileName,
                                           .folder = myFolder})
                Dim fullname As String = String.Format("{0}\{1}.csv", myFolder, FileName)
                oWb.SaveAs(Filename:=fullname, FileFormat:=Excel.XlFileFormat.xlUnicodeText, CreateBackup:=False)
            End If
        Next
    End Sub
    Private Sub PrepareMSMyDoc(DocType As String, MyDoc As List(Of DocCSV), oWb As Excel.Workbook)
        Dim ReportProperty1 As MSReportProperty = MSReportProperty.getInstance

        For i = 2 To oWb.Worksheets.Count - 1
            oWb.Worksheets(i).select()
            Dim oSheet = oWb.Worksheets(i)
            Dim buff = oSheet.Name.Split("-")

            If buff.length = 3 Then
                Dim FG = buff(2)
                Dim KAM = buff(1)
                Dim FileName = String.Format("{0}{1}", DocType, FG)
                MyDoc.Add(New DocCSV With {.KAM = KAM,
                                            .FG = FG,
                                           .Name = FileName,
                                           .folder = myFolder})
                Dim fullname As String = String.Format("{0}\{1}.csv", myFolder, FileName)
                oWb.SaveAs(Filename:=fullname, FileFormat:=Excel.XlFileFormat.xlUnicodeText, CreateBackup:=False)
            End If
        Next
    End Sub
    Private Sub PrepareSGMyDoc(DocType As String, MyDoc As List(Of DocCSV), oWb As Excel.Workbook)
        Dim ReportProperty1 As SGReportProperty = SGReportProperty.getInstance

        For i = 2 To oWb.Worksheets.Count - 1
            oWb.Worksheets(i).select()
            Dim oSheet = oWb.Worksheets(i)
            Dim buff = oSheet.Name.Split("-")

            If buff.length = 3 Then
                Dim FG = buff(2)
                Dim KAM = buff(1)
                Dim FileName = String.Format("{0}{1}", DocType, FG)
                MyDoc.Add(New DocCSV With {.KAM = KAM,
                                            .FG = FG,
                                           .Name = FileName,
                                           .folder = myFolder})
                Dim fullname As String = String.Format("{0}\{1}.csv", myFolder, FileName)
                oWb.SaveAs(Filename:=fullname, FileFormat:=Excel.XlFileFormat.xlUnicodeText, CreateBackup:=False)
            End If
        Next
    End Sub

    Private Sub PrepareTHMyDoc(DocType As String, MyDoc As List(Of DocCSV), oWb As Excel.Workbook)
        Dim ReportProperty1 As THReportProperty = THReportProperty.getInstance

        For i = 2 To oWb.Worksheets.Count - 1
            oWb.Worksheets(i).select()
            Dim oSheet = oWb.Worksheets(i)
            Dim buff = oSheet.Name.Split("-")

            If buff.length = 3 Then
                Dim FG = buff(2)
                Dim KAMAssignmentid = buff(1)
                Dim FileName = String.Format("{0}{1}{2}", DocType, KAMAssignmentid, FG)
                MyDoc.Add(New DocCSV With {.KAMAssignmentID = KAMAssignmentid,
                                            .FG = FG,
                                           .Name = FileName,
                                           .folder = myFolder})
                Dim fullname As String = String.Format("{0}\{1}.csv", myFolder, FileName)
                oWb.SaveAs(Filename:=fullname, FileFormat:=Excel.XlFileFormat.xlUnicodeText, CreateBackup:=False)
            End If
        Next
    End Sub

End Class

Public Class DocCSV
    Public Property Name As String
    Public Property folder As String
    Public Property Period As Date
    Public Property KAM As String
    Public Property FG As String
    Public Property KAMAssignmentID As Integer
End Class