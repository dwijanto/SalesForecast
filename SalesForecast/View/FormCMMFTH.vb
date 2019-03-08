Imports System.Threading
Imports System.Text

Public Class FormCMMFTH
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myThread2 As New System.Threading.Thread(AddressOf DoImport)
    Dim myController As CMMFTHController

    Dim drv As DataRowView = Nothing

    Dim ProductLine As AutoCompleteStringCollection
    Dim Status As AutoCompleteStringCollection
    Dim FamilyLVL2 As AutoCompleteStringCollection

    Dim PostgresqlDBAdapter1 As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        RemoveHandler DialogCMMFSG.RefreshDataGridView, AddressOf RefreshDataGridView
        AddHandler DialogCMMFSG.RefreshDataGridView, AddressOf RefreshDataGridView
    End Sub


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loadData()
    End Sub

    Private Sub loadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        myController = New CMMFTHController

        Try
            ProgressReport(1, "Loading..")
            If myController.loaddata() Then
                ProgressReport(4, "Init Data")
            End If
            ProductLine = myController.Model.getProductLine
            Status = myController.Model.getStatus
            FamilyLVL2 = myController.Model.getFamilyLVL2
            ProgressReport(1, "Done.")

        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try
    End Sub


    Public Sub showTx(ByVal tx As TxEnum)
        If Not myThread.IsAlive Then
            Select Case tx
                Case TxEnum.NewRecord
                    drv = myController.GetNewRecord

                Case TxEnum.UpdateRecord
                    drv = myController.GetCurrentRecord

            End Select


            Me.drv.BeginEdit()

            Dim myform = New DialogCMMFTH(drv, ProductLine, Status, FamilyLVL2)
            myform.ShowDialog()
        End If

    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = myController.BS
                Case 8
                    loadData()
            End Select
        End If
    End Sub


    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        Dim obj As ToolStripTextBox = DirectCast(sender, ToolStripTextBox)
        myController.ApplyFilter = obj.Text
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        showTx(TxEnum.UpdateRecord)
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        showTx(TxEnum.NewRecord)
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        If Not IsNothing(myController.GetCurrentRecord) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows

                    myController.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub
    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        loadData()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        myController.save()

    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        ToolStripButton5.PerformClick()
    End Sub

    Private Sub RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        Dim sqlstr As String = String.Empty
        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("CMMFTH{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            sqlstr = String.Format("select cmmf,productline,familylv2,commercialcode,itemdescription,status,rsp,cogs from sales.sfcmmfth order by cmmf")
            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx", "A1")
            myreport.Run(Me, e)

        End If
    End Sub
    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
    End Sub
    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        If Not myThread2.IsAlive Then
            'Get file
            If OpenFileDialog1.ShowDialog = DialogResult.OK Then
                myThread2 = New Thread(AddressOf DoImport)
                myThread2.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Private Sub DoImport()
        Try
            Dim mystr As New StringBuilder
            Dim myInsert As New System.Text.StringBuilder
            Dim myrecord() As String
            Using objTFParser = New FileIO.TextFieldParser(OpenFileDialog1.FileName)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0
                    ProgressReport(1, "Read Data")
                    Do Until .EndOfData
                        myrecord = .ReadFields
                        If count >= 1 Then
                            If myrecord(0).Length > 0 Then
                                myInsert.Append(PostgreSQLDbAdapter.validlong(myrecord(0)) & vbTab &
                                                PostgreSQLDbAdapter.validstr(myrecord(1)) & vbTab &
                                                PostgreSQLDbAdapter.validstr(myrecord(2)) & vbTab &
                                                PostgreSQLDbAdapter.validstr(myrecord(3)) & vbTab &
                                                PostgreSQLDbAdapter.validstr(myrecord(4)) & vbTab &
                                                PostgreSQLDbAdapter.validstr(myrecord(5)) & vbTab &
                                                PostgreSQLDbAdapter.validnumeric(myrecord(6).Replace(",", "")) & vbTab &
                                                PostgreSQLDbAdapter.validnumeric(myrecord(7).Replace(",", "")) & vbCrLf)
                            End If

                        End If
                        count += 1
                    Loop
                End With
            End Using
            'update record
            If myInsert.Length > 0 Then
                ProgressReport(1, "Start Add New Records")
                mystr.Append(String.Format("delete from sales.sfcmmfth"))
                Dim sqlstr As String = "copy sales.sfcmmfth(cmmf,productline,familylv2,commercialcode,itemdescription,status,rsp,cogs)  from stdin with null as 'Null';"
                Dim ra As Long = 0
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False

                Try

                    ProgressReport(1, "Delete Record Please wait!")
                    ra = PostgresqlDBAdapter1.ExecuteNonQuery(mystr.ToString)

                    ProgressReport(1, "Add Record Please wait!")
                    errmessage = PostgresqlDBAdapter1.copy(sqlstr, myInsert.ToString, myret)
                    If myret Then
                        ProgressReport(1, "Add Records Done.")
                    Else
                        ProgressReport(1, errmessage)
                    End If
                Catch ex As Exception
                    ProgressReport(1, ex.Message)

                End Try

            End If
        Catch ex As Exception
            ProgressReport(1, ex.Message)
            ProgressReport(3, "Set Continuous Again")
            Exit Sub
        End Try

        ProgressReport(3, "Set Continuous Again")
        ProgressReport(8, "Load Data")
    End Sub

End Class