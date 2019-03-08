Imports System.Threading
Public Class FormTousRawData
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myController As SFGroupTXHKController
    'Dim KAMController1 As KAMController
    Dim Criteria As String
    Dim FirstTime As Boolean = True
    Dim drv As DataRowView = Nothing
    Dim myform As New DialogSelectPeriod
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        RemoveHandler DialogTWParam.RefreshDataGridView, AddressOf RefreshDataGridView
        AddHandler DialogTWParam.RefreshDataGridView, AddressOf RefreshDataGridView
    End Sub


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loadData()
    End Sub

    Private Sub loadData()
        If Not myThread.IsAlive Then
            If FirstTime Then
                myform.ShowDialog()
                FirstTime = False
            End If

            If myform.DialogResult = Windows.Forms.DialogResult.OK Then
                Criteria = String.Format(" where txdate >= '{0:yyyy-MM-01}' and txdate <= '{1:yyyy-MM-01}'", myform.startperiod, myform.endperiod)
                ToolStripStatusLabel1.Text = ""
                myThread = New Thread(AddressOf DoWork)
                myThread.Start()
            Else
                Me.Close()
            End If

        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        myController = New SFGroupTXHKController
        'KAMController1 = New KAMController

        Try
            ProgressReport(1, "Loading..")
            If myController.loaddata(Criteria) Then
                ProgressReport(4, "Init Data")
            End If
            'Dim Criteria As String = "where location = 2"
            'KAMController1.loaddata(Criteria)

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
                    drv.Item("period") = Date.Today
                Case TxEnum.UpdateRecord
                    drv = myController.GetCurrentRecord
            End Select

            'Dim KAMBS = KAMController1.GetBindingSource

            'Me.drv.BeginEdit()
            'Dim myform = New DialogTWParam(drv, KAMBS)
            'myform.ShowDialog()
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

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub


End Class