Imports System.Threading

Public Class FormKAMTarget
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim ImportThread As New System.Threading.Thread(AddressOf DoImport)
    Dim myController As HKParamController
    Dim KAMController1 As KAMController
    'Dim FamilyController1 As FamilyController
    'Dim ProductLineGPSController1 As ProductLineGPSController

    'Dim VendorController As VendorTxAdapter
    'Dim PMController As PMAdapter

    Dim drv As DataRowView = Nothing
    Dim OpenFileDialog1 As New OpenFileDialog
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        RemoveHandler DialogHKParam.RefreshDataGridView, AddressOf RefreshDataGridView
        AddHandler DialogHKParam.RefreshDataGridView, AddressOf RefreshDataGridView
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
        myController = New HKParamController
        KAMController1 = New KAMController
        'FamilyController1 = New FamilyController
        'ProductLineGPSController1 = New ProductLineGPSController

        'VendorController = New VendorTxAdapter
        'PMController = New PMAdapter
        Try
            ProgressReport(1, "Loading..")

            If myController.loaddata() Then
                ProgressReport(4, "Init Data")
            End If
            Dim criteria As String = "where location = 1"
            KAMController1.loaddata(criteria)
            'FamilyController1.loaddata()
            'ProductLineGPSController1.loaddata()
            ProgressReport(1, "Done.")

        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try
    End Sub

    Sub DoImport()
        Dim ImportKAMTarget1 As New ImportKAMTarget(Me, OpenFileDialog1.FileName)
        If Not ImportKAMTarget1.Run() Then
            ProgressReport(1, ImportKAMTarget1.ErrMessage)
        Else
            loadData()
            ProgressReport(1, "Done")
        End If
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
            Dim KAMBS = KAMController1.GetBindingSource
            'Dim ProductLineGPSBS = ProductLineGPSController1.GetBindingSource
            'Dim FamilyBS = FamilyController1.GetBindingSource
            'Dim FamilyHelperBS = FamilyController1.GetBindingSource
            Me.drv.BeginEdit()
            Dim myform = New DialogHKParam(drv, KAMBS)
            myform.ShowDialog()
        End If

    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
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

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        If Not ImportThread.IsAlive Then
            'Get File name
            ToolStripStatusLabel1.Text = ""
            If OpenFileDialog1.ShowDialog = DialogResult.OK Then
                ImportThread = New Thread(AddressOf DoImport)
                ImportThread.Start()
            End If
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub



End Class