Imports System.Threading

Public Class FormCMMF
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myController As CMMFController
    Dim BrandController1 As BrandController
    Dim FamilyController1 As FamilyController
    Dim ProductLineGPSController1 As ProductLineGPSController
    Dim CMMFNSPController1 As CMMFNSPController    
    Dim PeriodMLAList As List(Of String)
    'Dim VendorController As VendorTxAdapter
    'Dim PMController As PMAdapter
    Dim MLANSPFieldNameAndType As String
    Dim MLANSPFieldName As String
    Dim drv As DataRowView = Nothing
    Dim NSPdrv As DataRowView = Nothing

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        RemoveHandler DialogCMMF.RefreshDataGridView, AddressOf RefreshDataGridView
        AddHandler DialogCMMF.RefreshDataGridView, AddressOf RefreshDataGridView
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
        myController = New CMMFController
        BrandController1 = New BrandController
        FamilyController1 = New FamilyController
        ProductLineGPSController1 = New ProductLineGPSController
        CMMFNSPController1 = New CMMFNSPController
        Dim MlaNspModel1 = New MlaNspModel
        Try

            MLANSPFieldNameAndType = MlaNspModel1.FieldNameAndType
            MLANSPFieldName = MlaNspModel1.FieldName
            If MLANSPFieldName.Length > 0 Then
                MLANSPFieldName = "," & MLANSPFieldName
            End If
            If MLANSPFieldNameAndType.Length = 0 Then
                MLANSPFieldNameAndType = """nsp"" text"
            End If
            ProgressReport(1, "Loading..")
            Dim sqlstr = String.Format("with m as (select * from  crosstab('select cmmf,period::text || ''_'' ||mla as desc ,nsp1 from sales.sfmlansp order by cmmf','select  period::text || ''_'' ||mla as desc from sales.sfmlansp group by mla,period order by mla desc,period') as (cmmf text,{0}))" &
                                       " select u.cmmf::text,u.origin,u.sbu,u.brand,u.reference,u.productdesc,u.description,u.companystatus,u.status,u.productsegmentation,u.productlinegpsid,u.familyid,u.activedate,u.launchingmonth,u.remarks,to_char(f.familyid,'FM000') || ' - ' || f.familyname as familyname,f.familylv2,sbu.productlinegpsname,p.nsp1,p.nsp2,u.catbrandid,cbv.catbrandname {1}" &
                                       " from sales.sfcmmf u left join sales.sffamily f on f.familyid = u.familyid left join sales.sfproductlinegps sbu on sbu.productlinegpsid = u.productlinegpsid " &
                                       " left join sales.sfcmmfnsp p on p.cmmf = u.cmmf  left join sales.catbrandview cbv on cbv.id = u.catbrandid left join m on m.cmmf::bigint = u.cmmf" &
                                       " order by u.cmmf", MLANSPFieldNameAndType, MLANSPFieldName)
            'If myController.loaddata() Then
            If myController.loaddata(sqlstr) Then
                ProgressReport(4, "Init Data")
            End If
            BrandController1.loaddata()
            FamilyController1.loaddata()
            ProductLineGPSController1.loaddata()
            CMMFNSPController1.loaddata()
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
                    NSPdrv = CMMFNSPController1.GetNewRecord
                    Me.NSPdrv.BeginEdit()
                    Me.drv.BeginEdit()
                    drv.Item("activedate") = Date.Today
                Case TxEnum.UpdateRecord
                    drv = myController.GetCurrentRecord
                    NSPdrv = CMMFNSPController1.getRecord(drv.Item("cmmf"))
                    Me.NSPdrv.BeginEdit()
                    Me.drv.BeginEdit()
                    NSPdrv.Item("cmmf") = drv.Item("cmmf")
                    NSPdrv.Item("nsp1") = drv.Item("nsp1")
                    NSPdrv.Item("nsp2") = drv.Item("nsp2")
                    'NSPdrv.EndEdit()
            End Select
            Dim BrandBS = BrandController1.GetBindingSource
            Dim ProductLineGPSBS = ProductLineGPSController1.GetBindingSource
            Dim FamilyBS = FamilyController1.GetBindingSource
            Dim FamilyHelperBS = FamilyController1.GetBindingSource
            Dim CatBrandBS = BrandController1.GetCatBrandBindingSource



            Dim myform = New DialogCMMF(drv, NSPdrv, BrandBS, ProductLineGPSBS, FamilyBS, FamilyHelperBS, CatBrandBS)
            myform.ShowDialog()
        End If

    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Try
                Me.Invoke(d, New Object() {id, message})
            Catch ex As Exception

            End Try
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4
                    DataGridView1.AutoGenerateColumns = False

                    If DataGridView1.Columns.Count <= 16 Then
                        'Add New Column in DatagridView1
                        If MLANSPFieldName.Length > 0 Then
                            For Each mydata In MLANSPFieldName.Split(",")
                                Dim mycolumn As New DataGridViewTextBoxColumn
                                mycolumn.DataPropertyName = mydata.Replace("""", "")
                                mycolumn.HeaderText = mydata.Replace("""", "")
                                DataGridView1.Columns.Add(mycolumn)
                            Next
                        End If
                       

                    End If
                   
                    DataGridView1.DataSource = myController.BS

                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    MessageBox.Show(message)
            End Select

        End If


    End Sub


    'Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
    '    If Me.InvokeRequired Then
    '        Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
    '        Me.Invoke(d, New Object() {id, message})
    '    Else
    '        Select Case id
    '            Case 1
    '                ToolStripStatusLabel1.Text = message
    '            Case 4
    '                DataGridView1.AutoGenerateColumns = False
    '                DataGridView1.DataSource = myController.BS
    '        End Select
    '    End If
    'End Sub


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
                    CMMFNSPController1.RemoveRecord(CLng(drv.Cells.Item(0).FormattedValue))
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
        CMMFNSPController1.save()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        ToolStripButton5.PerformClick()
    End Sub

    Private Sub RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)
        DataGridView1.Invalidate()
    End Sub


    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        Dim openFileDialog1 As New OpenFileDialog
        If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim myImport As New ImportAVGNSPHK(Me, openFileDialog1.FileName)
            myImport.Start()
        End If
    End Sub

    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        Dim openFileDialog1 As New OpenFileDialog
        If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim myImport As New ImportMlaNspMonthly(Me, openFileDialog1.FileName)
            myImport.Start()
        End If
    End Sub
End Class

Public Enum TxEnum
    NewRecord = 1
    UpdateRecord = 2
End Enum