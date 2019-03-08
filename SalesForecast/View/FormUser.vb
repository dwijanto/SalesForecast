Imports System.Threading

Public Class FormUser
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myAdapter As UserController

    Private Sub LoadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub FormUser_Load(sender As Object, e As EventArgs) Handles Me.Load
        LoadData()
    End Sub

    Sub DoWork()
        myAdapter = New UserController
        Try
            ProgressReport(1, "Loading..")
            If myAdapter.loaddata() Then
                ProgressReport(4, "Init Data")
            End If
            ProgressReport(1, "Done.")
        Catch ex As Exception

            ProgressReport(1, ex.Message)
        End Try
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
                    DataGridView1.DataSource = myAdapter.BS
            End Select
        End If
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        LoadData()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Try
            If User.can("createUser") Then
                ShowTx(TxRecord.AddRecord)
            Else

                MessageBox.Show("You cannot create a new record.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub ShowTx(ByVal StatusTx As TxRecord)
        Dim drv As DataRowView = Nothing
        Select Case StatusTx
            Case TxRecord.AddRecord
                drv = myAdapter.BS.AddNew
            Case TxRecord.UpdateRecord
                drv = myAdapter.BS.Current
        End Select
        Dim myform As New DialogUserInput(drv)
        myform.ShowDialog()
    End Sub

    Private Sub DataGridView1_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView1.DoubleClick
        'ShowTx(TxRecord.UpdateRecord)
        ToolStripButton5.PerformClick()
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        myAdapter.save()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If Not IsNothing(myAdapter.BS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    myAdapter.BS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        RemoveHandler DialogUserInput.FinishUpdate, AddressOf RefreshDataGrid
        AddHandler DialogUserInput.FinishUpdate, AddressOf RefreshDataGrid
    End Sub

    Private Sub RefreshDataGrid()
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs)
        'Dim RBAC As New DbManager
        'Dim CreateUserRule As New CreateUserRule
        'CreateUserRule.data = RBAC.Serialize(CreateUserRule)
        'Dim myitem = New Permission With {.name = "createUser",
        '                                 .type = TypeEnum.TYPE_PERMISSION,
        '                                  .description = "Create User",
        '                                  .ruleName = "CreateUserRule",
        '                                  .data = Nothing}
        'RBAC.addRule(CreateUserRule)
        'RBAC.add(myitem)
        'RBAC.Dispose()
        'RBAC = Nothing
    End Sub

    Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs)
        Dim RBAC As New DbManager
        'Dim AdminRole As New Item With {.name = "Admin",
        '                                .description = "Admin Role",
        '                                .type = TypeEnum.TYPE_ROLE,
        '                                .data = Nothing}

        'RBAC.addItem(AdminRole)
        'Dim HKRole As New Item With {.name = "HongKong",
        '                                .description = "Hong Kong User",
        '                                .type = TypeEnum.TYPE_ROLE,
        '                                .data = Nothing}
        'RBAC.addItem(HKRole)

        'Dim TWRole As New Item With {.name = "Taiwan",
        '                                .description = "Taiwan User",
        '                                .type = TypeEnum.TYPE_ROLE,
        '                                .data = Nothing}
        'RBAC.addItem(TWRole)
        'Dim HKSERule As New Item With {.name = "View-HKSalesExtract",
        '                              .description = "Hong Kong Sales Extract Rule",
        '                              .type = TypeEnum.TYPE_PERMISSION,
        '                               .data = Nothing}

        'RBAC.addItem(HKSERule)
        'Dim HKRule As New Item With {.name = "View-HKReport",
        '                              .description = "Hong Kong Report Rule ",
        '                              .type = TypeEnum.TYPE_PERMISSION,
        '                               .data = Nothing}
        'RBAC.addItem(HKRule)
        'Dim TWRule As New Item With {.name = "View-TWReport",
        '                      .description = "Taiwan Report Rule ",
        '                      .type = TypeEnum.TYPE_PERMISSION,
        '                       .data = Nothing}
        'RBAC.addItem(TWRule)

        'Dim TWRoleManager As New Item With {.name = "TaiwanManager",
        '                                .description = "TaiwanManager",
        '                                .type = TypeEnum.TYPE_ROLE,
        '                                .data = Nothing}
        'RBAC.addItem(TWRoleManager)
        'Dim HKRoleManager As New Item With {.name = "HKManager",
        '                        .description = "HKManager",
        '                        .type = TypeEnum.TYPE_ROLE,
        '                        .data = Nothing}
        'RBAC.addItem(HKRoleManager)
    End Sub

    Private Sub ToolStripButton9_Click(sender As Object, e As EventArgs)
        'Dim RBAC As New DbManager
        'Dim AdminRole As New Item With {.name = "Admin",
        '                        .description = "Admin Role",
        '                        .type = TypeEnum.TYPE_ROLE,
        '                        .data = Nothing}
        'Dim TWRole As New Item With {.name = "Taiwan",
        '                               .description = "Taiwan User",
        '                               .type = TypeEnum.TYPE_ROLE,
        '                               .data = Nothing}
        'Dim HKRole As New Item With {.name = "HongKong",
        '                               .description = "Hong Kong User",
        '                               .type = TypeEnum.TYPE_ROLE,
        '                               .data = Nothing}
        'RBAC.addChild(AdminRole, TWRole)
        'RBAC.addChild(AdminRole, HKRole)
        'RBAC.addChild(RBAC.getRole("HongKong"), RBAC.getItem("View-HKSalesExtract"))
        'RBAC.addChild(RBAC.getRole("Admin"), RBAC.getItem("createUser"))
        'RBAC.addChild(RBAC.getRole("HongKong"), RBAC.getItem("View-HKReport"))
        'RBAC.addChild(RBAC.getRole("Taiwan"), RBAC.getItem("View-TWReport"))
    End Sub

    Private Sub ToolStripButton10_Click(sender As Object, e As EventArgs) Handles ToolStripButton10.Click
        'Dim RBAC As New DbManager
        ''RBAC.assign(RBAC.getRole("Admin"), 1)
        ''RBAC.assign(RBAC.getRole("Taiwan"), 2)
        'RBAC.assign(RBAC.getRole("HongKong"), 3)
        Dim drv As DataRowView = myAdapter.BS.Current
        Dim myform As New FormUserAssignment(drv.Row.Item("id"))
        myform.ShowDialog()

    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        Dim myform = New DialogItem("Role", "Description")
        If myform.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim RBAC As New DbManager
            Dim myRole As New Item With {.name = myform.TextBox1.Text,
                                            .description = myform.TextBox2.Text,
                                            .type = TypeEnum.TYPE_ROLE,
                                            .data = Nothing}
            RBAC.addItem(myRole)
        End If
    End Sub

    Private Sub ToolStripButton7_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        Dim myform = New DialogItem("Permission", "Description")
        If myform.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim RBAC As New DbManager
            Dim myRule As New Item With {.name = myform.TextBox1.Text,
                                            .description = myform.TextBox2.Text,
                                            .type = TypeEnum.TYPE_PERMISSION,
                                            .data = Nothing}
            RBAC.addItem(myRule)
        End If
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        ShowTx(TxRecord.UpdateRecord)
    End Sub
End Class