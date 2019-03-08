Imports System.Windows.Forms

Public Class DialogGrossSalesTargetTW

    Dim DRV As DataRowView
    Dim GroupBS As BindingSource
    Dim KAMBS As BindingSource
    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Public Sub New(ByVal drv As DataRowView, ByVal GroupBS As BindingSource, ByVal KAMBS As BindingSource)
        InitializeComponent()
        Me.DRV = drv
        Me.DRV.BeginEdit()
        Me.GroupBS = GroupBS
        Me.KAMBS = KAMBS
    End Sub

    Public Overloads Function Validate() As Boolean
        'Check combobox
        validcombo1()
        validcombo2()
        Return True
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.Validate Then
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            DRV.EndEdit()
            RaiseEvent RefreshDataGridView(Me, e)
            Me.Close()
        Else
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        DRV.CancelEdit()
        RaiseEvent RefreshDataGridView(Me, e)
        Me.Close()
    End Sub
    Private Sub initData()
        ComboBox1.DataSource = GroupBS
        ComboBox1.DisplayMember = "groupname"
        ComboBox1.ValueMember = "id"

        ComboBox2.DataSource = KAMBS
        ComboBox2.DisplayMember = "username"
        ComboBox2.ValueMember = "username"

        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()

        TextBox1.DataBindings.Clear()

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "groupid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", DRV, "kam", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox1.DataBindings.Add(New Binding("Text", DRV, "amount", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))
        DateTimePicker1.DataBindings.Add(New Binding("Text", DRV, "period", True, DataSourceUpdateMode.OnPropertyChanged))

        If DRV.IsNew Then
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1
        End If
    End Sub


    Private Sub Dialog1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initData()
    End Sub


    Private Sub validcombo1()
        Dim drv As DataRowView = ComboBox1.SelectedItem
        If Not IsNothing(drv) Then
            Me.DRV.Item("groupname") = drv.Item("groupname")
            RaiseEvent RefreshDataGridView(Me, New EventArgs)
        End If
    End Sub
    Private Sub validcombo2()
        Dim drv As DataRowView = ComboBox2.SelectedItem
        If Not IsNothing(drv) Then
            Me.DRV.Item("KAM") = drv.Item("username")
            RaiseEvent RefreshDataGridView(Me, New EventArgs)
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        validcombo1()
    End Sub


    Private Sub ComboBox2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox2.SelectionChangeCommitted
        validcombo2()
    End Sub
End Class
