Imports System.Windows.Forms

Public Class DialogFamily

    Dim DRV As DataRowView
    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)
    Private ProductLineBS As BindingSource
    Public Sub New(ByVal drv As DataRowView, ByVal ProductLineBS As BindingSource)
        InitializeComponent()
        Me.DRV = drv
        Me.DRV.BeginEdit()
        Me.productlineBS = ProductLineBS
    End Sub

    Public Overloads Function Validate() As Boolean
        'Check combobox
        Dim mydrv As DataRowView = ComboBox1.SelectedItem
        ErrorProvider1.SetError(ComboBox1, "")
        If IsNothing(mydrv) Then
            ErrorProvider1.SetError(ComboBox1, "Please select from the list.")
            Return False
        End If
        'DRV.Item("locationname") = mydrv.locationname
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

        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()

        ComboBox1.DataBindings.Clear()

        ComboBox1.DataSource = Me.ProductLineBS
        ComboBox1.DisplayMember = "productlinedesc"
        ComboBox1.ValueMember = "productlinedesc"

        TextBox1.DataBindings.Add(New Binding("Text", DRV, "familyid", True, DataSourceUpdateMode.OnPropertyChanged, "", "000"))
        TextBox2.DataBindings.Add(New Binding("Text", DRV, "familyname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox3.DataBindings.Add(New Binding("Text", DRV, "familylv2", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "productlinedesc", True, DataSourceUpdateMode.OnPropertyChanged, ""))


    End Sub


    Private Sub Dialog1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initData()
    End Sub


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    End Sub
End Class
