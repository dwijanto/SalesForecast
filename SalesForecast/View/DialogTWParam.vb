Imports System.Windows.Forms

Public Class DialogTWParam
    Dim DRV As DataRowView
    Dim KAMBS As BindingSource

    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Public Sub New(ByVal drv As DataRowView, ByVal KAMBS As BindingSource)
        InitializeComponent()
        Me.DRV = drv
        Me.DRV.BeginEdit()
        Me.KAMBS = KAMBS
       
    End Sub

    Public Overloads Function Validate() As Boolean
        'Check combobox

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
        ComboBox1.DataSource = KAMBS
        ComboBox1.DisplayMember = "username"
        ComboBox1.ValueMember = "username"

        TextBox1.DataBindings.Clear()
        TextBox1.DataBindings.Clear()

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "kam", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox1.DataBindings.Add(New Binding("Text", DRV, "targetgross", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0"))

        DateTimePicker1.DataBindings.Add(New Binding("Text", DRV, "period", True, DataSourceUpdateMode.OnPropertyChanged))

        'If IsDBNull(DRV.Item(0)) Then
        If DRV.Row.RowState = DataRowState.Detached Then
            ComboBox1.SelectedIndex = -1
        End If
    End Sub


   
    Private Sub Dialog1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initData()
    End Sub



    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox1.TextChanged
        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    End Sub

End Class
