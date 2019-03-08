Imports System.Windows.Forms

Public Class DialogSalesDeduction

    Dim DRV As DataRowView
    Dim GroupBS As BindingSource
    Dim producttypes As New List(Of ProductType)
    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Public Sub New(ByVal drv As DataRowView, ByVal GroupBS As BindingSource)
        InitializeComponent()
        Me.DRV = drv
        Me.DRV.BeginEdit()
        Me.GroupBS = GroupBS

    End Sub

    Public Overloads Function Validate() As Boolean
        'Check combobox

        Return True
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.Validate Then
            Try
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                DRV.EndEdit()
                RaiseEvent RefreshDataGridView(Me, e)
                Me.Close()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DRV.CancelEdit()
            End Try
            
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

        ComboBox1.DataBindings.Clear()

        TextBox1.DataBindings.Clear()

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "groupid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox1.DataBindings.Add(New Binding("Text", DRV, "sd", True, DataSourceUpdateMode.OnPropertyChanged, "", "0%"))
  
        'If IsDBNull(DRV.Item(0)) Then
        If DRV.Row.RowState = DataRowState.Detached Then
            ComboBox1.SelectedIndex = -1
            ComboBox1.SelectedIndex = -1
        End If
    End Sub


    Private Sub Dialog1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initData()
    End Sub


  


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    End Sub
End Class
