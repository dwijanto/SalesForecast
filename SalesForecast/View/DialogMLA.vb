Imports System.Windows.Forms

Public Class DialogMLA

    Dim DRV As DataRowView


    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Public Sub New(ByVal drv As DataRowView)
        InitializeComponent()
        Me.DRV = drv
        Me.DRV.BeginEdit()


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

        TextBox1.DataBindings.Clear()

        TextBox1.DataBindings.Add(New Binding("Text", DRV, "mla", True, DataSourceUpdateMode.OnPropertyChanged, ""))

       
    End Sub


    Private Sub Dialog1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initData()
    End Sub





    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    End Sub

End Class
