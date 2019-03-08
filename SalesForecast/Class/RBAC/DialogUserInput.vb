Imports System.Windows.Forms

Public Class DialogUserInput

    Private DRV As DataRowView
    Public Shared Event FinishUpdate()

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        RaiseEvent FinishUpdate()
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        RaiseEvent FinishUpdate()
        Me.Close()
    End Sub

    Private Sub InitDataDRV()

        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()


        TextBox1.DataBindings.Add(New Binding("text", DRV, "username", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("text", DRV, "email", True, DataSourceUpdateMode.OnPropertyChanged, ""))

    End Sub

    Public Sub New(ByVal drv As DataRowView)

        ' This call is required by the designer.
        InitializeComponent()
        Me.DRV = drv
        ' Add any initialization after the InitializeComponent() call.
        InitDataDRV()
    End Sub

End Class
