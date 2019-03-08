Imports System.Windows.Forms

Public Class DialogKAM

    Dim DRV As DataRowView
    Private LocationList As New List(Of LocationModel)

    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Public Sub New(ByVal drv As DataRowView)
        InitializeComponent()
        Me.DRV = drv
        Me.DRV.BeginEdit()
        LocationList.Add(New LocationModel With {.locationid = 1, .locationname = "Hong Kong"})
        LocationList.Add(New LocationModel With {.locationid = 2, .locationname = "Taiwan"})
        LocationList.Add(New LocationModel With {.locationid = 3, .locationname = "Malaysia"})
        LocationList.Add(New LocationModel With {.locationid = 4, .locationname = "Singapore"})
        LocationList.Add(New LocationModel With {.locationid = 5, .locationname = "Thailand"})
    End Sub

    Public Overloads Function Validate() As Boolean
        'Check combobox
        Dim mydrv As LocationModel = ComboBox1.SelectedItem
        ErrorProvider1.SetError(ComboBox1, "")
        If IsNothing(mydrv) Then
            ErrorProvider1.SetError(ComboBox1, "Please select from the list.")
            Return False
        End If
        DRV.Item("locationname") = mydrv.locationname
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
        CheckBox1.DataBindings.Clear()
        ComboBox1.DataBindings.Clear()

        ComboBox1.DataSource = LocationList
        ComboBox1.DisplayMember = "locationname"
        ComboBox1.ValueMember = "locationid"

        TextBox1.DataBindings.Add(New Binding("Text", DRV, "userid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("Text", DRV, "username", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox3.DataBindings.Add(New Binding("Text", DRV, "fullname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "location", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        CheckBox1.DataBindings.Add(New Binding("Checked", DRV, "isactive", True, DataSourceUpdateMode.OnPropertyChanged, ""))

    End Sub


    Private Sub Dialog1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initData()
    End Sub


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    End Sub
End Class

Public Class LocationModel
    Public Property locationid
    Public Property locationname

    Public Overrides Function ToString() As String
        Return locationname
    End Function
End Class