Public Class DialogHKParam
    Dim DRV As DataRowView
    Dim KAMBS As BindingSource
    Dim producttypes As New List(Of ProductType)
    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Public Sub New(ByVal drv As DataRowView, ByVal KAMBS As BindingSource)
        InitializeComponent()
        Me.DRV = drv
        Me.DRV.BeginEdit()
        Me.KAMBS = KAMBS
        producttypes.Add(New ProductType With {.id = 1, .producttypename = "SDA"})
        producttypes.Add(New ProductType With {.id = 2, .producttypename = "CKW Tefal"})
        producttypes.Add(New ProductType With {.id = 3, .producttypename = "CKW LAGO"})
    End Sub

    Public Overloads Function Validate() As Boolean
        'Check combobox

        Dim cbdrv2 As ProductType = ComboBox2.SelectedItem
        DRV.Item("producttypename") = cbdrv2.producttypename
        
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

        ComboBox2.DataSource = producttypes
        ComboBox2.DisplayMember = "producttypename"
        ComboBox2.ValueMember = "id"


        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()



        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()


        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "kam", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", DRV, "producttype", True, DataSourceUpdateMode.OnPropertyChanged, ""))

        TextBox1.DataBindings.Add(New Binding("Text", DRV, "sdpct", True, DataSourceUpdateMode.OnPropertyChanged, "", "0.0%"))
        TextBox2.DataBindings.Add(New Binding("Text", DRV, "targetgross", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0"))
        TextBox3.DataBindings.Add(New Binding("Text", DRV, "targetnet", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0"))

        DateTimePicker1.DataBindings.Add(New Binding("Text", DRV, "period", True, DataSourceUpdateMode.OnPropertyChanged))

        'If IsDBNull(DRV.Item(0)) Then
        If DRV.Row.RowState = DataRowState.Detached Then
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1
        End If
    End Sub


    Private Sub validcombo2()
        Dim drv As ProductType = ComboBox2.SelectedItem
        If Not IsNothing(drv) Then
            Me.DRV.Item("producttypename") = drv.producttypename
            RaiseEvent RefreshDataGridView(Me, New EventArgs)
        End If
    End Sub
    Private Sub Dialog1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initData()
    End Sub


    Private Sub ComboBox2_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectionChangeCommitted
        validcombo2()
    End Sub




    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    End Sub
End Class

Public Class ProductType
    Public Property id As Integer
    Public Property producttypename As String
    Public Overrides Function ToString() As String
        Return producttypename
    End Function
End Class