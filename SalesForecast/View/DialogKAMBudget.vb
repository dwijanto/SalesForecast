Imports System.Windows.Forms

Public Class DialogKAMBudget

    Dim DRV As DataRowView
    Dim ProductTypeModel1 As New ProductTypeModel
    Dim KamList As List(Of KAMModel)
    Dim GroupList As List(Of GroupModel)
    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Public Sub New(ByVal drv As DataRowView, ByVal KAMList As List(Of KAMModel), ByVal GroupList As List(Of GroupModel))
        InitializeComponent()
        Me.DRV = drv
        Me.DRV.BeginEdit()
        Me.KamList = KAMList
        Me.GroupList = GroupList
    End Sub

    Public Overloads Function Validate() As Boolean
        'Check combobox
        Dim mydrv3 As ProductTypeModel = ComboBox3.SelectedItem
        ErrorProvider1.SetError(ComboBox3, "")
        If IsNothing(mydrv3) Then
            ErrorProvider1.SetError(ComboBox3, "Please select from the list.")
            Return False
        End If
        DRV.Item("producttypename") = mydrv3.producttypename
        Dim mydrv2 As GroupModel = ComboBox2.SelectedItem
        ErrorProvider1.SetError(ComboBox2, "")
        If IsNothing(mydrv2) Then
            ErrorProvider1.SetError(ComboBox2, "Please select from the list.")
            Return False
        End If
        DRV.Item("groupname") = mydrv2.Groupname
        Dim mydrv1 As KAMModel = ComboBox1.SelectedItem
        ErrorProvider1.SetError(ComboBox1, "")
        If IsNothing(mydrv1) Then
            ErrorProvider1.SetError(ComboBox1, "Please select from the list.")
            Return False
        End If

        DRV.Item("period") = CDate(String.Format("{0:yyyy-MM}-1", DRV.Item("period")))
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
        ComboBox3.DataSource = ProductTypeModel1.GetProductList
        ComboBox3.DisplayMember = "producttypename"
        ComboBox3.ValueMember = "producttypeid"

        ComboBox2.DataSource = GroupList
        ComboBox2.DisplayMember = "groupname"
        ComboBox2.ValueMember = "groupid"

        ComboBox1.DataSource = KamList
        ComboBox1.DisplayMember = "username"
        ComboBox1.ValueMember = "username"

        TextBox1.DataBindings.Add(New Binding("Text", DRV, "budgetnett", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        'TextBox2.DataBindings.Add(New Binding("Text", DRV, "username", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        'TextBox3.DataBindings.Add(New Binding("Text", DRV, "fullname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "kam", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", DRV, "groupid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox3.DataBindings.Add(New Binding("SelectedValue", DRV, "producttype", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker1.DataBindings.Add(New Binding("value", DRV, "period", True, DataSourceUpdateMode.OnPropertyChanged, ""))
    End Sub


    Private Sub Dialog1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initData()
    End Sub


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    End Sub

End Class
Public Class ProductTypeModel
    Public Property producttypeid
    Public Property producttypename

    Public Overrides Function ToString() As String
        Return producttypename
    End Function

    Public Function GetProductList() As List(Of ProductTypeModel)
        Dim ProductTypeList As New List(Of ProductTypeModel)
        ProductTypeList.Add(New ProductTypeModel With {.producttypeid = 1, .producttypename = "SDA"})
        ProductTypeList.Add(New ProductTypeModel With {.producttypeid = 4, .producttypename = "CKW"})
        Return ProductTypeList
    End Function

End Class