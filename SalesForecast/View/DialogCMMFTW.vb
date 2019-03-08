Imports System.Windows.Forms

Public Class DialogCMMFTW

    Dim DRV As DataRowView
    Dim PriceDRV As DataRowView

    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)
    Dim ProductRange As AutoCompleteStringCollection
    Public Sub New(ByVal drv As DataRowView, ByVal Pricedrv As DataRowView, ByRef ProductRange As AutoCompleteStringCollection)
        InitializeComponent()
        Me.DRV = drv
        Me.PriceDRV = Pricedrv
        Me.DRV.BeginEdit()
        Me.ProductRange = ProductRange
    End Sub

    Public Overloads Function Validate() As Boolean
        Dim myret As Boolean = True
        DRV.Item("nsp1") = PriceDRV.Item("nsp1")
        DRV.Item("nsp2") = PriceDRV.Item("nsp2")
        PriceDRV.Item("cmmf") = DRV.Item("cmmf")
        ErrorProvider1.SetError(TextBox1, "")
        If IsDBNull(DRV.Item("cmmf")) Then
            ErrorProvider1.SetError(TextBox1, "Value cannot be blank.")
            myret = False
        End If
        If ComboBox1.SelectedIndex = -1 Then
            ErrorProvider1.SetError(ComboBox1, "Please select from the list.")
            myret = False
        End If
        If ComboBox2.SelectedIndex = -1 Then
            ErrorProvider1.SetError(ComboBox2, "Please select from the list.")
            myret = False
        End If

        Return myret
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.Validate Then
            If Not ProductRange.Contains(TextBox4.Text) Then
                ProductRange.Add(TextBox4.Text)
            End If

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            DRV.EndEdit()
            PriceDRV.EndEdit()
            RaiseEvent RefreshDataGridView(Me, e)
            Me.Close()
        Else
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        DRV.CancelEdit()
        PriceDRV.CancelEdit()
        RaiseEvent RefreshDataGridView(Me, e)
        Me.Close()
    End Sub
    Private Sub initData()
        TextBox4.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox4.AutoCompleteCustomSource = ProductRange
        TextBox4.AutoCompleteMode = AutoCompleteMode.SuggestAppend

        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()

        TextBox1.DataBindings.Clear()

        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()

        TextBox2.DataBindings.Clear()
        TextBox6.DataBindings.Clear()
        TextBox7.DataBindings.Clear()
        TextBox8.DataBindings.Clear()
        TextBox9.DataBindings.Clear()

        ComboBox1.DataBindings.Add(New Binding("Text", DRV, "producttype", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox2.DataBindings.Add(New Binding("Text", DRV, "subproducttype", True, DataSourceUpdateMode.OnPropertyChanged, ""))

        TextBox1.DataBindings.Add(New Binding("Text", DRV, "cmmf", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("Text", DRV, "localcmmf", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox3.DataBindings.Add(New Binding("Text", DRV, "referenceno", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox4.DataBindings.Add(New Binding("Text", DRV, "productrange", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox5.DataBindings.Add(New Binding("Text", DRV, "chinesedesc", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox6.DataBindings.Add(New Binding("Text", DRV, "description", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox7.DataBindings.Add(New Binding("Text", PriceDRV, "nsp1", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00########"))
        TextBox8.DataBindings.Add(New Binding("Text", PriceDRV, "nsp2", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00########"))
        TextBox9.DataBindings.Add(New Binding("Text", DRV, "moq", True, DataSourceUpdateMode.OnPropertyChanged, ""))

        'If IsDBNull(DRV.Item(0)) Then
        If DRV.Row.RowState = DataRowState.Detached Then
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1
        End If
    End Sub

    Private Sub Dialog1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initData()
    End Sub

End Class
