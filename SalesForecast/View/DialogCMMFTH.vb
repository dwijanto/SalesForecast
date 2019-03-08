Imports System.Windows.Forms

Public Class DialogCMMFTH

    Dim DRV As DataRowView


    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)
    Dim ProductLine As AutoCompleteStringCollection
    Dim Status As AutoCompleteStringCollection
    Dim FamilyLVL2 As AutoCompleteStringCollection

    Public Sub New(ByVal drv As DataRowView, ByRef ProductLine As AutoCompleteStringCollection, ByRef Status As AutoCompleteStringCollection, ByRef FamilyLvl2 As AutoCompleteStringCollection)
        InitializeComponent()
        Me.DRV = drv

        Me.DRV.BeginEdit()
        Me.ProductLine = ProductLine
        Me.Status = Status
        Me.FamilyLVL2 = FamilyLvl2
    End Sub

    Public Overloads Function Validate() As Boolean
        Dim myret As Boolean = True

        ErrorProvider1.SetError(TextBox1, "")
        If IsDBNull(DRV.Item("cmmf")) Then
            ErrorProvider1.SetError(TextBox1, "Value cannot be blank.")
            myret = False
        End If

        Return myret
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.Validate Then
            If Not ProductLine.Contains(TextBox4.Text) Then
                ProductLine.Add(TextBox4.Text)
            End If

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
        TextBox2.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox2.AutoCompleteCustomSource = ProductLine
        TextBox2.AutoCompleteMode = AutoCompleteMode.SuggestAppend

        TextBox3.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox3.AutoCompleteCustomSource = FamilyLVL2
        TextBox3.AutoCompleteMode = AutoCompleteMode.SuggestAppend


        TextBox6.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox6.AutoCompleteCustomSource = Status
        TextBox6.AutoCompleteMode = AutoCompleteMode.SuggestAppend



        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()
        TextBox6.DataBindings.Clear()
        TextBox7.DataBindings.Clear()
        TextBox8.DataBindings.Clear()


        TextBox1.DataBindings.Add(New Binding("Text", DRV, "cmmf", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("Text", DRV, "productline", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox3.DataBindings.Add(New Binding("Text", DRV, "familylv2", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox4.DataBindings.Add(New Binding("Text", DRV, "commercialcode", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox5.DataBindings.Add(New Binding("Text", DRV, "itemdescription", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox6.DataBindings.Add(New Binding("Text", DRV, "status", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox7.DataBindings.Add(New Binding("Text", DRV, "rsp", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))
        TextBox8.DataBindings.Add(New Binding("Text", DRV, "cogs", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))

    End Sub

    Private Sub Dialog1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initData()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

End Class
