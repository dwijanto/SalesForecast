Imports System.Windows.Forms

Public Class DialogCMMF

    Dim DRV As DataRowView
    Dim PriceDRV As DataRowView
    Dim BrandBS As BindingSource
    Dim ProductLineGPSBS As BindingSource
    Dim FamilyBS As BindingSource
    Dim FamilyHelperBS As BindingSource
    Dim CatBrandBS As BindingSource
    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Public Sub New(ByVal drv As DataRowView, ByVal Pricedrv As DataRowView, ByVal BrandBS As BindingSource, ByVal ProductLineGPSBS As BindingSource, ByVal FamilyBS As BindingSource, ByVal FamilyHelperBS As BindingSource, ByVal CatBrandBS As BindingSource)
        InitializeComponent()
        Me.DRV = drv
        Me.PriceDRV = Pricedrv
        Me.DRV.BeginEdit()
        Me.BrandBS = BrandBS
        Me.ProductLineGPSBS = ProductLineGPSBS
        Me.FamilyBS = FamilyBS
        Me.FamilyHelperBS = FamilyHelperBS
        Me.CatBrandBS = CatBrandBS

        initData()
    End Sub

    Public Overloads Function Validate() As Boolean
        'Check combobox
        Dim myret As Boolean = True
        Dim cbdrv As DataRowView = ComboBox1.SelectedItem
        If Not IsNothing(cbdrv) Then
            DRV.Item("brand") = cbdrv.Item("brand")
        End If

        ErrorProvider1.SetError(ComboBox4, "")
        Dim cbdrv4 As DataRowView = ComboBox4.SelectedItem
        If Not IsNothing(cbdrv4) Then
            DRV.Item("productlinegpsname") = cbdrv4.Item("productlinegpsname")
            DRV.Item("familylv2") = DRV.Item("familylv2")
        Else
            myret = False
            ErrorProvider1.SetError(ComboBox4, "Product Line cannot be blank.")
        End If
        

        Dim cbdrv5 As DataRowView = ComboBox5.SelectedItem
        If Not IsNothing(cbdrv5) Then
            'DRV.Item("familyname") = cbdrv5.Item("familyname")
            DRV.Item("familyname") = String.Format("{0:000} - {1}", cbdrv5.Row.Item("familyid"), cbdrv5.Row.Item("familyname"))
            DRV.Item("nsp1") = PriceDRV.Item("nsp1")
            DRV.Item("nsp2") = PriceDRV.Item("nsp2")
        End If
        ErrorProvider1.SetError(TextBox1, "")
        If IsDBNull(DRV.Item("cmmf")) Then
            ErrorProvider1.SetError(TextBox1, "CMMF cannot be blank.")
            Return False
        End If

        PriceDRV.Item("cmmf") = DRV.Item("cmmf")

        Dim cbdrv7 As DataRowView = ComboBox7.SelectedItem
        If Not IsNothing(cbdrv7) Then
            DRV.Item("catbrandname") = cbdrv7.Row.Item("catbrandname")
            If DRV.Item("catbrandid") = 0 Then
                DRV.Item("catbrandid") = DBNull.Value
                DRV.Item("catbrandname") = String.Empty
            End If
        End If



        Return myret
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.Validate Then
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
        ComboBox1.DataSource = BrandBS
        ComboBox1.DisplayMember = "brand"
        ComboBox1.ValueMember = "brand"

        ComboBox4.DataSource = ProductLineGPSBS
        ComboBox4.DisplayMember = "productlinegpsname"
        ComboBox4.ValueMember = "productlinegpsid"

        ComboBox5.DataSource = FamilyBS
        ComboBox5.DisplayMember = "familydesc"
        ComboBox5.ValueMember = "familyid"

        ComboBox7.DataSource = CatBrandBS
        ComboBox7.DisplayMember = "catbrandname"
        ComboBox7.ValueMember = "id"


        ComboBox1.DataBindings.Clear()

        ComboBox2.DataBindings.Clear()
        ComboBox3.DataBindings.Clear()

        ComboBox4.DataBindings.Clear()
        ComboBox5.DataBindings.Clear()
        ComboBox6.DataBindings.Clear()
        ComboBox7.DataBindings.Clear()


        

       

        TextBox1.DataBindings.Clear()

        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()

        TextBox2.DataBindings.Clear()
        TextBox6.DataBindings.Clear()
        TextBox7.DataBindings.Clear()

        DateTimePicker1.DataBindings.Clear()
        DateTimePicker2.DataBindings.Clear()

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "brand", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox2.DataBindings.Add(New Binding("Text", DRV, "status", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox3.DataBindings.Add(New Binding("Text", DRV, "productsegmentation", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox4.DataBindings.Add(New Binding("SelectedValue", DRV, "productlinegpsid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox5.DataBindings.Add(New Binding("SelectedValue", DRV, "familyid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox6.DataBindings.Add(New Binding("Text", DRV, "origin", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox7.DataBindings.Add(New Binding("SelectedValue", DRV, "catbrandid", True, DataSourceUpdateMode.OnPropertyChanged, ""))


        TextBox1.DataBindings.Add(New Binding("Text", DRV, "cmmf", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox3.DataBindings.Add(New Binding("Text", DRV, "reference", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox4.DataBindings.Add(New Binding("Text", DRV, "description", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox5.DataBindings.Add(New Binding("Text", DRV, "companystatus", True, DataSourceUpdateMode.OnPropertyChanged, ""))

        TextBox2.DataBindings.Add(New Binding("Text", PriceDRV, "nsp1", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox6.DataBindings.Add(New Binding("Text", PriceDRV, "nsp2", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker1.DataBindings.Add(New Binding("Text", DRV, "activedate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker2.DataBindings.Add(New Binding("Text", DRV, "launchingmonth", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox7.DataBindings.Add(New Binding("Text", DRV, "remarks", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        If DRV.Row.RowState = DataRowState.Detached Then            
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1
            ComboBox3.SelectedIndex = -1
            ComboBox4.SelectedIndex = -1
            ComboBox5.SelectedIndex = -1
            ComboBox6.SelectedIndex = -1
            ComboBox7.SelectedIndex = -1
        End If
    End Sub

    Private Sub validcombo4()
        Dim drv = ComboBox4.SelectedItem
        If Not IsNothing(drv) Then
            Me.DRV.Item("productlinegpsname") = drv.item("productlinegpsname")
            RaiseEvent RefreshDataGridView(Me, New EventArgs)
        End If
    End Sub
    Private Sub validcombo5()
        Dim drv = ComboBox5.SelectedItem
        If Not IsNothing(drv) Then
            'Me.DRV.Item("familyname") = drv.item("familyname")
            Me.DRV.Row.Item("familyname") = String.Format("{0:000} - {1}", drv.Row.Item("familyid"), drv.Row.Item("familyname"))
            Me.DRV.Item("familylv2") = drv.item("familylv2")

            RaiseEvent RefreshDataGridView(Me, New EventArgs)
        End If
    End Sub
    Private Sub Dialog1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub


    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectionChangeCommitted
        validcombo4()
    End Sub
    Private Sub ComboBox2_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectionChangeCommitted
        validcombo5()
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim myobj As Button = CType(sender, Button)
        Select Case myobj.Name
            Case "Button1"
                Dim myform = New FormHelper(FamilyHelperBS)
                'myform.DataGridView1.Columns(0).DataPropertyName = "familyname"
                myform.DataGridView1.Columns(0).DataPropertyName = "familydesc"
                If myform.ShowDialog = DialogResult.OK Then
                    Dim drv As DataRowView = FamilyHelperBS.Current

                    Me.DRV.BeginEdit()
                    Me.DRV.Row.Item("familyname") = String.Format("{0:000} - {1}", drv.Row.Item("familyid"), drv.Row.Item("familyname"))

                    'Need bellow code to sync with combobox
                    Dim myposition = FamilyBS.Find("familyid", drv.Row.Item("familyid"))
                    FamilyBS.Position = myposition
                End If

        End Select
        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    End Sub

End Class
