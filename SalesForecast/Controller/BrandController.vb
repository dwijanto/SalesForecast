Public Class BrandController
    Implements IController
    Dim Model As New BrandModel
    Dim BS As BindingSource
    Dim BS2 As BindingSource
    Dim DS As DataSet
    Public ReadOnly Property GetTable As DataTable Implements IController.GetTable
        Get
            Return DS.Tables(Model.tablename).Copy()
        End Get
    End Property

    Public ReadOnly Property GetTableCatBrand As DataTable
        Get
            Return DS.Tables(1).Copy()
        End Get
    End Property

    Public ReadOnly Property GetBindingSource As BindingSource
        Get
            Dim BS As New BindingSource
            BS.DataSource = GetTable
            BS.Sort = Model.SortField
            Return BS
        End Get
    End Property

    Public ReadOnly Property GetCatBrandBindingSource As BindingSource
        Get
            'Dim BS As New BindingSource
            'BS.DataSource = GetTableCatBrand
            Return BS2
        End Get
    End Property

    Public Function loaddata() As Boolean Implements IController.loaddata
        Dim myret As Boolean = False
        Model = New BrandModel
        DS = New DataSet
        If Model.LoadData(DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("brand")
            DS.Tables(0).PrimaryKey = pk
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)

            Dim pk1(0) As DataColumn
            DS.Tables(1).PrimaryKey = pk1
            BS2 = New BindingSource
            BS2.DataSource = DS.Tables(1)
            myret = True
        End If
        Return myret

    End Function
   
    Public Function save() As Boolean Implements IController.save
        Return Nothing
    End Function
End Class
