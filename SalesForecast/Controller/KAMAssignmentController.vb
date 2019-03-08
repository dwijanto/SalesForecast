Public Class KAMAssignmentController
    Implements IController
    Public Model As New KAMAssignmentModel
    Dim ListModel As New List(Of KAMAssignmentModel)
    Dim BS As BindingSource
    Dim DS As DataSet
    Public ReadOnly Property GetTable As DataTable Implements IController.GetTable
        Get
            Return DS.Tables(Model.tablename).Copy()
        End Get
    End Property

    Public Function GetLastID() As Long
        Return Model.GetLastId
    End Function

    Public ReadOnly Property GetBindingSource As BindingSource
        Get
            Dim BS As New BindingSource
            BS.DataSource = GetTable
            BS.Sort = Model.sortField
            Return BS
        End Get
    End Property

    Public Function loaddata() As Boolean Implements IController.loaddata
        Dim myret As Boolean = False
        Model = New KAMAssignmentModel
        DS = New DataSet
        If Model.LoadData(DS) Then
            Dim pk(1) As DataColumn
            pk(0) = DS.Tables(0).Columns("mlacardnameid")
            pk(1) = DS.Tables(0).Columns("kam")
            DS.Tables(0).PrimaryKey = pk
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)
            myret = True
        End If
        Return myret
    End Function
    Public Function getAssignments(ByVal Criteria As String) As List(Of KAMAssignmentModel)
        Dim myret As Boolean = False
        Model = New KAMAssignmentModel
        ListModel = Model.PopulateKAMAssignment(Criteria)
        Return ListModel
    End Function


    Public Function save() As Boolean Implements IController.save
        Return Nothing
    End Function
End Class
