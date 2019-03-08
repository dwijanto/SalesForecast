Imports Npgsql

Public Class CMMFKAMAssignmentModel
    Implements IModel

    Dim myadapter As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance
    Public Property cmmf As Long
    Public Property kamassignmentid As Long

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "sales.sfcmmfkamassignment"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "id"
        End Get
    End Property

    Public ReadOnly Property FilterField
        Get
            Return Nothing '"[cmmf] like '*{0}*' or [reference] like '*{0}*' or [origin] like '*{0}*' or [description] like '*{0}*'  or [familyname] like '*{0}*' or [productlinegpsname] like '*{0}*'"
        End Get
    End Property

    Public Function GetLastId() As Long
        Dim sqlstr = String.Format("select id from {0}  order by id desc limit 1", tablename)
        Dim myret As Long = 0
        Using conn As Object = myadapter.getConnection
            conn.open()
            myadapter.ExecuteScalar(sqlstr, recordAffected:=myret)
        End Using
        Return myret
    End Function

    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u.* from {0} u order by {1}", TableName, SortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function

    Public Function Add(ByVal model As CMMFKAMAssignmentModel) As Long
        Dim myret As Long
        Dim sqlstr = String.Format("insert into sales.sfcmmfkamassignment(cmmf,kamassignmentid) values('{0}','{1}'); select currval('sales.sfcmmfkamassignment_id_seq');", model.cmmf, model.kamassignmentid)
        If Not myadapter.ExecuteScalar(sqlstr, recordAffected:=myret) Then
            Throw New Exception(String.Format("CMMFKamAssignmentModel.Add {0}", myadapter.ErrorMessage))
        End If
        Return myret
    End Function

    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save       
        Return Nothing
    End Function

End Class
