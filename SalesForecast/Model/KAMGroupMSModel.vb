Imports Npgsql
Public Class KAMGroupMSModel
    Implements IModel
    Dim myadapter As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance


    Public Property kam As String
    Public Property groupid As Long
    Public Property groupname As String

    Dim KAMGroupMSList As List(Of KAMGroupMSModel)

    Public Function LoadData(DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u.* from {0} u order by {1}", tablename, sortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, tablename)
            myret = True
        End Using
        Return myret
    End Function


    Public Function PopulateKAMGroupMS(criteria As String) As List(Of KAMGroupMSModel)
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim DS As DataSet = New DataSet
        Using conn As Object = myadapter.getConnection
            conn.Open()

            Dim sqlstr = String.Format("select distinct kam.username,kg.groupid,g.groupname from sales.sfkam kam  " &
                                       " left join sales.sfkamgroupms kg on kg.kam = kam.username " &
                                       " left join sales.sfgroup g on g.id = kg.groupid  " &
                                       " where kam.username = '{0}' order by 1,3", criteria)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, tablename)
            KAMGroupMSList = New List(Of KAMGroupMSModel)
            For Each dr As DataRow In DS.Tables(0).Rows
                KAMGroupMSList.Add(New KAMGroupMSModel With {.groupid = dr.Item("groupid"),
                                                                   .kam = dr.Item("username"),
                                                                   .groupname = "" & dr.Item("groupname")})
            Next
        End Using
        Return KAMGroupMSList
    End Function

    Public Function Add(ByVal model As KAMGroupMSModel) As Long
        Dim myret As Long
        Dim sqlstr = String.Format("insert into sales.sfkamgroupms(groupid,kam) values('{0}','{1}'); select currval('sales.sfkamgroupms_id_seq');", model.groupid, model.kam)
        If Not myadapter.ExecuteScalar(sqlstr, recordAffected:=myret) Then
            Throw New Exception(String.Format("KAMGroupMSModel.Add {0}", myadapter.ErrorMessage))
        End If
        Return myret
    End Function

    Public Function save(obj As Object, mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Return Nothing
    End Function

    Public ReadOnly Property sortField As String Implements IModel.sortField
        Get
            Return "kam"
        End Get
    End Property

    Public ReadOnly Property tablename As String Implements IModel.tablename
        Get
            Return "Sales.sfkamgroupms"
        End Get
    End Property
End Class
