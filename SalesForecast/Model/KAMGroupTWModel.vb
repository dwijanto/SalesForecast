Imports Npgsql

Public Class KAMGroupTWModel
    Implements IModel
    Dim myadapter As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance


    Public Property kam As String
    Public Property groupid As Long
    Public Property groupname As String
    Public Property sd As Decimal


    Dim KAMGroupTWList As List(Of KAMGroupTWModel)

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


    Public Function PopulateKAMAGroupTW(criteria As String) As List(Of KAMGroupTWModel)
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim DS As DataSet = New DataSet
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select kam.username,kg.groupid,g.groupname,tsd.sd from sales.sfkam kam  " &
                                       " left join sales.sfkamgrouptw kg on kg.kam = kam.username " &
                                       " left join sales.sfgroup g on g.id = kg.groupid  " &
                                       " left join sales.sftwgroupsd tsd on tsd.groupid = kg.groupid  " &
                                       " where kam.username = '{0}' order by 1,3", criteria)

            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, tablename)
            KAMGroupTWList = New List(Of KAMGroupTWModel)
            For Each dr As DataRow In DS.Tables(0).Rows
                KAMGroupTWList.Add(New KAMGroupTWModel With {.groupid = dr.Item("groupid"),
                                                                   .kam = dr.Item("username"),
                                                                   .groupname = "" & dr.Item("groupname"),
                                                                   .sd = dr.Item("sd")})
            Next
        End Using
        Return KAMGroupTWList
    End Function

    Public Function Add(ByVal model As KAMGroupTWModel) As Long
        Dim myret As Long
        Dim sqlstr = String.Format("insert into sales.sfkamgrouptw(groupid,kam) values('{0}','{1}'); select currval('sales.sfkamgrouptw_id_seq');", model.groupid, model.kam)
        If Not myadapter.ExecuteScalar(sqlstr, recordAffected:=myret) Then
            Throw New Exception(String.Format("KAMGroupTWModel.Add {0}", myadapter.ErrorMessage))
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
            Return "Sales.sfkamgrouptw"
        End Get
    End Property
End Class
