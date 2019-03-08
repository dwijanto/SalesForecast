Imports Npgsql

Public Class GroupModel
    Implements IModel

    Dim myadapter As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance
    Public Property Groupname As String
    Public Property GroupId As Integer

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "sales.sfgroup"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "groupname,location"
        End Get
    End Property

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
    Public Function LoadData(ByVal DS As DataSet, ByVal Criteria As String) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u.* from {0} u {1} order by {2}", TableName, Criteria, SortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function

    Public Function GetGroupDict(ByVal location As Integer) As Dictionary(Of String, Integer)
        Dim mydict As Dictionary(Of String, Integer)

        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim DS As DataSet = New DataSet
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select groupname,id from sales.sfgroup g  " &
                                       " where location = {0} order by 1,2", location)

            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            mydict = New Dictionary(Of String, Integer)
            For Each dr As DataRow In DS.Tables(0).Rows
                mydict.Add(dr.Item("groupname"), dr.Item("id"))
            Next
        End Using
        Return mydict
    End Function
    Public Function PopulateGroupList(ByVal location As Integer) As List(Of GroupModel)
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim GroupList As New List(Of GroupModel)
        Dim DS As DataSet = New DataSet
        Using conn As Object = myadapter.getConnection
            conn.Open()

            Dim sqlstr = String.Format("select groupname,id from sales.sfgroup g  " &
                                       " where location = {0} order by 1,2", location)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)

            For Each dr As DataRow In DS.Tables(0).Rows
                GroupList.Add(New GroupModel With {.GroupId = dr.Item("id"),
                                               .Groupname = dr.Item("groupname")})
            Next
        End Using
        Return GroupList

    End Function
    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Return Nothing
    End Function
End Class
