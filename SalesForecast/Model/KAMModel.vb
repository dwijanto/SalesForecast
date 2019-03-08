Imports Npgsql

Public Class KAMModel
    Implements IModel

    Dim myadapter As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance
    Public Property userid As String
    Public Property username As String
    Public Property location As Integer
    Public Property fullname As String

    Public ReadOnly Property FilterField
        Get
            Return "[userid] like '*{0}*' or [username] like '*{0}*' or [fullname] like '*{0}*' or [locationname] like '*{0}*'  "
        End Get
    End Property
    'Public Function save(obj As Object, mye As ContentBaseEventArgs) As Boolean Implements IModel.save
    '    Return Nothing
    'End Function

    Public Function LoadData(DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u.*,sales.get_location(location) as locationname from {0} u order by {1}", tablename, sortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, tablename)
            myret = True
        End Using
        Return myret
    End Function

    Public Function LoadData(DS As DataSet, criteria As String) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u.*,sales.get_location(location) as locationname from {0} u {1} order by {2}", tablename, criteria, sortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, tablename)
            myret = True
        End Using
        Return myret
    End Function

    Public ReadOnly Property sortField As String Implements IModel.sortField
        Get
            'Return "username"
            Return "fullname"
        End Get
    End Property

    Public ReadOnly Property tablename As String Implements IModel.tablename
        Get
            Return "Sales.sfkam"
        End Get
    End Property

    Public Function PopulateKAM(ByVal location As Integer) As List(Of KAMModel)
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim KAMList As New List(Of KAMModel)
        Dim DS As DataSet = New DataSet
        Using conn As Object = myadapter.getConnection
            conn.Open()

            Dim sqlstr = String.Format("select * from sales.sfkam where location = {0} order by username", location)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, tablename)

            For Each dr As DataRow In DS.Tables(0).Rows
                KAMList.Add(New KAMModel With {.fullname = dr.Item("fullname"),
                                               .location = dr.Item("location"),
                                               .userid = dr.Item("userid"),
                                               .username = dr.Item("username")})
            Next
        End Using
        Return KAMList
    End Function

    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        AddHandler dataadapter.RowUpdated, AddressOf myadapter.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Object = myadapter.getConnection
            conn.Open()
            mytransaction = conn.BeginTransaction
            'Update
            Dim sqlstr = "sales.sp_updatesfkam"
            dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "username").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "location").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "fullname").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive").SourceVersion = DataRowVersion.Current

            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sales.sp_insertsfkam"
            dataadapter.InsertCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "username").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "location").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "fullname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sales.sp_deletesfkam"
            dataadapter.DeleteCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").Direction = ParameterDirection.Input
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(TableName))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function
End Class
