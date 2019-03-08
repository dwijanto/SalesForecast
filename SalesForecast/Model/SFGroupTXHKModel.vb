Imports Npgsql
Public Class SFGroupTXHKModel
    Implements IModel

    Dim myadapter As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "sales.sfgrouptxhk"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "txdate,cmmf,username"
        End Get
    End Property

    Public ReadOnly Property FilterField
        Get
            Return "[cmmf] like '*{0}*' or [reference] like '*{0}*'"
        End Get
    End Property


    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u.id, u.txdate,u.cmmf::text,u.salesforecast from {0} u  order by {1}", TableName, SortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function
    Public Function LoadData(ByVal DS As DataSet, ByVal criteria As String) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u.id, u.txdate,u.cmmf::text,u.salesforecast,k.username,c.reference from {0} u left join sales._user _u on _u.id = u.userid left join sales.sfkam k on k.userid = _u.username left join sales.sfcmmf c on c.cmmf = u.cmmf {1} order by {2} ", TableName, criteria, SortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
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
            Dim sqlstr = "sales.sp_updatesfgrouptxhk"
            dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original            
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "salesforecast").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            'dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            'dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(TableName))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function
End Class
