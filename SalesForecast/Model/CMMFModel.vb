Imports Npgsql

Public Class CMMFModel
    Implements IModel



    Dim myadapter As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "sales.sfcmmf"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "cmmf"
        End Get
    End Property

    Public ReadOnly Property FilterField
        Get
            Return "[cmmf] like '*{0}*' or [reference] like '*{0}*' or [origin] like '*{0}*' or [description] like '*{0}*'  or [familyname] like '*{0}*' or [productlinegpsname] like '*{0}*'"
        End Get
    End Property

    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u.cmmf::text,u.origin,u.sbu,u.brand,u.reference,u.productdesc,u.description,u.companystatus,u.status,u.productsegmentation,u.productlinegpsid,u.familyid,u.activedate,u.launchingmonth,u.remarks,to_char(f.familyid,'FM000') || ' - ' || f.familyname as familyname,f.familylv2,sbu.productlinegpsname,p.nsp1,p.nsp2 from {0} u left join sales.sffamily f on f.familyid = u.familyid left join sales.sfproductlinegps sbu on sbu.productlinegpsid = u.productlinegpsid left join sales.sfcmmfnsp p on p.cmmf = u.cmmf  order by {1}", TableName, SortField)
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
            Dim sqlstr = "sales.sp_updatesfcmmf"
            dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "origin").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "brand").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "reference").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "companystatus").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "status").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "productsegmentation").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "productlinegpsid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "activedate").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "launchingmonth").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remarks").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sales.sp_insertsfcmmf"
            dataadapter.InsertCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "origin").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "brand").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "reference").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "companystatus").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "status").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "productsegmentation").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "productlinegpsid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "activedate").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "launchingmonth").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remarks").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sales.sp_deletesfcmmf"
            dataadapter.DeleteCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").Direction = ParameterDirection.Input
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
