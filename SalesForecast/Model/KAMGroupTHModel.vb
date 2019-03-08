Imports Npgsql
Public Class KAMGroupTHModel
    Implements IModel
    Dim myadapter As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance

    'Public Property groupid As Long
    'Public Property groupname As String

    Public Property kam As String    
    Public Property recid As Long
    Public Property account As String
    Public Property mla As String
    Public Property assignment As String
    Public Property assigmentid As Long
    Public Property mlacardnameid As Long
    Public Property gpcw As Decimal
    Public Property gpsda As Decimal
    Public Property ifrscw As Decimal
    Public Property ifrssda As Decimal


    Dim KAMGroupTHList As List(Of KAMGroupTHModel)

    Public Function LoadData(DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u.* from {0}  order by {1}", tablename, sortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, tablename)
            myret = True
        End Using
        Return myret
    End Function


    Public Function PopulateKAMGroupTH(criteria As String) As List(Of KAMGroupTHModel)
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim DS As DataSet = New DataSet
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("with ka as (select distinct kam.username,mc.account,mc.mla,mc.account || '_' ||  trim(mc.mla) as assignment, ka.id as assignmentid " &
                                       " ,mc.id as mlacardnameid,mg.gpcw,mg.gpsda,mg.ifrscw,mg.ifrssda  " &
                                        " from sales.sfkam kam  left join sales.sfkamassignmentth ka on ka.kam = kam.username  " &
                                        " left join sales.sfmlacardnameth mc on mc.id = ka.mlacardnameid" &
                                        " left join sales.sfmlagroupth mg on mg.mla = mc.mla" &
                                        " where kam.username = '{0}' order by mc.mla) " &
                                        " select 1 as id,row_number() over (order by ka.account, ka.mla) as recid,* from ka;", criteria)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, tablename)
            KAMGroupTHList = New List(Of KAMGroupTHModel)
            For Each dr As DataRow In DS.Tables(0).Rows

                KAMGroupTHList.Add(New KAMGroupTHModel With {.recid = dr.Item("recid"),
                                                                   .kam = dr.Item("username"),
                                                                   .account = "" & dr.Item("account"),
                                                                    .mla = "" & dr.Item("mla"),
                                                                    .assignment = "" & dr.Item("assignment"),
                                                                    .assigmentid = dr.Item("assignmentid"),
                                                                    .mlacardnameid = dr.Item("mlacardnameid"),
                                                                    .gpcw = dr.Item("gpcw"),
                                                                    .gpsda = dr.Item("gpsda"),
                                                                    .ifrscw = dr.Item("ifrscw"),
                                                                    .ifrssda = dr.Item("ifrssda")})
            Next
        End Using
        Return KAMGroupTHList
    End Function

    Public Function Add(ByVal model As KAMGroupMSModel) As Long
        Dim myret As Long
        Dim sqlstr = String.Format("insert into sales.sfkamgroupth(groupid,kam) values('{0}','{1}'); select currval('sales.sfkamgroupms_id_seq');", model.groupid, model.kam)
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
            Return "Sales.sfkamassignmentth"
        End Get
    End Property
End Class
