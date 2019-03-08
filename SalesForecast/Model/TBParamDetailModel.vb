Public Enum country As Integer
    HK = 1
    TW = 2
    MS = 3
    SGD = 4
End Enum
Public Class TBParamDetailModel
    Implements IModel

    Dim myadapter As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance

    Public Function getCurrency(myCountry As country, paramname As String) As Decimal
        Dim sqlstr = String.Empty
        sqlstr = String.Format("select d.nvalue from {0} d left join sales.sfparamhd h on h.paramhdid = d.paramhdid  where h.paramname = '{1}' and h.cvalue = '{2}'", tablename, paramname, myCountry.ToString)

        Dim ra As Decimal
        myadapter.ExecuteScalar(sqlstr, recordAffected:=ra)
        Return ra
    End Function

    Public Function UpdateCurrency(myCountry As country, paramname As String, newvalue As Decimal) As Boolean
        Dim sqlstr = String.Empty
        sqlstr = String.Format("update {0} d set nvalue = {1} from sales.sfparamhd h where h.paramname = '{2}' and h.cvalue = '{3}' and d.paramhdid = h.paramhdid", tablename, newvalue, paramname, myCountry.ToString)
        myadapter.ExecuteNonQuery(sqlstr)
        Return True
    End Function

    Public Function LoadData(DS As DataSet) As Boolean Implements IModel.LoadData
        Return Nothing
    End Function

    Public Function save(obj As Object, mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Return Nothing
    End Function

    Public ReadOnly Property sortField As String Implements IModel.sortField
        Get
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property tablename As String Implements IModel.tablename
        Get
            Return "sales.sfparamdt"
        End Get
    End Property
End Class
