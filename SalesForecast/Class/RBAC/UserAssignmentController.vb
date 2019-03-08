Public Class UserAssignmentController
    Public Function getRoles() As List(Of SalesForecast.Item)
        Dim RBAC = New DbManager
        Return RBAC.getRoles()
    End Function
End Class
