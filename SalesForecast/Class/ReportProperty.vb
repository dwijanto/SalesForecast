Public MustInherit Class ReportProperty
    Public Property ColumnStartKey As Integer = 1
    Public Property ColumnStartData As Integer = 7
    Public Property RowStartDataI As Integer = 7
    Public Property RowStartdData As String = "A7"
End Class

Public Class HKReportProperty
    Inherits ReportProperty
    Private Shared Property _instance As HKReportProperty

    Public Property FGColumnStartKey As Integer = 1
    Public Property FGColumnStartData As Integer = 7
    Public Property FGRowStartDataI As Integer = 7

    Public Property RowStartData As String = "A7"
    Public Property FGRowStartData As String = "A7"

    Public Shared Function getInstance()
        If IsNothing(_instance) Then
            _instance = New HKReportProperty
        End If
        Return _instance
    End Function

    Private Sub New()
        MyBase.new()
    End Sub
End Class
'Public Class HKGroupReportProperty
'    Inherits ReportProperty
'    Private Shared Property _instance As HKGroupReportProperty

'    Public Property FGColumnStartKey As Integer = 1
'    Public Property FGColumnStartData As Integer = 7
'    Public Property FGRowStartDataI As Integer = 7

'    Public Property RowStartData As String = "A7"

'    Public Shared Function getInstance()
'        If IsNothing(_instance) Then
'            _instance = New HKGroupReportProperty
'        End If
'        Return _instance
'    End Function

'    Private Sub New()
'        MyBase.New()
'    End Sub
'End Class

Public Class TWReportProperty
    Inherits ReportProperty
    Private Shared Property _instance As TWReportProperty

    Public Shared Function getInstance()
        If IsNothing(_instance) Then
            _instance = New TWReportProperty
        End If
        Return _instance
    End Function

    Private Sub New()
        MyBase.New()
    End Sub
End Class
Public Class MSReportProperty
    Inherits ReportProperty
    Private Shared Property _instance As MSReportProperty

    Public Shared Function getInstance()
        If IsNothing(_instance) Then
            _instance = New MSReportProperty
        End If
        Return _instance
    End Function

    Private Sub New()
        MyBase.New()
    End Sub
End Class

Public Class SGReportProperty
    Inherits ReportProperty
    Private Shared Property _instance As SGReportProperty

    Public Shared Function getInstance()
        If IsNothing(_instance) Then
            _instance = New SGReportProperty
        End If
        Return _instance
    End Function

    Private Sub New()
        MyBase.New()
    End Sub
End Class

Public Class THReportProperty
    Inherits ReportProperty
    Private Shared Property _instance As THReportProperty

    Public Shared Function getInstance()
        If IsNothing(_instance) Then
            _instance = New THReportProperty
        End If
        Return _instance
    End Function

    Private Sub New()
        MyBase.New()
    End Sub
End Class