Public Class PeriodRange
    Private startPeriod As Date
    Private Periodlength As Integer
    Dim periodDict As Dictionary(Of Integer, Date)

    Public Sub New(ByVal startPeriod As Date, ByVal Periodlength As Integer)
        Me.startPeriod = startPeriod
        Me.Periodlength = Periodlength
    End Sub

    Public Function getPeriod() As Dictionary(Of Integer, Date)
        periodDict = New Dictionary(Of Integer, Date)
        Dim myyear = Year(startPeriod)
        Dim mymonth = Month(startPeriod)
        For i = 0 To Periodlength - 1
            'startPeriod = startPeriod.AddMonths(1)
            'periodDict.Add(i, CDate(String.Format("{0}-{1}-01", myYear, myMonth)))
            periodDict.Add(i, CDate(String.Format("{0}-{1}-01", startPeriod.Year, startPeriod.Month)))
            startPeriod = startPeriod.AddMonths(1)
            'mymonth = mymonth + 1
            'If mymonth = 13 Then
            '    mymonth = 1
            '    myyear = myyear + 1
            'End If
        Next
        Return periodDict
    End Function

End Class
