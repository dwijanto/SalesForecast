Imports System.Windows.Forms

Public Class DialogSelectPeriod
    Public startperiod As Date
    Public endperiod As Date

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        startperiod = DateTimePicker1.Value.Date
        endperiod = DateTimePicker2.Value.Date
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker2.Value = DateTimePicker1.Value.AddMonths(6)
    End Sub

    Private Sub DialogSelectPeriod_Load(sender As Object, e As EventArgs) Handles Me.Load
        DateTimePicker1.Value = DateTimePicker1.Value.AddMonths(6)
        DateTimePicker2.Value = DateTimePicker1.Value.AddMonths(5)
    End Sub
End Class
