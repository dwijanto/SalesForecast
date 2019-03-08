Imports System.Threading

Public Class FormTWImport

    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Private OpenFileDialog1 As New OpenFileDialog
    Dim period As Date

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            If OpenFileDialog1.ShowDialog = DialogResult.OK Then
                period = DateTimePicker1.Value.Date
                myThread = New Thread(AddressOf DoWork)
                myThread.Start()
            End If
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoWork()
        Dim ImportSalesForecastHK1 = New ImportSalesForecastHK(Me, OpenFileDialog1.FileName)
        ProgressReport(1, "Processing. Please wait..")
        ProgressReport(2, "Marque")
        If ImportSalesForecastHK1.ValidateFile Then

            If ImportSalesForecastHK1.DoImportFile(period) Then
                'Thread.Sleep(5000)
                ProgressReport(1, "Done.")
            Else
                ProgressReport(1, String.Format("Error::{0}", ImportSalesForecastHK1.ErrorMsg))
            End If
            ProgressReport(3, "Continuous")
        End If
    End Sub
    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.ToolStripStatusLabel1.Text = message
                Case 2
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee

                Case 3
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
            End Select

        End If

    End Sub
End Class