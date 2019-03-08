Imports System.Threading
Public Class FormImportRAWDATAMS

    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Private OpenFileDialog1 As New OpenFileDialog
    Dim startperiod As Date
    Dim endperiod As Date

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            If MessageBox.Show("Please check your Start and End Period! Make sure the Period Range is correct. Continue?", "Continue", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.OK Then
                If OpenFileDialog1.ShowDialog = DialogResult.OK Then
                    startperiod = CDate(String.Format("{0:yyyy-MM-01}", DateTimePicker1.Value.Date))
                    endperiod = CDate(String.Format("{0:yyyy-MM-01}", DateTimePicker2.Value.Date))
                    myThread = New Thread(AddressOf DoWork)
                    myThread.Start()
                End If
            End If
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoWork()
        Dim ImportSalesForecastHK1 = New ImportMSRAWDATA(Me, OpenFileDialog1.FileName)
        Dim sw As New Stopwatch
        sw.Start()
        ProgressReport(1, "Processing. Please wait..")
        ProgressReport(2, "Marque")
        If ImportSalesForecastHK1.ValidateFile Then
            ProgressReport(1, "Importing File. Please wait...")
            If ImportSalesForecastHK1.doImportFile(startperiod, endperiod) Then
                'Thread.Sleep(5000)               
                sw.Stop()
                ProgressReport(1, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            Else
                ProgressReport(1, String.Format("Error::{0}", ImportSalesForecastHK1.ErrorMsg))
            End If
            ProgressReport(3, "Continuous")
        Else
            ProgressReport(1, String.Format("Error found: {0}", ImportSalesForecastHK1.ErrorMsg))
            ProgressReport(3, "Continuous")
        End If
        sw.Stop()

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

    Private Sub FormMSImport_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class