Imports System.Windows.Forms
Imports System.Threading
Imports System.ComponentModel

Public Class DialogTBParamDetail
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myCountry As country
    Dim myControl As New TBParamDetailController
    Dim ExchangeRate As Decimal


    Public Sub New(myCountry As country)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.myCountry = myCountry
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loadData()
    End Sub

    Private Sub loadData()
        If Not myThread.IsAlive Then

            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub
    Public Overloads Function validate() As Boolean
        Dim myreturn As Boolean = True
        ErrorProvider1.SetError(TextBox1, "")
        If Not IsNumeric(TextBox1.Text) Then
            myreturn = False
            ErrorProvider1.SetError(TextBox1, "Value is not valid.")
        End If
        Return myreturn
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.validate Then

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            myControl.Model.UpdateCurrency(myCountry, "EX-Rate", TextBox1.Text)
            Me.Close()
        End If
       
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public Sub DoWork()
        'ExchangeRate
        ExchangeRate = myControl.Model.getCurrency(myCountry, "EX-Rate")
        ProgressReport(4, "ExchangeRate")
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    'ToolStripStatusLabel1.Text = message
                Case 4
                    TextBox1.Text = ExchangeRate
            End Select
        End If
    End Sub
End Class

