<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormForecastGroupMS
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.StatusStrip1.SuspendLayout
        Me.SuspendLayout
        '
        'Label1
        '
        Me.Label1.AutoSize = true
        Me.Label1.Location = New System.Drawing.Point(185, 35)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 25
        Me.Label1.Text = "KAM"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = true
        Me.ComboBox1.Location = New System.Drawing.Point(221, 32)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(178, 21)
        Me.ComboBox1.TabIndex = 24
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(221, 107)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(178, 36)
        Me.Button1.TabIndex = 21
        Me.Button1.Text = "Generate Template"
        Me.Button1.UseVisualStyleBackColor = true
        '
        'Label2
        '
        Me.Label2.AutoSize = true
        Me.Label2.Location = New System.Drawing.Point(178, 65)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(37, 13)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "Period"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "yyyy-MM"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(221, 59)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(178, 20)
        Me.DateTimePicker1.TabIndex = 19
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel2, Me.ToolStripProgressBar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 186)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(647, 22)
        Me.StatusStrip1.TabIndex = 22
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(530, 17)
        Me.ToolStripStatusLabel1.Spring = true
        Me.ToolStripStatusLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(0, 17)
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(100, 16)
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = true
        Me.CheckBox1.Location = New System.Drawing.Point(221, 84)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(100, 17)
        Me.CheckBox1.TabIndex = 26
        Me.CheckBox1.Text = "Blank Template"
        Me.CheckBox1.UseVisualStyleBackColor = true
        Me.CheckBox1.Visible = false
        '
        'FormForecastGroupMS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(647, 208)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Name = "FormForecastGroupMS"
        Me.Text = "FormForecastGroupMS"
        Me.StatusStrip1.ResumeLayout(false)
        Me.StatusStrip1.PerformLayout
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Public WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Public WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
End Class
