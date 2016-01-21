<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.components = New System.ComponentModel.Container()
        Me.btnStartStop = New System.Windows.Forms.Button()
        Me.lblRunning = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnStartStop
        '
        Me.btnStartStop.Location = New System.Drawing.Point(221, 239)
        Me.btnStartStop.Name = "btnStartStop"
        Me.btnStartStop.Size = New System.Drawing.Size(199, 63)
        Me.btnStartStop.TabIndex = 0
        Me.btnStartStop.Text = "Start/Stop"
        Me.btnStartStop.UseVisualStyleBackColor = True
        '
        'lblRunning
        '
        Me.lblRunning.AutoSize = True
        Me.lblRunning.Font = New System.Drawing.Font("Microsoft Sans Serif", 19.97906!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRunning.Location = New System.Drawing.Point(210, 161)
        Me.lblRunning.Name = "lblRunning"
        Me.lblRunning.Size = New System.Drawing.Size(0, 63)
        Me.lblRunning.TabIndex = 1
        Me.lblRunning.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(276, 32)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(400, 31)
        Me.TextBox1.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 35)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(257, 26)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "File To Watch (Full Path):"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(688, 373)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.lblRunning)
        Me.Controls.Add(Me.btnStartStop)
        Me.Name = "frmMain"
        Me.Text = "Log File Monitor"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnStartStop As System.Windows.Forms.Button
    Friend WithEvents lblRunning As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class
