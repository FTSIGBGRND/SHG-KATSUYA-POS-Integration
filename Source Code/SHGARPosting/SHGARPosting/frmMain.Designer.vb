<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.lblToday = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.bgwIntegrate = New System.ComponentModel.BackgroundWorker()
        Me.timerToday = New System.Windows.Forms.Timer(Me.components)
        Me.bgwInitialize = New System.ComponentModel.BackgroundWorker()
        Me.timerReconnect = New System.Windows.Forms.Timer(Me.components)
        Me.timerStarter = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'lblToday
        '
        Me.lblToday.AutoSize = True
        Me.lblToday.Location = New System.Drawing.Point(320, 9)
        Me.lblToday.Name = "lblToday"
        Me.lblToday.Size = New System.Drawing.Size(37, 13)
        Me.lblToday.TabIndex = 3
        Me.lblToday.Text = "Today"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(15, 164)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(46, 13)
        Me.lblStatus.TabIndex = 2
        Me.lblStatus.Text = "Status..."
        '
        'bgwIntegrate
        '
        '
        'timerToday
        '
        Me.timerToday.Interval = 1000
        '
        'bgwInitialize
        '
        '
        'timerReconnect
        '
        Me.timerReconnect.Interval = 1800000
        '
        'timerStarter
        '
        Me.timerStarter.Interval = 1000
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(490, 186)
        Me.Controls.Add(Me.lblToday)
        Me.Controls.Add(Me.lblStatus)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "AR Posting"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblToday As Label
    Friend WithEvents lblStatus As Label
    Friend WithEvents bgwIntegrate As System.ComponentModel.BackgroundWorker
    Friend WithEvents timerToday As Timer
    Friend WithEvents bgwInitialize As System.ComponentModel.BackgroundWorker
    Friend WithEvents timerReconnect As Timer
    Friend WithEvents timerStarter As Timer
End Class
