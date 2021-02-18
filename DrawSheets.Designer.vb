<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DrawSheets
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
        Me.DateCheckTimer = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'DateCheckTimer
        '
        Me.DateCheckTimer.Enabled = True
        Me.DateCheckTimer.Interval = 3000
        '
        'DrawSheets
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(2539, 1281)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "DrawSheets"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DateCheckTimer As Timer
End Class
