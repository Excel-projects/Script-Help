<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Settings
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.pgdSettings = New System.Windows.Forms.PropertyGrid()
        Me.SuspendLayout()
        '
        'pgdSettings
        '
        Me.pgdSettings.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pgdSettings.LineColor = System.Drawing.SystemColors.ControlDark
        Me.pgdSettings.Location = New System.Drawing.Point(0, 0)
        Me.pgdSettings.Name = "pgdSettings"
        Me.pgdSettings.Size = New System.Drawing.Size(650, 750)
        Me.pgdSettings.TabIndex = 3
        '
        'Settings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.pgdSettings)
        Me.Name = "Settings"
        Me.Size = New System.Drawing.Size(650, 750)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pgdSettings As Windows.Forms.PropertyGrid
End Class
