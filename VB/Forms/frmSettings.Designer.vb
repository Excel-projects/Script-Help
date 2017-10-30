<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSettings
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
        Me.pgdSettings = New System.Windows.Forms.PropertyGrid()
        Me.SuspendLayout()
        '
        'pgdSettings
        '
        Me.pgdSettings.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pgdSettings.Location = New System.Drawing.Point(0, 0)
        Me.pgdSettings.Name = "pgdSettings"
        Me.pgdSettings.Size = New System.Drawing.Size(586, 416)
        Me.pgdSettings.TabIndex = 0
        '
        'frmSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(586, 416)
        Me.Controls.Add(Me.pgdSettings)
        Me.Name = "frmSettings"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "frmSettings"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pgdSettings As System.Windows.Forms.PropertyGrid
End Class
