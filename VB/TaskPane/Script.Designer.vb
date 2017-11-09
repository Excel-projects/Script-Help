<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Script
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
        Me.toolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.btnSave = New System.Windows.Forms.ToolStripButton()
        Me.btnCopy = New System.Windows.Forms.ToolStripButton()
        Me.txtScript = New System.Windows.Forms.RichTextBox()
        Me.toolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'toolStrip1
        '
        Me.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.toolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnSave, Me.btnCopy})
        Me.toolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.toolStrip1.Name = "toolStrip1"
        Me.toolStrip1.Size = New System.Drawing.Size(650, 25)
        Me.toolStrip1.TabIndex = 4
        Me.toolStrip1.Text = "toolStrip1"
        '
        'btnSave
        '
        Me.btnSave.Image = Global.ScriptHelp.My.Resources.Resources.Save
        Me.btnSave.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(51, 22)
        Me.btnSave.Text = "Save"
        Me.btnSave.ToolTipText = "Save script file"
        '
        'btnCopy
        '
        Me.btnCopy.Image = Global.ScriptHelp.My.Resources.Resources.Copy
        Me.btnCopy.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnCopy.Name = "btnCopy"
        Me.btnCopy.Size = New System.Drawing.Size(55, 22)
        Me.btnCopy.Text = "Copy"
        Me.btnCopy.ToolTipText = "Copy script text"
        '
        'txtScript
        '
        Me.txtScript.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtScript.BackColor = System.Drawing.Color.White
        Me.txtScript.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtScript.Location = New System.Drawing.Point(3, 30)
        Me.txtScript.Name = "txtScript"
        Me.txtScript.Size = New System.Drawing.Size(644, 719)
        Me.txtScript.TabIndex = 3
        Me.txtScript.Text = ""
        '
        'Script
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.toolStrip1)
        Me.Controls.Add(Me.txtScript)
        Me.Name = "Script"
        Me.Size = New System.Drawing.Size(650, 750)
        Me.toolStrip1.ResumeLayout(False)
        Me.toolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private WithEvents toolStrip1 As Windows.Forms.ToolStrip
    Private WithEvents btnSave As Windows.Forms.ToolStripButton
    Private WithEvents btnCopy As Windows.Forms.ToolStripButton
    Private WithEvents txtScript As Windows.Forms.RichTextBox
End Class
