<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSQL
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSQL))
        Me.lblEnterSQL = New System.Windows.Forms.Label()
        Me.btnSQLDone = New System.Windows.Forms.Button()
        Me.txtSQL = New System.Windows.Forms.RichTextBox()
        Me.SuspendLayout()
        '
        'lblEnterSQL
        '
        Me.lblEnterSQL.AutoSize = True
        Me.lblEnterSQL.Location = New System.Drawing.Point(15, 6)
        Me.lblEnterSQL.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblEnterSQL.Name = "lblEnterSQL"
        Me.lblEnterSQL.Size = New System.Drawing.Size(196, 13)
        Me.lblEnterSQL.TabIndex = 1
        Me.lblEnterSQL.Text = "Enter your source SQL Statement below"
        '
        'btnSQLDone
        '
        Me.btnSQLDone.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnSQLDone.Location = New System.Drawing.Point(207, 312)
        Me.btnSQLDone.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnSQLDone.Name = "btnSQLDone"
        Me.btnSQLDone.Size = New System.Drawing.Size(67, 29)
        Me.btnSQLDone.TabIndex = 2
        Me.btnSQLDone.Text = "Done"
        Me.btnSQLDone.UseVisualStyleBackColor = True
        '
        'txtSQL
        '
        Me.txtSQL.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSQL.Location = New System.Drawing.Point(18, 22)
        Me.txtSQL.Name = "txtSQL"
        Me.txtSQL.Size = New System.Drawing.Size(482, 271)
        Me.txtSQL.TabIndex = 3
        Me.txtSQL.Text = ""
        '
        'frmSQL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(512, 349)
        Me.Controls.Add(Me.txtSQL)
        Me.Controls.Add(Me.btnSQLDone)
        Me.Controls.Add(Me.lblEnterSQL)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Name = "frmSQL"
        Me.Text = "SQL Statement"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblEnterSQL As Label
    Friend WithEvents btnSQLDone As Button
    Friend WithEvents txtSQL As RichTextBox
End Class
