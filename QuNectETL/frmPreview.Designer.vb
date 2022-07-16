<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPreview
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
        Me.dgvPreview = New System.Windows.Forms.DataGridView()
        CType(Me.dgvPreview, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvPreview
        '
        Me.dgvPreview.AllowUserToAddRows = False
        Me.dgvPreview.AllowUserToDeleteRows = False
        Me.dgvPreview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPreview.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvPreview.Location = New System.Drawing.Point(0, 0)
        Me.dgvPreview.Name = "dgvPreview"
        Me.dgvPreview.ReadOnly = True
        Me.dgvPreview.Size = New System.Drawing.Size(800, 450)
        Me.dgvPreview.TabIndex = 0
        '
        'frmPreview
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.dgvPreview)
        Me.Name = "frmPreview"
        Me.Text = "Preview"
        CType(Me.dgvPreview, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents dgvPreview As DataGridView
End Class
