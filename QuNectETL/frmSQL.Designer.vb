﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
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
        Me.txtSQL = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnSQLDone = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtSQL
        '
        Me.txtSQL.Location = New System.Drawing.Point(26, 46)
        Me.txtSQL.MaxLength = 65536
        Me.txtSQL.Multiline = True
        Me.txtSQL.Name = "txtSQL"
        Me.txtSQL.Size = New System.Drawing.Size(705, 392)
        Me.txtSQL.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(22, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(294, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Enter your source SQL Statement below"
        '
        'btnSQLDone
        '
        Me.btnSQLDone.Location = New System.Drawing.Point(320, 462)
        Me.btnSQLDone.Name = "btnSQLDone"
        Me.btnSQLDone.Size = New System.Drawing.Size(100, 44)
        Me.btnSQLDone.TabIndex = 2
        Me.btnSQLDone.Text = "Done"
        Me.btnSQLDone.UseVisualStyleBackColor = True
        '
        'frmSQL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(768, 537)
        Me.Controls.Add(Me.btnSQLDone)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtSQL)
        Me.Name = "frmSQL"
        Me.Text = "SQL Statement"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtSQL As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnSQLDone As Button
End Class