<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmCopy
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCopy))
        Me.dgMapping = New System.Windows.Forms.DataGridView()
        Me.Source = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Destination = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.btnDestination = New System.Windows.Forms.Button()
        Me.pb = New System.Windows.Forms.ProgressBar()
        Me.lblAppToken = New System.Windows.Forms.Label()
        Me.txtAppToken = New System.Windows.Forms.TextBox()
        Me.lblServer = New System.Windows.Forms.Label()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.lblUsername = New System.Windows.Forms.Label()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.lblDestinationTable = New System.Windows.Forms.Label()
        Me.ckbDetectProxy = New System.Windows.Forms.CheckBox()
        Me.btnListFields = New System.Windows.Forms.Button()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.lblCatalog = New System.Windows.Forms.Label()
        Me.cmbPassword = New System.Windows.Forms.ComboBox()
        Me.btnAppToken = New System.Windows.Forms.Button()
        Me.btnUserToken = New System.Windows.Forms.Button()
        Me.btnSQLStatement = New System.Windows.Forms.Button()
        Me.lblDSN = New System.Windows.Forms.Label()
        Me.cmbDSN = New System.Windows.Forms.ComboBox()
        CType(Me.dgMapping, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgMapping
        '
        Me.dgMapping.AllowUserToAddRows = False
        Me.dgMapping.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgMapping.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgMapping.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgMapping.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Source, Me.Destination})
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgMapping.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgMapping.Location = New System.Drawing.Point(12, 251)
        Me.dgMapping.Name = "dgMapping"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgMapping.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgMapping.Size = New System.Drawing.Size(522, 716)
        Me.dgMapping.TabIndex = 0
        '
        'Source
        '
        Me.Source.HeaderText = "Source"
        Me.Source.Name = "Source"
        Me.Source.ReadOnly = True
        Me.Source.Width = 200
        '
        'Destination
        '
        Me.Destination.HeaderText = "Destination"
        Me.Destination.Name = "Destination"
        Me.Destination.Width = 200
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(349, 212)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(185, 27)
        Me.btnImport.TabIndex = 3
        Me.btnImport.Text = "Copy from Source to Destination"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'btnDestination
        '
        Me.btnDestination.Location = New System.Drawing.Point(12, 188)
        Me.btnDestination.Name = "btnDestination"
        Me.btnDestination.Size = New System.Drawing.Size(167, 23)
        Me.btnDestination.TabIndex = 4
        Me.btnDestination.Text = "Choose Destination Table..."
        Me.btnDestination.UseVisualStyleBackColor = True
        '
        'pb
        '
        Me.pb.Location = New System.Drawing.Point(286, 83)
        Me.pb.Maximum = 1000
        Me.pb.Name = "pb"
        Me.pb.Size = New System.Drawing.Size(248, 23)
        Me.pb.TabIndex = 33
        '
        'lblAppToken
        '
        Me.lblAppToken.AutoSize = True
        Me.lblAppToken.Location = New System.Drawing.Point(15, 64)
        Me.lblAppToken.Name = "lblAppToken"
        Me.lblAppToken.Size = New System.Drawing.Size(148, 13)
        Me.lblAppToken.TabIndex = 30
        Me.lblAppToken.Text = "QuickBase Application Token"
        '
        'txtAppToken
        '
        Me.txtAppToken.Location = New System.Drawing.Point(12, 83)
        Me.txtAppToken.Name = "txtAppToken"
        Me.txtAppToken.Size = New System.Drawing.Size(258, 20)
        Me.txtAppToken.TabIndex = 29
        '
        'lblServer
        '
        Me.lblServer.AutoSize = True
        Me.lblServer.Location = New System.Drawing.Point(360, 11)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.Size = New System.Drawing.Size(93, 13)
        Me.lblServer.TabIndex = 28
        Me.lblServer.Text = "QuickBase Server"
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(361, 30)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(173, 20)
        Me.txtServer.TabIndex = 27
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(188, 30)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(167, 20)
        Me.txtPassword.TabIndex = 25
        '
        'lblUsername
        '
        Me.lblUsername.AutoSize = True
        Me.lblUsername.Location = New System.Drawing.Point(15, 11)
        Me.lblUsername.Name = "lblUsername"
        Me.lblUsername.Size = New System.Drawing.Size(110, 13)
        Me.lblUsername.TabIndex = 24
        Me.lblUsername.Text = "QuickBase Username"
        '
        'txtUsername
        '
        Me.txtUsername.Location = New System.Drawing.Point(12, 30)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(172, 20)
        Me.txtUsername.TabIndex = 23
        '
        'lblDestinationTable
        '
        Me.lblDestinationTable.AutoSize = True
        Me.lblDestinationTable.Location = New System.Drawing.Point(184, 193)
        Me.lblDestinationTable.Name = "lblDestinationTable"
        Me.lblDestinationTable.Size = New System.Drawing.Size(0, 13)
        Me.lblDestinationTable.TabIndex = 34
        '
        'ckbDetectProxy
        '
        Me.ckbDetectProxy.AutoSize = True
        Me.ckbDetectProxy.Location = New System.Drawing.Point(286, 56)
        Me.ckbDetectProxy.Name = "ckbDetectProxy"
        Me.ckbDetectProxy.Size = New System.Drawing.Size(188, 17)
        Me.ckbDetectProxy.TabIndex = 36
        Me.ckbDetectProxy.Text = "Automatically detect proxy settings"
        Me.ckbDetectProxy.UseVisualStyleBackColor = True
        '
        'btnListFields
        '
        Me.btnListFields.Location = New System.Drawing.Point(12, 218)
        Me.btnListFields.Name = "btnListFields"
        Me.btnListFields.Size = New System.Drawing.Size(82, 27)
        Me.btnListFields.TabIndex = 37
        Me.btnListFields.Text = "List Fields"
        Me.btnListFields.UseVisualStyleBackColor = True
        '
        'lblProgress
        '
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(116, 225)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(0, 13)
        Me.lblProgress.TabIndex = 42
        '
        'lblCatalog
        '
        Me.lblCatalog.AutoSize = True
        Me.lblCatalog.Location = New System.Drawing.Point(185, 125)
        Me.lblCatalog.Name = "lblCatalog"
        Me.lblCatalog.Size = New System.Drawing.Size(0, 13)
        Me.lblCatalog.TabIndex = 45
        '
        'cmbPassword
        '
        Me.cmbPassword.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPassword.FormattingEnabled = True
        Me.cmbPassword.Items.AddRange(New Object() {"Please choose...", "QuickBase Password", "QuickBase User Token"})
        Me.cmbPassword.Location = New System.Drawing.Point(188, 8)
        Me.cmbPassword.Name = "cmbPassword"
        Me.cmbPassword.Size = New System.Drawing.Size(143, 21)
        Me.cmbPassword.TabIndex = 46
        '
        'btnAppToken
        '
        Me.btnAppToken.Location = New System.Drawing.Point(160, 60)
        Me.btnAppToken.Name = "btnAppToken"
        Me.btnAppToken.Size = New System.Drawing.Size(19, 20)
        Me.btnAppToken.TabIndex = 79
        Me.btnAppToken.Text = "?"
        Me.btnAppToken.UseVisualStyleBackColor = True
        '
        'btnUserToken
        '
        Me.btnUserToken.Location = New System.Drawing.Point(335, 9)
        Me.btnUserToken.Name = "btnUserToken"
        Me.btnUserToken.Size = New System.Drawing.Size(19, 20)
        Me.btnUserToken.TabIndex = 80
        Me.btnUserToken.Text = "?"
        Me.btnUserToken.UseVisualStyleBackColor = True
        '
        'btnSQLStatement
        '
        Me.btnSQLStatement.Location = New System.Drawing.Point(12, 152)
        Me.btnSQLStatement.Name = "btnSQLStatement"
        Me.btnSQLStatement.Size = New System.Drawing.Size(167, 28)
        Me.btnSQLStatement.TabIndex = 81
        Me.btnSQLStatement.Text = "Enter Source SQL Statement"
        Me.btnSQLStatement.UseVisualStyleBackColor = True
        '
        'lblDSN
        '
        Me.lblDSN.AutoSize = True
        Me.lblDSN.Location = New System.Drawing.Point(15, 107)
        Me.lblDSN.Name = "lblDSN"
        Me.lblDSN.Size = New System.Drawing.Size(30, 13)
        Me.lblDSN.TabIndex = 83
        Me.lblDSN.Text = "DSN"
        '
        'cmbDSN
        '
        Me.cmbDSN.FormattingEnabled = True
        Me.cmbDSN.Location = New System.Drawing.Point(12, 125)
        Me.cmbDSN.Name = "cmbDSN"
        Me.cmbDSN.Size = New System.Drawing.Size(522, 21)
        Me.cmbDSN.TabIndex = 82
        '
        'frmCopy
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(565, 998)
        Me.Controls.Add(Me.lblDSN)
        Me.Controls.Add(Me.cmbDSN)
        Me.Controls.Add(Me.btnSQLStatement)
        Me.Controls.Add(Me.btnUserToken)
        Me.Controls.Add(Me.btnAppToken)
        Me.Controls.Add(Me.cmbPassword)
        Me.Controls.Add(Me.lblCatalog)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.btnListFields)
        Me.Controls.Add(Me.ckbDetectProxy)
        Me.Controls.Add(Me.lblDestinationTable)
        Me.Controls.Add(Me.pb)
        Me.Controls.Add(Me.lblAppToken)
        Me.Controls.Add(Me.txtAppToken)
        Me.Controls.Add(Me.lblServer)
        Me.Controls.Add(Me.txtServer)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.lblUsername)
        Me.Controls.Add(Me.txtUsername)
        Me.Controls.Add(Me.btnDestination)
        Me.Controls.Add(Me.btnImport)
        Me.Controls.Add(Me.dgMapping)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmCopy"
        Me.Text = "QuNect ETL"
        CType(Me.dgMapping, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgMapping As System.Windows.Forms.DataGridView
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents btnDestination As System.Windows.Forms.Button
    Friend WithEvents pb As System.Windows.Forms.ProgressBar
    Friend WithEvents lblAppToken As System.Windows.Forms.Label
    Friend WithEvents txtAppToken As System.Windows.Forms.TextBox
    Friend WithEvents lblServer As System.Windows.Forms.Label
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents lblUsername As System.Windows.Forms.Label
    Friend WithEvents txtUsername As System.Windows.Forms.TextBox
    Friend WithEvents lblDestinationTable As System.Windows.Forms.Label
    Friend WithEvents ckbDetectProxy As System.Windows.Forms.CheckBox
    Friend WithEvents btnListFields As System.Windows.Forms.Button
    Friend WithEvents Source As DataGridViewTextBoxColumn
    Friend WithEvents Destination As DataGridViewComboBoxColumn
    Friend WithEvents lblProgress As Label
    Friend WithEvents lblCatalog As Label
    Friend WithEvents cmbPassword As ComboBox
    Friend WithEvents btnAppToken As Button
    Friend WithEvents btnUserToken As Button
    Friend WithEvents btnSQLStatement As Button
    Friend WithEvents lblDSN As Label
    Friend WithEvents cmbDSN As ComboBox
End Class
