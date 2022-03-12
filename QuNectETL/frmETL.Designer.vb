<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmETL
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmETL))
        Me.dgMapping = New System.Windows.Forms.DataGridView()
        Me.Source = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Destination = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.btnDestination = New System.Windows.Forms.Button()
        Me.btnListFields = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.saveDialog = New System.Windows.Forms.SaveFileDialog()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.openFile = New System.Windows.Forms.OpenFileDialog()
        Me.lblJobFile = New System.Windows.Forms.Label()
        Me.ckbLogSQL = New System.Windows.Forms.CheckBox()
        Me.TabControl = New System.Windows.Forms.TabControl()
        Me.TabPageSource = New System.Windows.Forms.TabPage()
        Me.GroupBoxSQL = New System.Windows.Forms.GroupBox()
        Me.btnSourceTable = New System.Windows.Forms.Button()
        Me.txtSQL = New System.Windows.Forms.RichTextBox()
        Me.GroupBoxSource = New System.Windows.Forms.GroupBox()
        Me.txtSourcePWD = New System.Windows.Forms.TextBox()
        Me.txtSourceUID = New System.Windows.Forms.TextBox()
        Me.lblSourcePWD = New System.Windows.Forms.Label()
        Me.lblSourceUID = New System.Windows.Forms.Label()
        Me.rdbSourceConnectionString = New System.Windows.Forms.RadioButton()
        Me.rdbSourceDSN = New System.Windows.Forms.RadioButton()
        Me.txtSourceConnectionString = New System.Windows.Forms.TextBox()
        Me.cmbSourceDSN = New System.Windows.Forms.ComboBox()
        Me.TabPageDestination = New System.Windows.Forms.TabPage()
        Me.GroupBoxDestination = New System.Windows.Forms.GroupBox()
        Me.lblDestinationPWD = New System.Windows.Forms.Label()
        Me.lblDestinationUID = New System.Windows.Forms.Label()
        Me.txtDestinationPWD = New System.Windows.Forms.TextBox()
        Me.txtDestinationUID = New System.Windows.Forms.TextBox()
        Me.rdbDestinationDSN = New System.Windows.Forms.RadioButton()
        Me.rdbDestinationConnectionString = New System.Windows.Forms.RadioButton()
        Me.txtDestinationConnectionString = New System.Windows.Forms.TextBox()
        Me.cmbDestinationDSN = New System.Windows.Forms.ComboBox()
        Me.lblDestinationTable = New System.Windows.Forms.Label()
        Me.TabPageMapping = New System.Windows.Forms.TabPage()
        Me.lblProgress = New System.Windows.Forms.Label()
        CType(Me.dgMapping, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl.SuspendLayout()
        Me.TabPageSource.SuspendLayout()
        Me.GroupBoxSQL.SuspendLayout()
        Me.GroupBoxSource.SuspendLayout()
        Me.TabPageDestination.SuspendLayout()
        Me.GroupBoxDestination.SuspendLayout()
        Me.TabPageMapping.SuspendLayout()
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
        Me.dgMapping.Location = New System.Drawing.Point(21, 130)
        Me.dgMapping.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.dgMapping.Name = "dgMapping"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgMapping.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgMapping.RowHeadersWidth = 62
        Me.dgMapping.Size = New System.Drawing.Size(904, 697)
        Me.dgMapping.TabIndex = 0
        '
        'Source
        '
        Me.Source.HeaderText = "Source"
        Me.Source.MinimumWidth = 8
        Me.Source.Name = "Source"
        Me.Source.ReadOnly = True
        Me.Source.Width = 200
        '
        'Destination
        '
        Me.Destination.HeaderText = "Destination"
        Me.Destination.MinimumWidth = 8
        Me.Destination.Name = "Destination"
        Me.Destination.Width = 200
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(514, 24)
        Me.btnImport.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(278, 42)
        Me.btnImport.TabIndex = 3
        Me.btnImport.Text = "Copy from Source to Destination"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'btnDestination
        '
        Me.btnDestination.Location = New System.Drawing.Point(28, 242)
        Me.btnDestination.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnDestination.Name = "btnDestination"
        Me.btnDestination.Size = New System.Drawing.Size(231, 31)
        Me.btnDestination.TabIndex = 4
        Me.btnDestination.Text = "Choose Destination Table..."
        Me.btnDestination.UseVisualStyleBackColor = True
        '
        'btnListFields
        '
        Me.btnListFields.Location = New System.Drawing.Point(21, 26)
        Me.btnListFields.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnListFields.Name = "btnListFields"
        Me.btnListFields.Size = New System.Drawing.Size(123, 42)
        Me.btnListFields.TabIndex = 37
        Me.btnListFields.Text = "List Fields"
        Me.btnListFields.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(396, 25)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(98, 42)
        Me.btnSave.TabIndex = 85
        Me.btnSave.Text = "Save Job"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(161, 27)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(106, 40)
        Me.btnLoad.TabIndex = 86
        Me.btnLoad.Text = "Load Job"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'openFile
        '
        Me.openFile.Filter = "QuNect ETL files|*.job"
        '
        'lblJobFile
        '
        Me.lblJobFile.AutoSize = True
        Me.lblJobFile.Location = New System.Drawing.Point(26, 92)
        Me.lblJobFile.Name = "lblJobFile"
        Me.lblJobFile.Size = New System.Drawing.Size(0, 20)
        Me.lblJobFile.TabIndex = 91
        '
        'ckbLogSQL
        '
        Me.ckbLogSQL.AutoSize = True
        Me.ckbLogSQL.Location = New System.Drawing.Point(292, 32)
        Me.ckbLogSQL.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.ckbLogSQL.Name = "ckbLogSQL"
        Me.ckbLogSQL.Size = New System.Drawing.Size(91, 24)
        Me.ckbLogSQL.TabIndex = 92
        Me.ckbLogSQL.Text = "Log SQL"
        Me.ckbLogSQL.UseVisualStyleBackColor = True
        '
        'TabControl
        '
        Me.TabControl.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl.Controls.Add(Me.TabPageSource)
        Me.TabControl.Controls.Add(Me.TabPageDestination)
        Me.TabControl.Controls.Add(Me.TabPageMapping)
        Me.TabControl.Location = New System.Drawing.Point(28, 23)
        Me.TabControl.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabControl.Name = "TabControl"
        Me.TabControl.SelectedIndex = 0
        Me.TabControl.Size = New System.Drawing.Size(1208, 1249)
        Me.TabControl.TabIndex = 93
        '
        'TabPageSource
        '
        Me.TabPageSource.Controls.Add(Me.GroupBoxSQL)
        Me.TabPageSource.Controls.Add(Me.GroupBoxSource)
        Me.TabPageSource.Location = New System.Drawing.Point(4, 29)
        Me.TabPageSource.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabPageSource.Name = "TabPageSource"
        Me.TabPageSource.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabPageSource.Size = New System.Drawing.Size(1200, 1216)
        Me.TabPageSource.TabIndex = 1
        Me.TabPageSource.Text = "Source"
        Me.TabPageSource.UseVisualStyleBackColor = True
        '
        'GroupBoxSQL
        '
        Me.GroupBoxSQL.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBoxSQL.Controls.Add(Me.btnSourceTable)
        Me.GroupBoxSQL.Controls.Add(Me.txtSQL)
        Me.GroupBoxSQL.Location = New System.Drawing.Point(28, 237)
        Me.GroupBoxSQL.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBoxSQL.Name = "GroupBoxSQL"
        Me.GroupBoxSQL.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBoxSQL.Size = New System.Drawing.Size(1150, 969)
        Me.GroupBoxSQL.TabIndex = 107
        Me.GroupBoxSQL.TabStop = False
        Me.GroupBoxSQL.Text = "Table or SQL Statement"
        '
        'btnSourceTable
        '
        Me.btnSourceTable.Location = New System.Drawing.Point(9, 45)
        Me.btnSourceTable.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnSourceTable.Name = "btnSourceTable"
        Me.btnSourceTable.Size = New System.Drawing.Size(264, 40)
        Me.btnSourceTable.TabIndex = 96
        Me.btnSourceTable.Text = "Choose source table"
        Me.btnSourceTable.UseVisualStyleBackColor = True
        '
        'txtSQL
        '
        Me.txtSQL.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSQL.Location = New System.Drawing.Point(14, 102)
        Me.txtSQL.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtSQL.Name = "txtSQL"
        Me.txtSQL.Size = New System.Drawing.Size(1128, 857)
        Me.txtSQL.TabIndex = 93
        Me.txtSQL.Text = ""
        '
        'GroupBoxSource
        '
        Me.GroupBoxSource.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBoxSource.Controls.Add(Me.txtSourcePWD)
        Me.GroupBoxSource.Controls.Add(Me.txtSourceUID)
        Me.GroupBoxSource.Controls.Add(Me.lblSourcePWD)
        Me.GroupBoxSource.Controls.Add(Me.lblSourceUID)
        Me.GroupBoxSource.Controls.Add(Me.rdbSourceConnectionString)
        Me.GroupBoxSource.Controls.Add(Me.rdbSourceDSN)
        Me.GroupBoxSource.Controls.Add(Me.txtSourceConnectionString)
        Me.GroupBoxSource.Controls.Add(Me.cmbSourceDSN)
        Me.GroupBoxSource.Location = New System.Drawing.Point(28, 23)
        Me.GroupBoxSource.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBoxSource.Name = "GroupBoxSource"
        Me.GroupBoxSource.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBoxSource.Size = New System.Drawing.Size(1150, 195)
        Me.GroupBoxSource.TabIndex = 106
        Me.GroupBoxSource.TabStop = False
        Me.GroupBoxSource.Text = "Choose Source"
        '
        'txtSourcePWD
        '
        Me.txtSourcePWD.Location = New System.Drawing.Point(870, 60)
        Me.txtSourcePWD.Name = "txtSourcePWD"
        Me.txtSourcePWD.Size = New System.Drawing.Size(261, 26)
        Me.txtSourcePWD.TabIndex = 116
        '
        'txtSourceUID
        '
        Me.txtSourceUID.Location = New System.Drawing.Point(540, 60)
        Me.txtSourceUID.Name = "txtSourceUID"
        Me.txtSourceUID.Size = New System.Drawing.Size(310, 26)
        Me.txtSourceUID.TabIndex = 115
        '
        'lblSourcePWD
        '
        Me.lblSourcePWD.AutoSize = True
        Me.lblSourcePWD.Location = New System.Drawing.Point(866, 30)
        Me.lblSourcePWD.Name = "lblSourcePWD"
        Me.lblSourcePWD.Size = New System.Drawing.Size(148, 20)
        Me.lblSourcePWD.TabIndex = 114
        Me.lblSourcePWD.Text = "Password (optional)"
        '
        'lblSourceUID
        '
        Me.lblSourceUID.AutoSize = True
        Me.lblSourceUID.Location = New System.Drawing.Point(536, 30)
        Me.lblSourceUID.Name = "lblSourceUID"
        Me.lblSourceUID.Size = New System.Drawing.Size(153, 20)
        Me.lblSourceUID.TabIndex = 113
        Me.lblSourceUID.Text = "Username (optional)"
        '
        'rdbSourceConnectionString
        '
        Me.rdbSourceConnectionString.AutoSize = True
        Me.rdbSourceConnectionString.Location = New System.Drawing.Point(24, 109)
        Me.rdbSourceConnectionString.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.rdbSourceConnectionString.Name = "rdbSourceConnectionString"
        Me.rdbSourceConnectionString.Size = New System.Drawing.Size(204, 24)
        Me.rdbSourceConnectionString.TabIndex = 105
        Me.rdbSourceConnectionString.Text = "Enter a connection string"
        Me.rdbSourceConnectionString.UseVisualStyleBackColor = True
        '
        'rdbSourceDSN
        '
        Me.rdbSourceDSN.AutoSize = True
        Me.rdbSourceDSN.Checked = True
        Me.rdbSourceDSN.Location = New System.Drawing.Point(24, 26)
        Me.rdbSourceDSN.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.rdbSourceDSN.Name = "rdbSourceDSN"
        Me.rdbSourceDSN.Size = New System.Drawing.Size(61, 24)
        Me.rdbSourceDSN.TabIndex = 104
        Me.rdbSourceDSN.TabStop = True
        Me.rdbSourceDSN.Text = "DSN"
        Me.rdbSourceDSN.UseVisualStyleBackColor = True
        '
        'txtSourceConnectionString
        '
        Me.txtSourceConnectionString.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSourceConnectionString.Enabled = False
        Me.txtSourceConnectionString.Location = New System.Drawing.Point(24, 141)
        Me.txtSourceConnectionString.Name = "txtSourceConnectionString"
        Me.txtSourceConnectionString.Size = New System.Drawing.Size(1118, 26)
        Me.txtSourceConnectionString.TabIndex = 103
        '
        'cmbSourceDSN
        '
        Me.cmbSourceDSN.FormattingEnabled = True
        Me.cmbSourceDSN.Location = New System.Drawing.Point(24, 60)
        Me.cmbSourceDSN.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmbSourceDSN.Name = "cmbSourceDSN"
        Me.cmbSourceDSN.Size = New System.Drawing.Size(500, 28)
        Me.cmbSourceDSN.TabIndex = 98
        '
        'TabPageDestination
        '
        Me.TabPageDestination.Controls.Add(Me.GroupBoxDestination)
        Me.TabPageDestination.Controls.Add(Me.btnDestination)
        Me.TabPageDestination.Controls.Add(Me.lblDestinationTable)
        Me.TabPageDestination.Location = New System.Drawing.Point(4, 29)
        Me.TabPageDestination.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabPageDestination.Name = "TabPageDestination"
        Me.TabPageDestination.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabPageDestination.Size = New System.Drawing.Size(1200, 1216)
        Me.TabPageDestination.TabIndex = 2
        Me.TabPageDestination.Text = "Destination"
        Me.TabPageDestination.UseVisualStyleBackColor = True
        '
        'GroupBoxDestination
        '
        Me.GroupBoxDestination.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBoxDestination.Controls.Add(Me.lblDestinationPWD)
        Me.GroupBoxDestination.Controls.Add(Me.lblDestinationUID)
        Me.GroupBoxDestination.Controls.Add(Me.txtDestinationPWD)
        Me.GroupBoxDestination.Controls.Add(Me.txtDestinationUID)
        Me.GroupBoxDestination.Controls.Add(Me.rdbDestinationDSN)
        Me.GroupBoxDestination.Controls.Add(Me.rdbDestinationConnectionString)
        Me.GroupBoxDestination.Controls.Add(Me.txtDestinationConnectionString)
        Me.GroupBoxDestination.Controls.Add(Me.cmbDestinationDSN)
        Me.GroupBoxDestination.Location = New System.Drawing.Point(28, 23)
        Me.GroupBoxDestination.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBoxDestination.Name = "GroupBoxDestination"
        Me.GroupBoxDestination.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBoxDestination.Size = New System.Drawing.Size(1150, 195)
        Me.GroupBoxDestination.TabIndex = 97
        Me.GroupBoxDestination.TabStop = False
        Me.GroupBoxDestination.Text = "Choose Destination"
        '
        'lblDestinationPWD
        '
        Me.lblDestinationPWD.AutoSize = True
        Me.lblDestinationPWD.Location = New System.Drawing.Point(866, 30)
        Me.lblDestinationPWD.Name = "lblDestinationPWD"
        Me.lblDestinationPWD.Size = New System.Drawing.Size(148, 20)
        Me.lblDestinationPWD.TabIndex = 112
        Me.lblDestinationPWD.Text = "Password (optional)"
        '
        'lblDestinationUID
        '
        Me.lblDestinationUID.AutoSize = True
        Me.lblDestinationUID.Location = New System.Drawing.Point(536, 30)
        Me.lblDestinationUID.Name = "lblDestinationUID"
        Me.lblDestinationUID.Size = New System.Drawing.Size(153, 20)
        Me.lblDestinationUID.TabIndex = 111
        Me.lblDestinationUID.Text = "Username (optional)"
        '
        'txtDestinationPWD
        '
        Me.txtDestinationPWD.Location = New System.Drawing.Point(870, 60)
        Me.txtDestinationPWD.Name = "txtDestinationPWD"
        Me.txtDestinationPWD.Size = New System.Drawing.Size(261, 26)
        Me.txtDestinationPWD.TabIndex = 110
        '
        'txtDestinationUID
        '
        Me.txtDestinationUID.Location = New System.Drawing.Point(540, 60)
        Me.txtDestinationUID.Name = "txtDestinationUID"
        Me.txtDestinationUID.Size = New System.Drawing.Size(310, 26)
        Me.txtDestinationUID.TabIndex = 109
        '
        'rdbDestinationDSN
        '
        Me.rdbDestinationDSN.AutoSize = True
        Me.rdbDestinationDSN.Checked = True
        Me.rdbDestinationDSN.Location = New System.Drawing.Point(24, 26)
        Me.rdbDestinationDSN.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.rdbDestinationDSN.Name = "rdbDestinationDSN"
        Me.rdbDestinationDSN.Size = New System.Drawing.Size(61, 24)
        Me.rdbDestinationDSN.TabIndex = 108
        Me.rdbDestinationDSN.TabStop = True
        Me.rdbDestinationDSN.Text = "DSN"
        Me.rdbDestinationDSN.UseVisualStyleBackColor = True
        '
        'rdbDestinationConnectionString
        '
        Me.rdbDestinationConnectionString.AutoSize = True
        Me.rdbDestinationConnectionString.Location = New System.Drawing.Point(24, 109)
        Me.rdbDestinationConnectionString.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.rdbDestinationConnectionString.Name = "rdbDestinationConnectionString"
        Me.rdbDestinationConnectionString.Size = New System.Drawing.Size(204, 24)
        Me.rdbDestinationConnectionString.TabIndex = 107
        Me.rdbDestinationConnectionString.Text = "Enter a connection string"
        Me.rdbDestinationConnectionString.UseVisualStyleBackColor = True
        '
        'txtDestinationConnectionString
        '
        Me.txtDestinationConnectionString.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDestinationConnectionString.Enabled = False
        Me.txtDestinationConnectionString.Location = New System.Drawing.Point(24, 141)
        Me.txtDestinationConnectionString.Name = "txtDestinationConnectionString"
        Me.txtDestinationConnectionString.Size = New System.Drawing.Size(1119, 26)
        Me.txtDestinationConnectionString.TabIndex = 106
        '
        'cmbDestinationDSN
        '
        Me.cmbDestinationDSN.FormattingEnabled = True
        Me.cmbDestinationDSN.Location = New System.Drawing.Point(24, 60)
        Me.cmbDestinationDSN.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmbDestinationDSN.Name = "cmbDestinationDSN"
        Me.cmbDestinationDSN.Size = New System.Drawing.Size(500, 28)
        Me.cmbDestinationDSN.TabIndex = 91
        '
        'lblDestinationTable
        '
        Me.lblDestinationTable.AutoSize = True
        Me.lblDestinationTable.Location = New System.Drawing.Point(287, 247)
        Me.lblDestinationTable.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblDestinationTable.Name = "lblDestinationTable"
        Me.lblDestinationTable.Size = New System.Drawing.Size(0, 20)
        Me.lblDestinationTable.TabIndex = 35
        '
        'TabPageMapping
        '
        Me.TabPageMapping.Controls.Add(Me.lblProgress)
        Me.TabPageMapping.Controls.Add(Me.lblJobFile)
        Me.TabPageMapping.Controls.Add(Me.dgMapping)
        Me.TabPageMapping.Controls.Add(Me.ckbLogSQL)
        Me.TabPageMapping.Controls.Add(Me.btnImport)
        Me.TabPageMapping.Controls.Add(Me.btnLoad)
        Me.TabPageMapping.Controls.Add(Me.btnListFields)
        Me.TabPageMapping.Controls.Add(Me.btnSave)
        Me.TabPageMapping.Location = New System.Drawing.Point(4, 29)
        Me.TabPageMapping.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabPageMapping.Name = "TabPageMapping"
        Me.TabPageMapping.Size = New System.Drawing.Size(1200, 1216)
        Me.TabPageMapping.TabIndex = 3
        Me.TabPageMapping.Text = "Column Mapping"
        Me.TabPageMapping.UseVisualStyleBackColor = True
        '
        'lblProgress
        '
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(818, 36)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(0, 20)
        Me.lblProgress.TabIndex = 93
        '
        'frmETL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1250, 1295)
        Me.Controls.Add(Me.TabControl)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "frmETL"
        Me.Text = "QuNect ETL"
        CType(Me.dgMapping, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl.ResumeLayout(False)
        Me.TabPageSource.ResumeLayout(False)
        Me.GroupBoxSQL.ResumeLayout(False)
        Me.GroupBoxSource.ResumeLayout(False)
        Me.GroupBoxSource.PerformLayout()
        Me.TabPageDestination.ResumeLayout(False)
        Me.TabPageDestination.PerformLayout()
        Me.GroupBoxDestination.ResumeLayout(False)
        Me.GroupBoxDestination.PerformLayout()
        Me.TabPageMapping.ResumeLayout(False)
        Me.TabPageMapping.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgMapping As System.Windows.Forms.DataGridView
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents btnDestination As System.Windows.Forms.Button
    Friend WithEvents btnListFields As System.Windows.Forms.Button
    Friend WithEvents Source As DataGridViewTextBoxColumn
    Friend WithEvents Destination As DataGridViewComboBoxColumn
    Friend WithEvents btnSave As Button
    Friend WithEvents saveDialog As SaveFileDialog
    Friend WithEvents btnLoad As Button
    Friend WithEvents openFile As OpenFileDialog
    Friend WithEvents lblJobFile As Label
    Friend WithEvents ckbLogSQL As CheckBox
    Friend WithEvents TabControl As TabControl
    Friend WithEvents TabPageSource As TabPage
    Friend WithEvents TabPageDestination As TabPage
    Friend WithEvents TabPageMapping As TabPage
    Friend WithEvents btnSourceTable As Button
    Friend WithEvents txtSQL As RichTextBox
    Friend WithEvents lblDestinationTable As Label
    Friend WithEvents cmbDestinationDSN As ComboBox
    Friend WithEvents lblProgress As Label
    Friend WithEvents GroupBoxSource As GroupBox
    Friend WithEvents rdbSourceConnectionString As RadioButton
    Friend WithEvents rdbSourceDSN As RadioButton
    Friend WithEvents txtSourceConnectionString As TextBox
    Friend WithEvents cmbSourceDSN As ComboBox
    Friend WithEvents GroupBoxSQL As GroupBox
    Friend WithEvents GroupBoxDestination As GroupBox
    Friend WithEvents rdbDestinationConnectionString As RadioButton
    Friend WithEvents txtDestinationConnectionString As TextBox
    Friend WithEvents rdbDestinationDSN As RadioButton
    Friend WithEvents txtSourcePWD As TextBox
    Friend WithEvents txtSourceUID As TextBox
    Friend WithEvents lblSourcePWD As Label
    Friend WithEvents lblSourceUID As Label
    Friend WithEvents lblDestinationPWD As Label
    Friend WithEvents lblDestinationUID As Label
    Friend WithEvents txtDestinationPWD As TextBox
    Friend WithEvents txtDestinationUID As TextBox
End Class
