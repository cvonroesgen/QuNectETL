Public Class frmTableChooser

    Private Sub tvAppsTables_DoubleClick(sender As Object, e As EventArgs) Handles tvAppsTables.DoubleClick
        btnDone_Click(sender, e)
    End Sub


    Private Sub btnDone_Click(sender As Object, e As EventArgs) Handles btnDone.Click
        If tvAppsTables.SelectedNode Is Nothing Then
            Me.Hide()
            Exit Sub
        End If

        If tvAppsTables.SelectedNode.Level = 1 Then
            Dim selectedTableFullName As String = tvAppsTables.SelectedNode.Name
            If selectedTableFullName.Length = 0 Then
                selectedTableFullName = tvAppsTables.SelectedNode.Text()
            End If

            Dim selectedTableSqlName As String = ""
            If tvAppsTables.SelectedNode.Tag IsNot Nothing Then
                selectedTableSqlName = tvAppsTables.SelectedNode.Tag.ToString()
            End If
            If selectedTableSqlName.Length = 0 Then
                selectedTableSqlName = buildQualifiedTableName(frmETL.txtSourceConnectionString.Text, "", tvAppsTables.SelectedNode.Text())
            End If

            If frmETL.TabControl.SelectedTab.Name = "TabPageSource" Then
                Dim sourceColumns As DataTable = frmETL.getColumnsDataTable("SELECT * FROM " & selectedTableSqlName, frmETL.txtSourceConnectionString.Text)
                If sourceColumns Is Nothing Then
                    Return
                End If

                frmETL.txtSQL.Text = "SELECT "
                Dim comma As String = ""
                For Each columnRow As DataRow In sourceColumns.Rows
                    frmETL.txtSQL.Text &= comma & quoteIdentifier(frmETL.txtSourceConnectionString.Text, columnRow(SchemaColumnName).ToString())
                    comma = ","
                Next
                frmETL.txtSQL.Text &= " FROM " & selectedTableSqlName
                frmETL.lblSourceTable.Text = selectedTableFullName
            Else
                frmETL.lblDestinationTable.Text = selectedTableFullName
            End If
        Else
            If frmETL.TabControl.SelectedTab.Name <> "TabPageSource" Then
                frmETL.lblDestinationTable.Text = ""
                frmETL.lblSourceTable.Text = ""
            End If
        End If

        hideButtons()
        Me.Hide()
    End Sub
    Private Sub hideButtons()
        frmETL.btnImport.Visible = False
        frmETL.dgMapping.Visible = False
    End Sub
End Class