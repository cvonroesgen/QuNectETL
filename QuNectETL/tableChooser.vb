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
            If frmETL.TabControl.SelectedTab.Name = "TabPageSource" Then
                Dim sourceColumns As DataTable = frmETL.getColumnsDataTable("SELECT * FROM """ & tvAppsTables.SelectedNode.Text() & """", frmETL.txtSourceConnectionString.Text)
                If sourceColumns Is Nothing Then
                    Return
                End If
                frmETL.txtSQL.Text = "SELECT "
                Dim comma As String = ""
                For Each columnRow As DataRow In sourceColumns.Rows
                    frmETL.txtSQL.Text &= comma & """" & columnRow(0) & """"
                    comma = ","
                Next
                frmETL.txtSQL.Text &= " FROM """ & tvAppsTables.SelectedNode.Text() & """"
            Else
                frmETL.lblDestinationTable.Text = tvAppsTables.SelectedNode.Text()
            End If
        Else
            If frmETL.TabControl.SelectedTab.Name <> "TabPageSource" Then
                frmETL.lblDestinationTable.Text = ""
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