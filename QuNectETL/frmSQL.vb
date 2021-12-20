Public Class frmSQL
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        txtSQL.Text = frmETL.strSourceSQL
    End Sub

    Private Sub txtSQL_TextChanged(sender As Object, e As EventArgs)
        frmETL.displaySQL(txtSQL.Text)
    End Sub

    Private Sub btnSQLDone_Click(sender As Object, e As EventArgs) Handles btnSQLDone.Click
        frmETL.displaySQL(txtSQL.Text)
        frmETL.listFields(frmETL.lblDestinationTable.Text, txtSQL.Text)
        Me.Close()
    End Sub
End Class