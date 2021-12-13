Public Class frmSQL
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        txtSQL.Text = frmETL.strSourceSQL
    End Sub

    Private Sub txtSQL_TextChanged(sender As Object, e As EventArgs) Handles txtSQL.TextChanged
        frmETL.strSourceSQL = txtSQL.Text
        frmETL.displaySQL()
    End Sub

    Private Sub btnSQLDone_Click(sender As Object, e As EventArgs) Handles btnSQLDone.Click
        frmETL.strSourceSQL = txtSQL.Text
        frmETL.displaySQL()
        Me.Close()
    End Sub
End Class