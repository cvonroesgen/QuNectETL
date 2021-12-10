Public Class frmSQL
    Private Sub txtSQL_TextChanged(sender As Object, e As EventArgs) Handles txtSQL.TextChanged
        frmCopy.strSourceSQL = txtSQL.Text
    End Sub
End Class