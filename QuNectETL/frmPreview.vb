Imports System.Data
Imports System.Data.Odbc

Public Class frmPreview
    Private mSQL As String
    Public Property sql() As String
        Get
            Return mSQL
        End Get
        Set(ByVal value As String)
            mSQL = value
        End Set
    End Property

    Private mConnectionString As String
    Public Property connectionString() As String
        Get
            Return mConnectionString
        End Get
        Set(ByVal value As String)
            mConnectionString = value
        End Set
    End Property

    Private Sub frmPreview_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            dgvPreview.DataSource =
                executeSourceReader(mConnectionString, mSQL, CommandBehavior.Default,
                    Function(dr As IDataReader) loadDataTableFromReader(dr, CInt(frmETL.nudPreview.Value)))
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MsgBox("Could not preview because " & ex.Message)
            Me.Close()
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

End Class