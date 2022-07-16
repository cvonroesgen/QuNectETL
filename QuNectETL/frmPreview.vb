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
            Dim tblPreview As New DataTable()
            'we have to open up a reader on the source
            Dim srcConnection As OdbcConnection
            srcConnection = New OdbcConnection(mConnectionString)
            srcConnection.Open()
            Dim fileLineCounter As Integer = 0
            Dim conversionErrors As String = ""
            Using srcCmd As OdbcCommand = New OdbcCommand(mSQL, srcConnection)
                Dim dr As OdbcDataReader
                Try
                    dr = srcCmd.ExecuteReader()
                    For i As Integer = 0 To dr.FieldCount - 1
                        Dim col As New DataColumn()
                        'Specify the column data type  
                        col.DataType = dr.GetFieldType(i) 'System.Type.GetType("System.Int32")
                        'This name is to identify the column
                        col.ColumnName = dr.GetName(i)
                        'This value will be displayed in the header of the column
                        col.Caption = dr.GetName(i)
                        'This will make this column read only,Since it is autoincrement
                        col.ReadOnly = True
                        tblPreview.Columns.Add(col)
                    Next
                    Dim recordCounter As Integer = 0
                    While (dr.Read())
                        'Add a new row
                        Dim r As DataRow
                        r = tblPreview.NewRow()
                        For i As Integer = 0 To dr.FieldCount - 1
                            r(i) = dr.GetValue(i)
                        Next
                        'Add the row to the tabledata
                        tblPreview.Rows.Add(r)
                        recordCounter += 1
                        If recordCounter >= frmETL.nudPreview.Value Then
                            Exit While
                        End If
                    End While

                Catch excpt As Exception
                    srcCmd.Cancel()
                    srcCmd.Dispose()
                    srcConnection.Close()
                    Me.Cursor = Cursors.Default
                    Throw New System.Exception("Could not get preview records from " & mSQL & vbCrLf & excpt.Message)
                End Try
            End Using
            dgvPreview.DataSource = tblPreview
        Catch ex As Exception
            MsgBox("Could not preview because " & ex.Message)
            Me.Cursor = Cursors.Default
        Finally

        End Try

    End Sub

End Class