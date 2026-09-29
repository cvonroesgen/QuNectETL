Imports System.Data
Imports System.Data.Odbc
Imports System.Data.SqlClient

Module ConnectionHelpers
    Public Const SchemaColumnName As String = "ColumnName"
    Public Const SchemaDataType As String = "DataType"

    Public Const TableChooserGroupName As String = "GroupName"
    Public Const TableChooserDisplayName As String = "DisplayName"
    Public Const TableChooserFullName As String = "FullName"
    Public Const TableChooserSelectName As String = "SelectName"

    Private Function getConnectionStringValue(builder As OdbcConnectionStringBuilder, key As String) As String
        If builder.ContainsKey(key) AndAlso builder(key) IsNot Nothing Then
            Return builder(key).ToString()
        End If
        Return ""
    End Function

    Private Function getDataRowValue(row As DataRow, ParamArray columnNames() As String) As String
        For Each columnName As String In columnNames
            If row.Table.Columns.Contains(columnName) AndAlso Not IsDBNull(row(columnName)) Then
                Return row(columnName).ToString()
            End If
        Next
        Return ""
    End Function

    Private Function isTruthyConnectionSetting(value As String) As Boolean
        If value Is Nothing Then
            Return False
        End If

        Select Case value.Trim().ToLower()
            Case "true", "yes", "sspi"
                Return True
        End Select

        Return False
    End Function

    Public Function tryBuildSqlServerConnectionString(odbcConnectionString As String, ByRef sqlConnectionString As String) As Boolean
        sqlConnectionString = ""

        If String.IsNullOrWhiteSpace(odbcConnectionString) Then
            Return False
        End If

        Try
            Dim directBuilder As New SqlConnectionStringBuilder(odbcConnectionString)
            If directBuilder.DataSource.Length > 0 Then
                sqlConnectionString = directBuilder.ConnectionString
                Return True
            End If
        Catch
        End Try

        Try
            Dim odbcBuilder As New OdbcConnectionStringBuilder(odbcConnectionString)

            Dim driver As String = getConnectionStringValue(odbcBuilder, "Driver")
            Dim server As String = getConnectionStringValue(odbcBuilder, "Server")
            Dim database As String = getConnectionStringValue(odbcBuilder, "Database")
            Dim uid As String = getConnectionStringValue(odbcBuilder, "UID")
            Dim pwd As String = getConnectionStringValue(odbcBuilder, "PWD")
            Dim trusted As String = getConnectionStringValue(odbcBuilder, "Trusted_Connection")
            Dim integrated As String = getConnectionStringValue(odbcBuilder, "Integrated Security")

            If server.Length = 0 Then
                server = getConnectionStringValue(odbcBuilder, "Address")
            End If
            If server.Length = 0 Then
                server = getConnectionStringValue(odbcBuilder, "Addr")
            End If
            If server.Length = 0 Then
                server = getConnectionStringValue(odbcBuilder, "Network Address")
            End If

            Dim isSqlServer As Boolean =
                driver.IndexOf("sql server", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                driver.IndexOf("odbc driver", StringComparison.OrdinalIgnoreCase) >= 0

            If Not isSqlServer OrElse server.Length = 0 Then
                Using probe As New OdbcConnection(odbcConnectionString)
                    probe.Open()

                    If probe.Driver IsNot Nothing AndAlso probe.Driver.IndexOf("sql server", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        isSqlServer = True
                    End If

                    If server.Length = 0 Then
                        server = probe.DataSource
                    End If

                    If database.Length = 0 Then
                        database = probe.Database
                    End If
                End Using
            End If

            If Not isSqlServer OrElse server.Length = 0 Then
                Return False
            End If

            Dim sqlBuilder As New SqlConnectionStringBuilder()
            sqlBuilder.DataSource = server

            If database.Length > 0 Then
                sqlBuilder.InitialCatalog = database
            End If

            If uid.Length > 0 Then
                sqlBuilder.UserID = uid
                sqlBuilder.Password = pwd
            ElseIf isTruthyConnectionSetting(trusted) OrElse isTruthyConnectionSetting(integrated) Then
                sqlBuilder.IntegratedSecurity = True
            Else
                sqlBuilder.IntegratedSecurity = True
            End If

            sqlConnectionString = sqlBuilder.ConnectionString
            Return True
        Catch
            Return False
        End Try
    End Function

    Public Function isSqlServerConnection(connectionString As String) As Boolean
        Dim sqlConnectionString As String = ""
        Return tryBuildSqlServerConnectionString(connectionString, sqlConnectionString)
    End Function

    Public Function quoteIdentifier(connectionString As String, identifier As String) As String
        If isSqlServerConnection(connectionString) Then
            Return "[" & identifier.Replace("]", "]]") & "]"
        End If

        Dim quote As String = ChrW(34)
        Return quote & identifier.Replace(quote, quote & quote) & quote
    End Function

    Public Function buildQualifiedTableName(connectionString As String, schemaName As String, tableName As String) As String
        If isSqlServerConnection(connectionString) AndAlso schemaName.Length > 0 Then
            Return quoteIdentifier(connectionString, schemaName) & "." & quoteIdentifier(connectionString, tableName)
        End If

        Return quoteIdentifier(connectionString, tableName)
    End Function

    Private Function createTableChooserTable() As DataTable
        Dim tables As New DataTable()
        tables.Columns.Add(TableChooserGroupName, GetType(String))
        tables.Columns.Add(TableChooserDisplayName, GetType(String))
        tables.Columns.Add(TableChooserFullName, GetType(String))
        tables.Columns.Add(TableChooserSelectName, GetType(String))
        Return tables
    End Function

    Public Function getTablesForChooser(connectionString As String) As DataTable
        If isSqlServerConnection(connectionString) Then
            Return getSqlServerTablesForChooser(connectionString)
        End If

        Return getOdbcTablesForChooser(connectionString)
    End Function

    Private Function getSqlServerTablesForChooser(connectionString As String) As DataTable
        Dim tables As DataTable = createTableChooserTable()
        Dim sqlConnectionString As String = ""

        If Not tryBuildSqlServerConnectionString(connectionString, sqlConnectionString) Then
            Return tables
        End If

        Using srcConnection As New SqlConnection(sqlConnectionString)
            srcConnection.Open()

            Using srcCmd As New SqlCommand("
SELECT
    TABLE_SCHEMA,
    TABLE_NAME
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_TYPE IN ('BASE TABLE', 'VIEW')
ORDER BY TABLE_SCHEMA, TABLE_NAME", srcConnection)
                Using dr As SqlDataReader = srcCmd.ExecuteReader()
                    While dr.Read()
                        Dim schemaName As String = dr.GetString(0)
                        Dim tableName As String = dr.GetString(1)

                        Dim row As DataRow = tables.NewRow()
                        row(TableChooserGroupName) = schemaName
                        row(TableChooserDisplayName) = tableName
                        row(TableChooserFullName) = schemaName & "." & tableName
                        row(TableChooserSelectName) = buildQualifiedTableName(connectionString, schemaName, tableName)
                        tables.Rows.Add(row)
                    End While
                End Using
            End Using
        End Using

        Return tables
    End Function

    Private Function getOdbcTablesForChooser(connectionString As String) As DataTable
        Dim tables As DataTable = createTableChooserTable()

        Using srcConnection As New OdbcConnection(connectionString)
            srcConnection.Open()

            Dim schemaTables As DataTable = srcConnection.GetSchema("Tables")
            Try
                Dim schemaViews As DataTable = srcConnection.GetSchema("Views")
                schemaTables.Merge(schemaViews)
            Catch
            End Try

            For Each schemaRow As DataRow In schemaTables.Rows
                Dim tableType As String = getDataRowValue(schemaRow, "TABLE_TYPE")
                If tableType.Length > 0 AndAlso
                   tableType <> "TABLE" AndAlso
                   tableType <> "VIEW" AndAlso
                   tableType <> "BASE TABLE" Then
                    Continue For
                End If

                Dim groupName As String = getDataRowValue(schemaRow, "TABLE_SCHEM", "TABLE_CAT")
                Dim tableName As String = getDataRowValue(schemaRow, "TABLE_NAME")

                If tableName.Length = 0 AndAlso schemaRow.Table.Columns.Count > 2 AndAlso Not IsDBNull(schemaRow(2)) Then
                    tableName = schemaRow(2).ToString()
                End If

                If groupName.Length = 0 AndAlso schemaRow.Table.Columns.Count > 0 AndAlso Not IsDBNull(schemaRow(0)) Then
                    groupName = schemaRow(0).ToString()
                End If

                If tableName.Length = 0 Then
                    Continue For
                End If

                Dim row As DataRow = tables.NewRow()
                row(TableChooserGroupName) = groupName
                row(TableChooserDisplayName) = tableName
                row(TableChooserFullName) = tableName
                row(TableChooserSelectName) = buildQualifiedTableName(connectionString, "", tableName)
                tables.Rows.Add(row)
            Next
        End Using

        Return tables
    End Function

    Public Function executeSourceReader(Of TResult)(connectionString As String, sql As String, behavior As CommandBehavior, readerFunc As Func(Of IDataReader, TResult)) As TResult
        Dim sqlConnectionString As String = ""

        If tryBuildSqlServerConnectionString(connectionString, sqlConnectionString) Then
            Using srcConnection As New SqlConnection(sqlConnectionString)
                srcConnection.Open()
                Using srcCmd As New SqlCommand(sql, srcConnection)
                    Using dr As SqlDataReader = srcCmd.ExecuteReader(behavior)
                        Return readerFunc(dr)
                    End Using
                End Using
            End Using
        End If

        Using srcConnection As New OdbcConnection(connectionString)
            srcConnection.Open()
            Using srcCmd As New OdbcCommand(sql, srcConnection)
                Using dr As OdbcDataReader = srcCmd.ExecuteReader(behavior)
                    Return readerFunc(dr)
                End Using
            End Using
        End Using
    End Function

    Public Function loadDataTableFromReader(dr As IDataReader, maxRows As Integer) As DataTable
        Dim table As New DataTable()

        For i As Integer = 0 To dr.FieldCount - 1
            Dim col As New DataColumn()
            col.DataType = dr.GetFieldType(i)
            col.ColumnName = dr.GetName(i)
            col.Caption = dr.GetName(i)
            col.ReadOnly = True
            table.Columns.Add(col)
        Next

        Dim recordCounter As Integer = 0
        While dr.Read()
            Dim r As DataRow = table.NewRow()
            For i As Integer = 0 To dr.FieldCount - 1
                r(i) = dr.GetValue(i)
            Next
            table.Rows.Add(r)

            recordCounter += 1
            If recordCounter >= maxRows Then
                Exit While
            End If
        End While

        Return table
    End Function
End Module