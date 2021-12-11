﻿Imports System.ComponentModel
Imports System.Data.Odbc
Imports System.Text.RegularExpressions

Public Class frmETL
    Private Structure connectionStrings
        Public src As String
        Public dst As String
    End Structure

    Public Structure config
        Public uid As String
        Public pwd As String
        Public server As String
        Public apptoken As String
        Public pwdIsPassword As Boolean
        Public detectProxy As Boolean
        Public DSN As String
        Public dbid As String
        Public sourceFieldOrdinals As String
        Public fidsForImport As String
        Public srcSQL As String
    End Structure
    Private Class qdbField
        Public Sub New(_fid As String, _label As String, _type As String, _parentFieldID As String, _unique As Boolean, _required As Boolean, _base_type As String, _decimal_places As Integer)
            fid = _fid
            label = _label
            parentFieldID = _parentFieldID
            unique = _unique
            required = _required
            base_type = _base_type
            decimal_places = _decimal_places
        End Sub
        Public fid As String
        Public label As String
        Public type As String
        Public parentFieldID As String
        Public unique As Boolean
        Public required As Boolean
        Public base_type As String
        Public decimal_places As Integer
        Overrides Function toString() As String
            Return fid
        End Function
    End Class

    Private Class srcField
        Public Sub New(_label As String, _type As String)
            label = _label
            type = _type
        End Sub
        Public label As String
        Public type As String
        Overrides Function toString() As String
            Return label
        End Function
    End Class
    Private cmdLineArgs() As String
    Private automode As Boolean = True
    Private Const AppName = "QuNectETL"
    Public Shared strSourceSQL As String
    Private Title = "QuNect ETL"
    Private destinationFieldNodes As New Dictionary(Of String, qdbField)
    Private qdbConnections As New Dictionary(Of String, OdbcConnection)
    Private rids As HashSet(Of String)
    Private keyfid As String
    Private uniqueExistingFieldValues As Dictionary(Of String, HashSet(Of String))
    Private uniqueNewFieldValues As Dictionary(Of String, HashSet(Of String))
    Private sourceLabelToFieldType As Dictionary(Of String, String)
    Private sourceFieldNames As New Dictionary(Of String, Integer)
    Private destinationLabelsToFids As Dictionary(Of String, String)
    Private sourceLabelsToFids As Dictionary(Of String, String)
    Private isBooleanTrue As Regex = New Regex("y|tr|c|[1-9]", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Private copyFinished As Boolean = False
    Dim copyCount As Integer
    Dim existingCount As Integer
    Private Class qdbVersion
        Public year As Integer
        Public major As Integer
        Public minor As Integer
    End Class
    Private qdbVer As qdbVersion = New qdbVersion
    Enum mapping
        source
        destination
    End Enum


    Private Sub restore_Load(sender As Object, e As EventArgs) Handles Me.Load
        cmdLineArgs = System.Environment.GetCommandLineArgs()
        If cmdLineArgs.Length > 1 Then
            automode = True
            Dim cnfg As New config
            loadConfig(cmdLineArgs(1), cnfg)
            Dim cnctStrings As connectionStrings
            cnctStrings.src = "DSN=" & cnfg.DSN & ";"
            cnctStrings.dst = getQDBConnectionString(True, cnfg.uid, cnfg.pwd, cnfg.server, cnfg.apptoken, cnfg.pwdIsPassword)
            listDestinationFields(cnfg.dbid)
            executeUpload(cnctStrings, New ArrayList(cnfg.fidsForImport.Split(".")), New ArrayList(cnfg.sourceFieldOrdinals.Split(".")), cnfg.srcSQL)
            Me.Close()
            Exit Sub
        Else
            automode = False
        End If
        Dim myBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
        Title &= " " & myBuildInfo.ProductVersion
        Text = Title
        txtUsername.Text = GetSetting(AppName, "Credentials", "username")
        cmbPassword.SelectedIndex = CInt(GetSetting(AppName, "Credentials", "passwordOrToken", "0"))
        txtPassword.Text = GetSetting(AppName, "Credentials", "password")
        txtServer.Text = GetSetting(AppName, "Credentials", "server", "")
        txtAppToken.Text = GetSetting(AppName, "Credentials", "apptoken", "b2fr52jcykx3tnbwj8s74b8ed55b")
        lblDestinationTable.Text = GetSetting(AppName, "config", "destinationtable", "")
        strSourceSQL = GetSetting(AppName, "config", "SQL", "")
        Dim dsn As String = GetSetting(AppName, "Connection", "DSN", "")
        GetDSNs()
        cmbDSN.SelectedIndex = cmbDSN.FindStringExact(dsn)

        Dim detectProxySetting As String = GetSetting(AppName, "Credentials", "detectproxysettings", "0")
        If detectProxySetting = "1" Then
            ckbDetectProxy.Checked = True
        Else
            ckbDetectProxy.Checked = False
        End If
        lblSQL.Text = strSourceSQL
        showHideControls()
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        saveDialog.Filter = "JOB Files (*.job)|*.job"
        If saveDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim strJob As String = txtUsername.Text
            strJob &= vbCrLf & txtPassword.Text
            strJob &= vbCrLf & cmbPassword.SelectedIndex
            strJob &= vbCrLf & txtServer.Text
            strJob &= vbCrLf & txtAppToken.Text
            strJob &= vbCrLf & lblDestinationTable.Text
            If ckbDetectProxy.Checked Then
                strJob &= vbCrLf & "1"
            Else
                strJob &= vbCrLf & "0"
            End If
            strJob &= vbCrLf & cmbDSN.Text
            Dim sourceFieldOrdinals As String = ""
            Dim fidsForImport As String = ""
            For i As Integer = 0 To dgMapping.Rows.Count - 1
                Dim destComboBoxCell As DataGridViewComboBoxCell = dgMapping.Rows(i).Cells(mapping.destination)
                If destComboBoxCell.Value Is Nothing Then Continue For
                Dim destDDIndex = destComboBoxCell.Items.IndexOf(destComboBoxCell.Value)
                If destDDIndex = 0 Then Continue For

                Dim fieldNode As qdbField
                fieldNode = destinationFieldNodes(destinationLabelsToFids(destComboBoxCell.Value))

                If keyfid = "3" And fieldNode.fid = "3" Then
                    Dim copyAnyway As MsgBoxResult = Alert("Copying into the key field " & fieldNode.label & " will update existing records without creating New records. Do you want To Continue?", MsgBoxStyle.YesNo)
                    If copyAnyway = MsgBoxResult.No Then
                        Exit For
                    End If
                End If
                If fidsForImport.Contains(fieldNode.fid) Then
                    Alert("You cannot import two different columns into the same field: " & destComboBoxCell.Value, MsgBoxStyle.OkOnly, AppName)
                    Exit For
                End If
                sourceFieldOrdinals &= "." & i
                fidsForImport &= "." & fieldNode.fid
            Next
            strJob &= vbCrLf & sourceFieldOrdinals.Substring(1)
            strJob &= vbCrLf & fidsForImport.Substring(1)
            strJob &= vbCrLf
            strJob &= vbCrLf & strSourceSQL
            My.Computer.FileSystem.WriteAllText(saveDialog.FileName, strJob, False)
        End If
    End Sub
    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        loadConfig("")
    End Sub
    Sub loadConfig(filename As String)
        Dim cnfg As New config
        loadConfig(filename, cnfg)

        txtUsername.Text = cnfg.uid
        txtPassword.Text = cnfg.pwd
        cmbPassword.SelectedIndex = CInt(cnfg.pwdIsPassword)
        txtServer.Text = cnfg.server
        txtAppToken.Text = cnfg.apptoken
        lblDestinationTable.Text = cnfg.dbid
        ckbDetectProxy.Checked = cnfg.detectProxy
        cmbDSN.SelectedIndex = cmbDSN.FindStringExact(cnfg.DSN)
        Dim sourceFieldOrdinals As String = cnfg.sourceFieldOrdinals
        Dim fidsForImport As String = cnfg.fidsForImport
        strSourceSQL = cnfg.srcSQL
        Dim fidsToLabels As Dictionary(Of String, String) = listFields(lblDestinationTable.Text, strSourceSQL)
        Dim ordinals As String() = sourceFieldOrdinals.Split(".")
        Dim fids As String() = fidsForImport.Split(".")
        For i As Integer = 0 To ordinals.Count - 1
            Dim destComboBoxCell As DataGridViewComboBoxCell = dgMapping.Rows(CInt(ordinals(i))).Cells(mapping.destination)
            destComboBoxCell.Value = fidsToLabels(fids(i))
        Next
    End Sub
    Sub loadConfig(filename As String, ByRef cnfg As config)
        If Not automode AndAlso openFile.ShowDialog = Windows.Forms.DialogResult.OK Then
            filename = openFile.FileName
        End If
        Dim jobFileReader As System.IO.StreamReader
        jobFileReader = My.Computer.FileSystem.OpenTextFileReader(filename)
        cnfg.uid = jobFileReader.ReadLine()
        cnfg.pwd = jobFileReader.ReadLine()
        cnfg.pwdIsPassword = jobFileReader.ReadLine()
        cnfg.server = jobFileReader.ReadLine()
        cnfg.apptoken = jobFileReader.ReadLine()
        cnfg.dbid = jobFileReader.ReadLine()
        cnfg.detectProxy = CBool(jobFileReader.ReadLine())
        cnfg.DSN = jobFileReader.ReadLine()
        cnfg.sourceFieldOrdinals = jobFileReader.ReadLine()
        cnfg.fidsForImport = jobFileReader.ReadLine()
        jobFileReader.ReadLine()
        cnfg.srcSQL = ""
        While Not jobFileReader.EndOfStream
            cnfg.srcSQL &= jobFileReader.ReadLine()
        End While
        jobFileReader.Close()
    End Sub
    Sub showHideControls()
        cmbPassword.Visible = txtUsername.Text.Length > 0
        txtPassword.Visible = cmbPassword.Visible And cmbPassword.SelectedIndex <> 0
        txtServer.Visible = txtPassword.Visible And txtPassword.Text.Length > 0
        lblServer.Visible = txtServer.Visible
        lblAppToken.Visible = cmbPassword.Visible And cmbPassword.SelectedIndex = 1
        btnAppToken.Visible = lblAppToken.Visible
        txtAppToken.Visible = lblAppToken.Visible
        btnUserToken.Visible = cmbPassword.Visible And cmbPassword.SelectedIndex = 2
        btnListFields.Visible = False
        btnDestination.Visible = True
        btnImport.Visible = False
        dgMapping.Visible = False
        btnSave.Visible = False


        If lblDestinationTable.Text <> "" Then
            btnListFields.Visible = True
        Else
            Exit Sub
        End If
        If lblDestinationTable.Text <> "" Then
            btnListFields.Visible = True
        Else
            Exit Sub
        End If
        If dgMapping.RowCount > 0 Then
            dgMapping.Visible = True
            btnImport.Visible = True
            btnSave.Visible = True
        End If
    End Sub
    Private Sub txtServer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtServer.TextChanged
        SaveSetting(AppName, "Credentials", "server", txtServer.Text)
        showHideControls()
    End Sub
    Private Sub ckbDetectProxy_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ckbDetectProxy.CheckStateChanged
        If ckbDetectProxy.Checked Then
            SaveSetting(AppName, "Credentials", "detectproxysettings", "1")
        Else
            SaveSetting(AppName, "Credentials", "detectproxysettings", "0")
        End If
        showHideControls()
    End Sub
    Private Sub txtAppToken_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAppToken.TextChanged
        SaveSetting(AppName, "Credentials", "apptoken", txtAppToken.Text)
        showHideControls()
    End Sub
    Private Sub lblDestinationTable_TextChanged(sender As Object, e As EventArgs) Handles lblDestinationTable.TextChanged
        SaveSetting(AppName, "config", "destinationtable", lblDestinationTable.Text)
        showHideControls()
    End Sub


    Private Sub txtUsername_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUsername.TextChanged
        SaveSetting(AppName, "Credentials", "username", txtUsername.Text)
        showHideControls()

    End Sub

    Private Sub txtPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPassword.TextChanged
        SaveSetting(AppName, "Credentials", "password", txtPassword.Text)

        showHideControls()
    End Sub
    Private Function getConnectionString(usefids As Boolean) As String
        Dim pwdIsPassword As Boolean
        If cmbPassword.SelectedIndex = 0 Then
            cmbPassword.Focus()
            Throw New System.Exception("Please indicate whether you are Using a password Or a user token.")
            Return ""
        ElseIf cmbPassword.SelectedIndex = 1 Then
            pwdIsPassword = True
        Else
            pwdIsPassword = False
        End If
        Return getQDBConnectionString(usefids, txtUsername.Text, txtPassword.Text, txtServer.Text, txtAppToken.Text, pwdIsPassword)
    End Function
    Private Function getQDBConnectionString(usefids As Boolean, txtUsername As String, txtPassword As String, txtServer As String, txtAppToken As String, pwdIsPassword As Boolean) As String

        getQDBConnectionString = "Driver={QuNect ODBC For QuickBase};FIELDNAMECHARACTERS=all;ALLREVISIONS=ALL;uid=" & txtUsername & ";pwd=" & txtPassword & ";QUICKBASESERVER=" & txtServer & ";APPTOKEN=" & txtAppToken
        If usefids Then
            getQDBConnectionString &= ";USEFIDS=1"
        End If
        If pwdIsPassword Then
            getQDBConnectionString &= ";PWDISPASSWORD=1"
        Else
            getQDBConnectionString &= ";PWDISPASSWORD=0"
        End If

    End Function
    Function countRecords(sql As String, connection As OdbcConnection) As Integer
        Using command As OdbcCommand = New OdbcCommand(sql, connection)
            Dim dr As OdbcDataReader
            Try
                dr = command.ExecuteReader()
            Catch excpt As Exception
                Throw New System.Exception("Could Not count records " & excpt.Message)
            End Try
            If Not dr.HasRows Then
                Throw New System.Exception("Counting records failed. " & sql)
            Else
                Return dr(0)
            End If
        End Using
    End Function
    Function setODBCParameter(val As Object, fid As String, command As OdbcCommand, ByRef fileLineCounter As Integer, ByRef conversionErrors As String) As Boolean
        Dim qdbType As OdbcType = command.Parameters("@fid" & fid).OdbcType
        If IsDBNull(val) Then
            command.Parameters("@fid" & fid).Value = val
            Return False
        End If
        Try
            Select Case qdbType
                Case OdbcType.Int
                    If val = "" Then
                        Dim nullInteger As Integer
                        command.Parameters("@fid" & fid).Value = nullInteger
                    Else
                        command.Parameters("@fid" & fid).Value = Convert.ToInt32(val)
                    End If
                Case OdbcType.Double, OdbcType.Numeric
                    command.Parameters("@fid" & fid).Value = Convert.ToDouble(val)
                Case OdbcType.Date
                    command.Parameters("@fid" & fid).Value = val
                Case OdbcType.DateTime
                    command.Parameters("@fid" & fid).Value = Convert.ToDateTime(val)
                Case OdbcType.Time
                    command.Parameters("@fid" & fid).Value = val
                Case OdbcType.Bit
                    Dim match As Match = isBooleanTrue.Match(val)
                    If match.Success Then
                        command.Parameters("@fid" & fid).Value = True
                    Else
                        command.Parameters("@fid" & fid).Value = False
                    End If
                Case Else
                    command.Parameters("@fid" & fid).Value = val
            End Select
        Catch ex As Exception
            Return True
        End Try
        Return False
    End Function
    Private Function getODBCTypeFromQuickBaseFieldNode(fieldNode As qdbField) As OdbcType
        Select Case fieldNode.base_type
            Case "text"
                Return OdbcType.VarChar
            Case "float"
                If fieldNode.decimal_places > 0 Then
                    Return OdbcType.Numeric
                Else
                    Return OdbcType.Double
                End If

            Case "bool"
                Return OdbcType.Bit
            Case "int32"
                Return OdbcType.Int
            Case "int64"
                Select Case fieldNode.type
                    Case "timestamp"
                        Return OdbcType.DateTime
                    Case "timeofday"
                        Return OdbcType.Time
                    Case "duration"
                        Return OdbcType.Double
                    Case Else
                        Return OdbcType.Date
                End Select
            Case Else
                Return OdbcType.VarChar
        End Select
        Return OdbcType.VarChar
    End Function

    Private Function uploadToQuickbase(cnctStrings As connectionStrings) As Boolean
        Dim destinationFields As New ArrayList
        Dim sourceFieldOrdinals As New ArrayList
        Dim fidsForImport As New HashSet(Of String)
        Dim strSQL As String = "INSERT INTO """ & lblDestinationTable.Text & """ (fid"

        For i As Integer = 0 To dgMapping.Rows.Count - 1
            Dim destComboBoxCell As DataGridViewComboBoxCell = DirectCast(dgMapping.Rows(i).Cells(mapping.destination), System.Windows.Forms.DataGridViewComboBoxCell)
            If destComboBoxCell.Value Is Nothing Then Continue For
            Dim destDDIndex = destComboBoxCell.Items.IndexOf(destComboBoxCell.Value)
            If destDDIndex = 0 Then Continue For

            Dim fieldNode As qdbField
            fieldNode = destinationFieldNodes(destinationLabelsToFids(destComboBoxCell.Value))

            If keyfid = "3" And fieldNode.fid = "3" And Not automode Then
                Dim copyAnyway As MsgBoxResult = Alert("Copying into the key field " & fieldNode.label & " will update existing records without creating New records. Do you want To Continue?", MsgBoxStyle.YesNo)
                If copyAnyway = MsgBoxResult.No Then
                    VolatileWrite(copyFinished, True)
                    Return False
                End If
            End If
            If fidsForImport.Contains(fieldNode.fid) And Not automode Then
                Alert("You cannot import two different columns into the same field: " & destComboBoxCell.Value, MsgBoxStyle.OkOnly, AppName)
                VolatileWrite(copyFinished, True)
                Return False
            End If
            fidsForImport.Add(fieldNode.fid)
            destinationFields.Add(fieldNode)
            sourceFieldOrdinals.Add(i)
        Next
        If fidsForImport.Count = 0 And Not automode Then
            Alert("You must map at least one field from the source table to the destination table.", MsgBoxStyle.OkOnly, AppName)
            VolatileWrite(copyFinished, True)
            Return False
        End If
        Return executeUpload(cnctStrings, destinationFields, sourceFieldOrdinals, strSQL)
    End Function
    Private Function executeUpload(cnctStrings As connectionStrings, destinationFields As ArrayList, sourceFieldOrdinals As ArrayList, strSQL As String) As Boolean
        Try

            Dim quNectConn As OdbcConnection = getquNectConn(cnctStrings.dst)
            strSQL &= String.Join(", fid", destinationFields.ToArray) & ") VALUES ("
            For Each var In destinationFields
                strSQL &= "?,"
            Next
            strSQL = strSQL.Substring(0, strSQL.Length - 1) & ")"
            Using command As OdbcCommand = New OdbcCommand(strSQL, quNectConn)
                For Each field In destinationFields
                    Dim qdbType As OdbcType
                    qdbType = getODBCTypeFromQuickBaseFieldNode(field)
                    command.Parameters.Add("@fid" & field.fid, qdbType)
                Next

                Dim transaction As OdbcTransaction = Nothing
                transaction = quNectConn.BeginTransaction()
                command.Transaction = transaction
                command.CommandType = CommandType.Text
                command.CommandTimeout = 0
                Dim progressThread As System.Threading.Thread = New Threading.Thread(AddressOf showProgress)
                progressThread.Start()
                'we have to open up a reader on the source
                Dim srcConnection As OdbcConnection
                srcConnection = New OdbcConnection(cnctStrings.src)

                srcConnection.Open()
                Dim fileLineCounter As Integer = 0
                Dim conversionErrors As String = ""
                Using srcCmd As OdbcCommand = New OdbcCommand(strSourceSQL, srcConnection)
                    Dim dr As OdbcDataReader
                    Try
                        dr = srcCmd.ExecuteReader()
                        Dim rowCount As Integer = 0
                        While (dr.Read())
                            fileLineCounter += 1
                            For i As Integer = 0 To sourceFieldOrdinals.Count - 1
                                If setODBCParameter(dr.GetValue(sourceFieldOrdinals(i)), destinationFields(i).fid, command, fileLineCounter, conversionErrors) Then
                                    Exit While
                                End If
                            Next
                            rowCount = command.ExecuteNonQuery()
                        End While
                        transaction.Commit()
                        If Not automode Then
                            Alert(rowCount & " rows were uploaded to Quickbase.")
                        End If
                    Catch excpt As Exception
                        srcCmd.Cancel()
                        srcCmd.Dispose()
                        transaction.Rollback()
                        srcConnection.Close()
                        Throw New System.Exception("Could not get record " & fileLineCounter & " from " & cnctStrings.src & vbCrLf & excpt.Message)
                    End Try

                End Using
            End Using

        Catch ex As Exception
            Alert("Could Not copy because " & ex.Message)
        End Try
        VolatileWrite(copyFinished, True)
        Return True

    End Function
    Function Alert(msg As String) As MsgBoxResult
        If automode Then
            Throw New System.Exception(msg)
        Else
            Return MsgBox(msg)
        End If
    End Function
    Function Alert(msg As String, style As MsgBoxStyle) As MsgBoxResult
        If automode Then
            Throw New System.Exception(msg)
        Else
            Return MsgBox(msg, style)
        End If
    End Function
    Function Alert(msg As String, style As MsgBoxStyle, title As String) As MsgBoxResult
        If automode Then
            Throw New System.Exception(msg)
        Else
            Return MsgBox(msg, style, title)
        End If
    End Function
    Private Function getquNectConn(connectionString As String) As OdbcConnection
        If qdbConnections.ContainsKey(connectionString) Then
            Return qdbConnections(connectionString)
        End If
        Dim quNectConn As OdbcConnection = New OdbcConnection(connectionString)
        Try
            quNectConn.Open()
        Catch excpt As Exception
            Me.Cursor = Cursors.Default
            If excpt.Message.StartsWith("Error [IM003]") Or excpt.Message.Contains("Data source name Not found") Then
                Alert("Please install QuNect ODBC For QuickBase from http://qunect.com/download/QuNect.exe and try again.")
            Else
                Alert(excpt.Message.Substring(13))
            End If
            Return Nothing
            Exit Function
        End Try

        Dim ver As String = quNectConn.ServerVersion
        Me.Text = Title & " with QuNect ODBC for QuickBase " & ver
        Dim m As Match = Regex.Match(ver, "\d+\.(\d+)\.(\d+)\.(\d+)")
        qdbVer.year = CInt(m.Groups(1).Value)
        qdbVer.major = CInt(m.Groups(2).Value)
        qdbVer.minor = CInt(m.Groups(3).Value)
        If (qdbVer.major < 7) Or (qdbVer.major = 7 And qdbVer.minor < 37) Then
            Alert("You are running the " & ver & " version of QuNect ODBC for QuickBase. Please install the latest version from http://qunect.com/download/QuNect.exe")
            quNectConn.Close()
            Me.Cursor = Cursors.Default
            Return Nothing
            Exit Function
        End If
        qdbConnections.Add(connectionString, quNectConn)
        Return quNectConn
    End Function



    Private Sub btnDestination_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDestination.Click
        listTables()
    End Sub
    Private Sub listTables()
        Me.Cursor = Cursors.WaitCursor
        Try
            Dim connectionString As String = getConnectionString(False)
            Dim quNectConn As OdbcConnection = getquNectConn(connectionString)
            Dim tables As DataTable = quNectConn.GetSchema("Tables")
            Dim views As DataTable = quNectConn.GetSchema("Views")
            tables.Merge(views)
            listTablesFromGetSchema(tables)

        Catch ex As Exception
            Alert("Could not list tables because " & ex.Message)
        End Try
    End Sub
    Sub listCatalogs()
        Try
            Dim connectionString As String = getConnectionString(False)
            Dim quNectConn As OdbcConnection = getquNectConn(connectionString)
            Using quNectCmd = New OdbcCommand("SELECT * FROM CATALOGS", quNectConn)
                Dim dr As OdbcDataReader = quNectCmd.ExecuteReader()
                frmTableChooser.tvAppsTables.BeginUpdate()
                frmTableChooser.tvAppsTables.Nodes.Clear()
                While (dr.Read())
                    Dim applicationName As String = dr.GetString(0)
                    Dim appDBID As String = dr.GetString(4)
                    Dim appNode As TreeNode = frmTableChooser.tvAppsTables.Nodes.Add(applicationName)
                    appNode.Tag = appDBID
                End While
                frmTableChooser.Text = "Choose an Application"
                frmTableChooser.tvAppsTables.EndUpdate()
            End Using
        Catch ex As Exception
            Alert("Could not list catalogs because " & ex.Message)
            Exit Sub
        End Try
        frmTableChooser.Show()
    End Sub
    Sub listTablesFromGetSchema(tables As DataTable)
        frmTableChooser.tvAppsTables.BeginUpdate()
        frmTableChooser.tvAppsTables.Nodes.Clear()
        frmTableChooser.tvAppsTables.ShowNodeToolTips = True
        Dim dbName As String
        Dim applicationName As String = ""
        Dim prevAppName As String = ""
        Dim dbid As String
        pb.Minimum = 0
        pb.Value = 0
        pb.Visible = True
        pb.Maximum = tables.Rows.Count
        Dim getDBIDfromdbName As New Regex("([a-z0-9~]+)$")


        For i = 0 To tables.Rows.Count - 1
            pb.Value = i
            Application.DoEvents()
            dbName = tables.Rows(i)(2)
            applicationName = tables.Rows(i)(0)
            Dim dbidMatch As Match = getDBIDfromdbName.Match(dbName)
            dbid = dbidMatch.Value
            If applicationName <> prevAppName Then

                Dim appNode As TreeNode = frmTableChooser.tvAppsTables.Nodes.Add(applicationName)
                appNode.Tag = dbid
                prevAppName = applicationName
            End If
            Dim tableName As String = dbName
            If applicationName.Length And dbName.Length > applicationName.Length Then
                tableName = dbName.Substring(applicationName.Length + 2)
            End If
            If applicationName.Length Then
                Dim tableNode As TreeNode = frmTableChooser.tvAppsTables.Nodes(frmTableChooser.tvAppsTables.Nodes.Count - 1).Nodes.Add(tableName)
                tableNode.Tag = dbid
            Else
                Dim tableNode As TreeNode = frmTableChooser.tvAppsTables.Nodes.Add(tableName)
                tableNode.Tag = dbid
            End If

        Next
        frmTableChooser.Text = "Choose a Table"
        frmTableChooser.tvAppsTables.EndUpdate()
        pb.Value = 0
        btnImport.Visible = True
        lblDestinationTable.Visible = True
        frmTableChooser.Show()
        Me.Cursor = Cursors.Default
    End Sub
    Function listDestinationFields(destinationDBID As String) As Dictionary(Of String, String)
        destinationLabelsToFids = New Dictionary(Of String, String)
        Dim destinationFIDtoLabel As New Dictionary(Of String, String)
        destinationFieldNodes.Clear()
        listDestinationFields = destinationFIDtoLabel
        Try
            Dim connectionString As String = getConnectionString(False)
            Dim connection As OdbcConnection = getquNectConn(connectionString)
            If connection Is Nothing Then Exit Function
            Dim strSQL As String = "SELECT label, fid, field_type, parentFieldID, ""isunique"", required, ""iskey"", base_type, decimal_places  FROM """ & destinationDBID & "~fields"" WHERE (mode = '' and role = '') or fid = '3'"

            Using quNectCmd As OdbcCommand = New OdbcCommand(strSQL, connection)
                Dim dr As OdbcDataReader
                Try
                    dr = quNectCmd.ExecuteReader()
                Catch excpt As Exception
                    Throw New System.Exception("Could not get field information for " & destinationDBID & vbCrLf & excpt.Message)
                End Try
                If Not dr.HasRows Then
                    Exit Function
                End If

                'Loop through all of the fields in the schema.
                DirectCast(dgMapping.Columns(mapping.destination), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Clear()

                DirectCast(dgMapping.Columns(mapping.destination), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Add("")
                While (dr.Read())
                    Dim field As New qdbField(dr.GetString(1), dr.GetString(0), dr.GetString(2), dr.GetString(3), dr.GetBoolean(4), dr.GetBoolean(5), dr.GetString(7), 0)
                    destinationFIDtoLabel.Add(field.fid, field.label)
                    If field.parentFieldID <> "" Then
                        field.label = destinationFIDtoLabel(field.parentFieldID) & ": " & field.label
                    End If
                    If dr.GetBoolean(6) Then
                        keyfid = field.fid
                    End If
                    If Not IsDBNull(dr(8)) Then
                        field.decimal_places = dr.GetInt32(8)
                    End If
                    destinationFieldNodes.Add(field.fid, field)
                    destinationLabelsToFids.Add(field.label, field.fid)
                    DirectCast(dgMapping.Columns(mapping.destination), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Add(field.label)
                End While
                dr.Close()
            End Using

        Catch ex As Exception
            Alert(ex.Message)
        End Try
    End Function
    Function listFields(destinationDBID As String, strSourceSQL As String) As Dictionary(Of String, String)
        listFields = listDestinationFields(destinationDBID)
        Try
            'here we need to open the source and get the field names
            Dim srcConnection As OdbcConnection
            srcConnection = New OdbcConnection("DSN=" & cmbDSN.Text & ";")
            srcConnection.Open()

            sourceFieldNames.Clear()
            sourceLabelToFieldType = New Dictionary(Of String, String)

            Using srcCmd As OdbcCommand = New OdbcCommand(strSourceSQL, srcConnection)
                Dim dr As OdbcDataReader
                Try
                    dr = srcCmd.ExecuteReader()
                Catch excpt As Exception
                    Throw New System.Exception("Could not get field information from " & cmbDSN.Text & vbCrLf & excpt.Message)
                End Try
                Dim sourceColumns As DataTable
                Try
                    sourceColumns = dr.GetSchemaTable()
                Catch excpt As Exception
                    Throw New System.Exception("Could not get field information for " & cmbDSN.Text & excpt.Message)
                End Try

                dgMapping.Rows.Clear()
                Dim i As Integer = 0
                For Each columnRow As DataRow In sourceColumns.Rows
                    Dim field As New srcField(columnRow(0), columnRow(5).Name)
                    sourceLabelToFieldType.Add(field.label, field.type)
                    dgMapping.Rows.Add(New String() {field.label})
                    sourceFieldNames.Add(field.label, i)
                    i += 1
                Next
                dr.Close()
            End Using

            For Each row In dgMapping.Rows
                guessDestination(row.Cells(mapping.source).Value, row.Index)
            Next
            showHideControls()

            Me.Cursor = Cursors.Default
        Catch ex As Exception
            Alert(ex.Message)
        End Try
    End Function
    Sub guessDestination(sourceFieldName As String, sourceFieldOrdinal As Integer)

        For Each field As KeyValuePair(Of String, qdbField) In destinationFieldNodes
            If (field.Value.label.ToLower = sourceFieldName.ToLower Or Regex.Replace(field.Value.label, "[^a-zA-Z]", "_").ToLower = sourceFieldName.ToLower) AndAlso field.Value.type <> "address" AndAlso field.Value.fid <> "3" Then
                DirectCast(dgMapping.Rows(sourceFieldOrdinal).Cells(mapping.destination), System.Windows.Forms.DataGridViewComboBoxCell).Value = field.Value.label
                Exit For
            End If
        Next

    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        import()
    End Sub
    Sub import()
        VolatileWrite(copyFinished, False)
        pb.Visible = True
        Dim cnctStrings As connectionStrings
        cnctStrings.src = "DSN=" & cmbDSN.Text & ";"
        cnctStrings.dst = getConnectionString(True)
        Dim copyThread As System.Threading.Thread = New Threading.Thread(AddressOf uploadToQuickbase)
        copyThread.Start(cnctStrings)
    End Sub

    Private Sub btnListFields_Click(sender As Object, e As EventArgs) Handles btnListFields.Click
        If lblDestinationTable.Text = "" Then
            Alert("Please choose a table to copy into.", MsgBoxStyle.OkOnly, AppName)
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        listFields(lblDestinationTable.Text, strSourceSQL)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub dgMapping_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgMapping.CellMouseClick
        If e.RowIndex > 0 AndAlso e.ColumnIndex = mapping.source Then
            If dgMapping.Rows(e.RowIndex).Cells(mapping.destination).Value = "" Then
                guessDestination(dgMapping.Rows(e.RowIndex).Cells(mapping.source).Value, e.RowIndex)
            Else
                dgMapping.Rows(e.RowIndex).Cells(mapping.destination).Value = ""
            End If

        End If
    End Sub

    Private Sub showProgress()
        Try
            Dim quNectConn As OdbcConnection = New OdbcConnection(getConnectionString(True))
            quNectConn.Open()
            Dim pbValue As Integer
            While Not VolatileRead(copyFinished)
                Dim currentCount = countRecords("SELECT count(1) FROM """ & lblDestinationTable.Text & """ WHERE fid4 = '" & txtUsername.Text & "'", quNectConn)
                Debug.WriteLine(Now & currentCount.ToString)
                If currentCount > copyCount + existingCount Then
                    pbValue = copyCount + existingCount
                ElseIf currentCount < existingCount Then
                    pbValue = existingCount
                Else
                    pbValue = currentCount
                End If
                SetLabelProgress("Copying record " & (pbValue - existingCount) & " of " & copyCount)
                SetPBProgress(pbValue)
                Threading.Thread.Sleep(2000)
            End While
            SetLabelProgress("")
            SetPBProgress(copyCount + existingCount)
        Catch ex As Exception
            'Alert(ex.Message)
        End Try
    End Sub
    Delegate Sub LabelDelegate(progressMessage As String)
    Private Sub SetLabelProgress(ByVal progressMessage As String)

        ' InvokeRequired required compares the thread ID of the  
        ' calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If Me.lblProgress.InvokeRequired Or Me.pb.InvokeRequired Then
            Dim d As New LabelDelegate(AddressOf SetLabelProgress)
            Me.Invoke(d, New Object() {progressMessage})
        Else
            Me.lblProgress.Text = progressMessage
        End If
    End Sub
    Delegate Sub PBDelegate(numRecords As Integer)
    Private Sub SetPBProgress(ByVal numRecords As Integer)

        ' InvokeRequired required compares the thread ID of the  
        ' calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If Me.pb.InvokeRequired Then
            Dim d As New PBDelegate(AddressOf SetPBProgress)
            Me.Invoke(d, New Object() {numRecords})
        Else
            Me.pb.Maximum = existingCount + copyCount
            Me.pb.Minimum = existingCount
            Me.pb.Value = numRecords
        End If
    End Sub


    Private Sub frmETL_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        SaveSetting(AppName, "config", "SQL", strSourceSQL)
        VolatileWrite(copyFinished, True)
    End Sub

    Function VolatileRead(Of T)(ByRef Address As T) As T
        VolatileRead = Address
        Threading.Thread.MemoryBarrier()
    End Function

    Sub VolatileWrite(Of T)(ByRef Address As T, ByVal Value As T)
        Threading.Thread.MemoryBarrier()
        Address = Value
    End Sub

    Private Sub cmbPassword_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPassword.SelectedIndexChanged
        SaveSetting(AppName, "Credentials", "passwordOrToken", cmbPassword.SelectedIndex)
        showHideControls()
    End Sub
    Private Sub btnAppToken_Click(sender As Object, e As EventArgs) Handles btnAppToken.Click
        Process.Start("https://qunect.com/flash/AppToken.html")
    End Sub

    Private Sub btnUserToken_Click(sender As Object, e As EventArgs) Handles btnUserToken.Click
        Process.Start("https://qunect.com/flash/UserToken.html")
    End Sub

    Private Sub cmbDSN_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDSN.SelectedIndexChanged
        SaveSetting(AppName, "Connection", "DSN", cmbDSN.Text)
        cmbDSN_TextChanged(sender, e)
    End Sub
    Private Sub cmbDSN_TextChanged(sender As Object, e As EventArgs) Handles cmbDSN.TextChanged

    End Sub

    Private Sub GetDSNs()
        Dim strKeyNames() As String
        Dim intKeyCount As Integer
        Dim intCount As Integer
        Dim key As Microsoft.Win32.RegistryKey
        cmbDSN.Items.Add("Please choose...")
        key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("software\odbc\odbc.ini\odbc data sources")
        strKeyNames = key.GetValueNames() 'Get an array of the key names
        intKeyCount = key.ValueCount() 'Get the number of keys
        For intCount = 0 To intKeyCount - 1
            cmbDSN.Items.Add(strKeyNames(intCount))
        Next
        key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("software\odbc\odbc.ini\odbc data sources")
        strKeyNames = key.GetValueNames() 'Get an array of the value names
        intKeyCount = key.ValueCount() 'Get the number of values
        For intCount = 0 To intKeyCount - 1
            cmbDSN.Items.Add(strKeyNames(intCount))
        Next
        cmbDSN.SelectedIndex = 0
    End Sub

    Private Sub btnSQLStatement_Click_1(sender As Object, e As EventArgs) Handles btnSQLStatement.Click
        Me.Cursor = Cursors.WaitCursor
        frmSQL.Show()
        Me.Cursor = Cursors.Default
    End Sub




End Class

