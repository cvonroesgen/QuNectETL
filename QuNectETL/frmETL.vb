Imports System.Data.Odbc
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports QuNectETL.frmETL


Public Class frmETL
    Private Structure connectionStrings
        Public src As String
        Public dst As String
    End Structure
    Enum arg
        configFile = 1
        logFile = 2
    End Enum
    Enum cmbPasswordIndex
        neither = 0
        password = 1
        usertoken = 2
    End Enum
    Public Structure config
        Public sourceConnectionString As String
        Public sourceSQL As String
        Public sourceFieldOrdinals As ArrayList
        Public destinationConnectionString As String
        Public destinationTable As String
        Public destinationFields As ArrayList
        Public logSQL As Boolean
        Overrides Function toString() As String
            Return _
            toStringNoSQL() _
            & "srcSQL: " & sourceSQL & vbCrLf
        End Function
        Public Function toStringNoSQL() As String
            Return _
             "destination: " & destinationTable & vbCrLf &
             "destination fields: " & String.Join(",", destinationFields.ToArray()) & vbCrLf &
             "sourceFieldOrdinals: " & String.Join(",", sourceFieldOrdinals.ToArray()) & vbCrLf
        End Function
        Public Sub loadFromJSON(ByRef json As String)
            jsonPointer = 0
            Dim prop As String
            Dim val As Object
            Dim propValues As New Hashtable
            goToFirstProperty(json)
            While jsonPointer < json.Length And json.Chars(jsonPointer) <> "}"
                prop = getPropertyName(json)
                val = getPropertyValue(json)
                propValues.Add(prop, val)
                skipWhiteSpace(json)
                If json.Chars(jsonPointer) = "," Then
                    jsonPointer += 1 'skip over comma
                    skipWhiteSpace(json)
                End If
            End While

            Me.sourceFieldOrdinals = propValues("sourceFieldOrdinals")
            Me.destinationFields = propValues("destinationFields")
            Me.sourceSQL = propValues("sourceSQL")
            Me.destinationConnectionString = propValues("destinationConnectionString")
            Me.destinationTable = propValues("destinationTable")
            Me.logSQL = propValues("logSQL")
            Me.sourceConnectionString = propValues("sourceConnectionString")
        End Sub

        Private jsonPointer As Integer
        Private Sub goToFirstProperty(ByRef json As String)
            While jsonPointer < json.Length And json.Chars(jsonPointer) <> """"
                jsonPointer += 1
            End While
        End Sub
        Private Function getPropertyName(ByRef json As String) As String
            getPropertyName = ""
            If json.Chars(jsonPointer) <> """" Then
                Throw New Exception("Missing opening double quote on property name." & json.Substring(jsonPointer))
            End If
            jsonPointer += 1
            While jsonPointer < json.Length And json.Chars(jsonPointer) <> """"
                getPropertyName &= json.Chars(jsonPointer)
                jsonPointer += 1
            End While
            While jsonPointer < json.Length And json.Chars(jsonPointer) <> ":"
                jsonPointer += 1
            End While
            jsonPointer += 1 'skip over colon
        End Function
        Private Function getPropertyValue(ByRef json As String) As Object
            getPropertyValue = Nothing
            skipWhiteSpace(json)
            If json.Chars(jsonPointer) = "[" Then
                getPropertyValue = New ArrayList
                jsonPointer += 1 'skip over opening square bracket
                skipWhiteSpace(json)
                While json.Chars(jsonPointer) <> "]"
                    skipWhiteSpace(json)
                    getPropertyValue.add(getJSONString(json))
                    skipWhiteSpace(json)
                    If json.Chars(jsonPointer) = "]" Then Exit While
                    If json.Chars(jsonPointer) <> "," Then
                        Throw New Exception("Missing comma in array." & json.Substring(jsonPointer))
                    End If
                    jsonPointer += 1 'skip over comma
                End While
                jsonPointer += 1 'skip over closing square bracket
            ElseIf json.Chars(jsonPointer) = """" Then
                getPropertyValue = getJSONString(json)
            Else
                Dim token = getJSONToken(json)
                If Regex.IsMatch(token, "^true$", RegexOptions.IgnoreCase) Then
                    getPropertyValue = True
                ElseIf Regex.IsMatch(token, "^false$", RegexOptions.IgnoreCase) Then
                    getPropertyValue = False
                Else
                    getPropertyValue = Convert.ToDecimal(token)
                End If
            End If
        End Function
        Private Sub skipWhiteSpace(ByRef json As String)
            While jsonPointer < json.Length And String.IsNullOrWhiteSpace(json.Chars(jsonPointer))
                jsonPointer += 1
            End While
        End Sub
        Private Function getJSONString(ByRef json As String) As String
            getJSONString = ""
            If json.Chars(jsonPointer) <> """" Then
                Throw New Exception("Missing opening double quote on property name. " & json.Substring(jsonPointer))
            End If
            jsonPointer += 1 'skip over opening double quote
            While jsonPointer < json.Length And json.Chars(jsonPointer) <> """"
                If json.Chars(jsonPointer) = "\" And json.Chars(jsonPointer + 1) = """" Then
                    jsonPointer += 1
                End If
                getJSONString &= json.Chars(jsonPointer)
                jsonPointer += 1
            End While
            jsonPointer += 1 'skip over closing double quote
        End Function
        Private Function getJSONToken(ByRef json As String) As String
            getJSONToken = ""
            While jsonPointer < json.Length AndAlso Regex.IsMatch(json.Chars(jsonPointer), "[\\d\\-\\.a-z]", RegexOptions.IgnoreCase)
                getJSONToken &= json.Chars(jsonPointer)
                jsonPointer += 1
            End While

        End Function
    End Structure

    Private Class odbcField
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
    Public Const AppName = "QuNectETL"
    Private Const fieldTypeDelimiter = ":"
    Private Const fieldDelimiter = ","
    Private Const ordinalDelimter = "."
    Private Const truncateSQL = 100
    Private Title = "QuNect ETL"
    Private odbcConnections As New Dictionary(Of String, OdbcConnection)
    Private rids As HashSet(Of String)
    Private keyfid As String
    Private uniqueExistingFieldValues As Dictionary(Of String, HashSet(Of String))
    Private uniqueNewFieldValues As Dictionary(Of String, HashSet(Of String))
    Private sourceFieldNames As New Dictionary(Of String, Integer)
    Private ReadOnly isBooleanTrue As New Regex("y|tr|c|[1-9]", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Private progressMessage As String
    Private sourceConnection As OdbcConnection
    Private destinationConnection As OdbcConnection
    Private destinationFieldNameToType As Dictionary(Of String, String)
    Private destinationFIDToFieldName As New Dictionary(Of String, String)
    Private destinationFieldNameToFID As New Dictionary(Of String, String)

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
    Private logFile As StreamWriter
    Dim cnfg As New config
    Private Sub ETL_Load(sender As Object, e As EventArgs) Handles Me.Load
        dgMapping.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        cmdLineArgs = System.Environment.GetCommandLineArgs()
        If cmdLineArgs.Length > arg.configFile Then
            automode = True

            Try
                loadConfigFromFile(cmdLineArgs(arg.configFile))
                If cmdLineArgs.Length > arg.logFile Then
                    'open log file
                    logFile = New StreamWriter(File.Open(cmdLineArgs(arg.logFile), FileMode.Append))
                    logFile.WriteLine()
                    logFile.WriteLine(DateTime.Now)
                    logFile.WriteLine("Running Job File: " & cmdLineArgs(arg.configFile))
                    If cnfg.logSQL Then
                        logFile.Write(cnfg.toString())
                    Else
                        logFile.Write(cnfg.toStringNoSQL())
                    End If
                End If
                txtSQL.Text = cnfg.sourceSQL
                Dim cnctStrings As connectionStrings
                cnctStrings.src = txtSourceConnectionString.Text
                cnctStrings.dst = txtDestinationConnectionString.Text
                Dim fieldNodes As New ArrayList

                executeUpload(cnfg)
                If cmdLineArgs.Length > arg.logFile Then
                    logFile.WriteLine(DateTime.Now)
                    logFile.WriteLine("Finished Job")
                    logFile.Flush()
                    logFile.Close()
                End If
            Catch excpt As Exception
                Alert(excpt.Message)
            Finally
                Me.Close()
            End Try
            Exit Sub
        Else
            automode = False
        End If
        Me.Cursor = Cursors.WaitCursor
        Dim myBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
        Title &= " " & myBuildInfo.ProductVersion
        Text = Title
        Dim sourceDSN As String = GetSetting(AppName, "Connection", "sourceDSN", "")
        Dim destinationDSN As String = GetSetting(AppName, "Connection", "destinationDSN", "")
        txtSourceUID.Text = GetSetting(AppName, "Credentials", "sourceUID", "")
        txtDestinationUID.Text = GetSetting(AppName, "Credentials", "destinationUID", "")
        txtSourcePWD.Text = GetSetting(AppName, "Credentials", "sourcePWD", "")
        txtDestinationPWD.Text = GetSetting(AppName, "Credentials", "destinationPWD", "")
        GetDSNs()
        cmbSourceDSN.SelectedIndex = cmbSourceDSN.FindStringExact(sourceDSN)
        If cmbSourceDSN.SelectedIndex = -1 Then
            cmbSourceDSN.SelectedIndex = 0
        End If
        cmbDestinationDSN.SelectedIndex = cmbDestinationDSN.FindStringExact(destinationDSN)
        If cmbDestinationDSN.SelectedIndex = -1 Then
            cmbDestinationDSN.SelectedIndex = 0
        End If
        txtSourceConnectionString.Text = GetSetting(AppName, "Connection", "sourceConnectionString", "")
        If txtSourceConnectionString.Text <> "" Then
            rdbSourceConnectionString.Checked = True
        End If
        txtDestinationConnectionString.Text = GetSetting(AppName, "Connection", "destinationConnectionString", "")
        If txtDestinationConnectionString.Text <> "" Then
            rdbDestinationConnectionString.Checked = True
        End If
        lblDestinationTable.Text = GetSetting(AppName, "config", "destinationtable", "")
        txtSQL.Text = GetSetting(AppName, "config", "SQL", "")

        showHideControls()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        saveConfig()
    End Sub
    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        Try
            loadConfigFromFile("")
        Catch excpt As Exception
            Alert("Could not load job. " & excpt.Message)
        End Try
        Me.Cursor = Cursors.Default
    End Sub

    Sub loadConfigFromFile(filename As String)
        Dim jobFileReader As System.IO.StreamReader = Nothing
        Try
            If Not automode AndAlso openFile.ShowDialog = Windows.Forms.DialogResult.OK Then
                filename = openFile.FileName
                lblJobFile.Text = filename
            End If
            jobFileReader = My.Computer.FileSystem.OpenTextFileReader(filename)
            Dim firstLine As String = jobFileReader.ReadLine()
            If firstLine Is Nothing Then
                Throw New Exception("Job file is blank.")
            End If
            If firstLine.StartsWith("{") Then
                loadJsonFile(firstLine, jobFileReader)

            Else
                loadOldConfig(jobFileReader, firstLine)
                loadConfigFromUI()
            End If
        Catch excpt As Exception
            Throw New Exception("Could not read job file: " & excpt.Message)
        Finally
            If Not jobFileReader Is Nothing Then
                jobFileReader.Close()
            End If
        End Try
    End Sub
    Sub loadUIfromConfig()
        txtDestinationConnectionString.Text = cnfg.destinationConnectionString
        txtSourceConnectionString.Text = cnfg.sourceConnectionString
        lblDestinationTable.Text = cnfg.destinationTable
        txtSQL.Text = cnfg.sourceSQL
        ckbLogSQL.Checked = cnfg.logSQL
        listFields(lblDestinationTable.Text, txtSQL.Text, False)
        If cnfg.sourceFieldOrdinals.Count = cnfg.destinationFields.Count Then
            For i As Integer = 0 To cnfg.sourceFieldOrdinals.Count - 1
                If Regex.IsMatch(cnfg.sourceFieldOrdinals(i), "^\d+$") Then
                    Dim srcOrdinal As Integer = CInt(cnfg.sourceFieldOrdinals(i))
                    If srcOrdinal > (dgMapping.Rows.Count - 1) Then
                        Throw New Exception("The source SQL statement has fewer columns than this configuration requires.")
                    End If
                    Dim destComboBoxCell As DataGridViewComboBoxCell = dgMapping.Rows(srcOrdinal).Cells(mapping.destination)

                    destComboBoxCell.Value = cnfg.destinationFields(i)
                End If
            Next
        End If


    End Sub

    Sub loadConfigFromUI()
        cnfg.destinationConnectionString = txtDestinationConnectionString.Text
        cnfg.sourceConnectionString = txtSourceConnectionString.Text
        Dim sourceFieldOrdinals As String = ""
        Dim DestinationFields As String = ""
        cnfg.sourceFieldOrdinals = New ArrayList()
        cnfg.destinationFields = New ArrayList()
        For i As Integer = 0 To dgMapping.Rows.Count - 1
            Dim destComboBoxCell As DataGridViewComboBoxCell = dgMapping.Rows(i).Cells(mapping.destination)
            If destComboBoxCell.Value Is Nothing Then Continue For
            Dim destDDIndex = destComboBoxCell.Items.IndexOf(destComboBoxCell.Value)
            If destDDIndex = 0 Then Continue For
            cnfg.sourceFieldOrdinals.Add(CStr(i))
            cnfg.destinationFields.Add(destComboBoxCell.Value)
        Next
        cnfg.destinationTable = lblDestinationTable.Text
        cnfg.sourceSQL = txtSQL.Text
        cnfg.logSQL = ckbLogSQL.Checked

    End Sub
    Sub loadOldConfig(jobFileReader As System.IO.StreamReader, firstLine As String)
        txtDestinationConnectionString.Text = "Driver={QuNect ODBC for QuickBase};FIELDNAMECHARACTERS=all;ALLREVISIONS=ALL"
        txtDestinationUID.Text = firstLine
        txtDestinationConnectionString.Text &= ";UID=" + firstLine
        txtDestinationPWD.Text = jobFileReader.ReadLine()
        txtDestinationConnectionString.Text &= ";PWD=" + txtDestinationPWD.Text
        Dim pwdIsPassword As String = jobFileReader.ReadLine()

        Select Case CInt(pwdIsPassword)
            Case cmbPasswordIndex.password
                txtDestinationConnectionString.Text &= ";PWDISPASSWORD=1"
            Case cmbPasswordIndex.usertoken
                txtDestinationConnectionString.Text &= ";PWDISPASSWORD=0"
            Case Else
                txtDestinationConnectionString.Text &= ";PWDISPASSWORD=0"
        End Select

        txtDestinationConnectionString.Text &= ";QUICKBASESERVER=" & jobFileReader.ReadLine()
        txtDestinationConnectionString.Text &= ";APPTOKEN=" & jobFileReader.ReadLine()
        lblDestinationTable.Text = jobFileReader.ReadLine()
        txtDestinationConnectionString.Text &= ";DETECTPROXY=" & jobFileReader.ReadLine()
        rdbDestinationConnectionString.Checked = True
        Dim nextLine As String = jobFileReader.ReadLine()
        cmbSourceDSN.SelectedIndex = cmbSourceDSN.FindStringExact(nextLine)
        txtSourceConnectionString.Text = "DSN=" & nextLine


        '####################################
        cnfg.sourceFieldOrdinals = New ArrayList(jobFileReader.ReadLine().Split(ordinalDelimter))
        Dim fidsForImport As String = jobFileReader.ReadLine()
        '####################################
        nextLine = jobFileReader.ReadLine()
        If nextLine.Length Then
            txtSourceConnectionString.Text &= ";UID=" & nextLine
            txtSourceUID.Text = nextLine
        End If
        nextLine = jobFileReader.ReadLine()
        If nextLine.Length Then
            txtSourceConnectionString.Text &= ";PWD=" & nextLine
            txtSourcePWD.Text = nextLine
        End If
        rdbSourceConnectionString.Checked = True
        Dim logSQL As String = jobFileReader.ReadLine()
        If logSQL.Length > 0 AndAlso CInt(logSQL) > 0 Then
            ckbLogSQL.Checked = True
        Else
            ckbLogSQL.Checked = False
        End If
        txtSQL.Text = jobFileReader.ReadLine()
        While Not jobFileReader.EndOfStream
            txtSQL.Text &= jobFileReader.ReadLine() & vbCrLf
        End While

        listFields(lblDestinationTable.Text, txtSQL.Text, False)
        Dim fids As String() = fidsForImport.Split(fieldDelimiter)
        If cnfg.sourceFieldOrdinals.Count = fids.Length Then
            For i As Integer = 0 To cnfg.sourceFieldOrdinals.Count - 1
                If Regex.IsMatch(cnfg.sourceFieldOrdinals(i), "^\d+$") Then
                    Dim srcOrdinal As Integer = CInt(cnfg.sourceFieldOrdinals(i))
                    If srcOrdinal > (dgMapping.Rows.Count - 1) Then
                        Throw New Exception("The source SQL statement has fewer columns than this configuration requires.")
                    End If
                    Dim destComboBoxCell As DataGridViewComboBoxCell = dgMapping.Rows(srcOrdinal).Cells(mapping.destination)
                    Dim fidType As New ArrayList(fids(i).Split(fieldTypeDelimiter))
                    Dim fid As String = fidType(0)
                    destComboBoxCell.Value = destinationFIDToFieldName(fid.Substring(3))
                End If
            Next
        End If
    End Sub

    Sub saveConfig()
        Me.Cursor = Cursors.WaitCursor
        loadConfigFromUI()
        If destinationFieldNameToFID.Count > 0 Then
            For i As Integer = 0 To cnfg.destinationFields.Count - 1
                cnfg.destinationFields(i) = destinationFieldNameToFID(cnfg.destinationFields(i))
            Next
        End If
        saveDialog.Filter = "JOB Files (*.job)|*.job"
        saveDialog.FileName = lblJobFile.Text
        If saveDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim strJob As String = "{"
            strJob &= createJSONString("sourceConnectionString", cnfg.sourceConnectionString, False)
            strJob &= createJSONString("sourceSQL", cnfg.sourceSQL, True)
            strJob &= createJSONArray("sourceFieldOrdinals", cnfg.sourceFieldOrdinals, True)
            strJob &= createJSONString("destinationConnectionString", cnfg.destinationConnectionString, True)
            strJob &= createJSONString("destinationTable", cnfg.destinationTable, True)
            strJob &= createJSONArray("destinationFields", cnfg.destinationFields, True)
            strJob &= createJSONBoolean("logSQL", cnfg.logSQL, True)
            strJob &= "}"
            My.Computer.FileSystem.WriteAllText(saveDialog.FileName, strJob, False, Encoding.ASCII)
        End If
        lblJobFile.Text = saveDialog.FileName
        showHideControls()
        Me.Cursor = Cursors.Default
    End Sub
    Function createJSONString(ByRef name As String, ByRef value As String, comma As Boolean) As String
        createJSONString = """" & name & """:""" & Replace(value, """", "\""") & """"
        If comma Then
            createJSONString = "," & createJSONString
        End If
    End Function
    Function createJSONBoolean(ByRef name As String, bool As Boolean, comma As Boolean) As String
        Dim value As String = "false"
        If bool Then value = "true"
        createJSONBoolean = """" & name & """:""" & value & """"
        If comma Then
            createJSONBoolean = "," & createJSONBoolean
        End If
    End Function
    Function createJSONArray(ByRef name As String, ByRef ary As ArrayList, notfirst As Boolean) As String
        Dim strArray As String = ""
        Dim comma As String = ""
        For Each val As String In ary
            strArray &= comma & """" & Replace(val, """", "\""") & """"
            comma = ","
        Next
        createJSONArray = """" & name & """:[" & strArray & "]"
        If notfirst Then
            createJSONArray = "," & createJSONArray
        End If
    End Function
    Sub loadJsonFile(json As String, jobFileReader As System.IO.StreamReader)
        While Not jobFileReader.EndOfStream
            json &= jobFileReader.ReadLine & vbCrLf
        End While
        cnfg.loadFromJSON(json)
        loadUIfromConfig()
        listDestinationFields(cnfg.destinationTable)
        If destinationFieldNameToFID.Count > 0 Then
            For i As Integer = 0 To cnfg.destinationFields.Count - 1
                cnfg.destinationFields(i) = destinationFIDToFieldName(cnfg.destinationFields(i))
            Next
        End If
    End Sub
    Sub showHideControls()
        btnListFields.Visible = False
        btnDestination.Visible = True
        btnImport.Visible = False
        dgMapping.Visible = False
        btnSave.Visible = False
        btnCommandLine.Visible = False

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
        SaveSetting(AppName, "Connection", "sourceDSN", cmbSourceDSN.Text)
        SaveSetting(AppName, "Connection", "destinationDSN", cmbDestinationDSN.Text)
        SaveSetting(AppName, "Credentials", "sourceUID", txtSourceUID.Text)
        SaveSetting(AppName, "Credentials", "destinationUID", txtDestinationUID.Text)
        SaveSetting(AppName, "Credentials", "sourcePWD", txtSourcePWD.Text)
        SaveSetting(AppName, "Credentials", "destinationPWD", txtDestinationPWD.Text)
        If rdbDestinationDSN.Checked And cmbDestinationDSN.Text <> "" Then
            txtDestinationConnectionString.Text = "DSN=" & cmbDestinationDSN.Text & ";"
            If txtDestinationUID.TextLength > 0 Then
                txtDestinationConnectionString.Text &= "UID=" & txtDestinationUID.Text & ";"
            End If
            If txtDestinationPWD.TextLength > 0 Then
                txtDestinationConnectionString.Text &= "PWD=" & txtDestinationPWD.Text & ";"
            End If
        End If
        If rdbSourceDSN.Checked And cmbSourceDSN.Text <> "" Then
            txtSourceConnectionString.Text = "DSN=" & cmbSourceDSN.Text & ";"
            If txtSourceUID.TextLength > 0 Then
                txtSourceConnectionString.Text &= "UID=" & txtSourceUID.Text & ";"
            End If
            If txtSourcePWD.TextLength > 0 Then
                txtSourceConnectionString.Text &= "PWD=" & txtSourcePWD.Text & ";"
            End If
        End If
        If lblJobFile.Text <> "" Then
            btnCommandLine.Visible = True
        End If
    End Sub
    Private Sub lblDestinationTable_TextChanged(sender As Object, e As EventArgs) Handles lblDestinationTable.TextChanged
        SaveSetting(AppName, "config", "destinationtable", lblDestinationTable.Text)
        showHideControls()
    End Sub


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
                Throw New System.Exception("Could not count records " & excpt.Message)
            End Try
            If Not dr.HasRows Then
                Throw New System.Exception("Counting records failed. " & sql)
            Else
                Return dr(0)
            End If
        End Using
    End Function
    Function setODBCParameter(ByRef odbcTyp As OdbcType, ByRef val As Object, fid As String, ByRef command As OdbcCommand, ByRef conversionErrors As String) As Boolean

        If IsDBNull(val) Then
            command.Parameters("@fid" & fid).Value = val
            Return False
        End If
        Try
            Select Case odbcTyp
                Case OdbcType.Int
                    If IsDBNull(val) Then
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
            conversionErrors = ex.Message
            Return True
        End Try
        Return False
    End Function
    Private Function getODBCType(strType As String) As OdbcType
        Select Case strType.ToLower
            Case "bigint"
                Return OdbcType.BigInt
            Case "binary"
                Return OdbcType.Binary
            Case "bit"
                Return OdbcType.Bit
            Case "char"
                Return OdbcType.Char
            Case "datetime"
                Return OdbcType.DateTime
            Case "decimal"
                Return OdbcType.Numeric
            Case "numeric"
                Return OdbcType.Numeric
            Case "double"
                Return OdbcType.Double
            Case "image"
                Return OdbcType.Image
            Case "int"
                Return OdbcType.Int
            Case "nchar"
                Return OdbcType.NChar
            Case "ntext"
                Return OdbcType.NText
            Case "nvarchar"
                Return OdbcType.NVarChar
            Case "real"
                Return OdbcType.Real
            Case "uniqueidentifier"
                Return OdbcType.UniqueIdentifier
            Case "smalldatetime"
                Return OdbcType.SmallDateTime
            Case "smallint"
                Return OdbcType.SmallInt
            Case "text"
                Return OdbcType.Text
            Case "timestamp"
                Return OdbcType.Timestamp
            Case "tinyint"
                Return OdbcType.TinyInt
            Case "varbinary"
                Return OdbcType.VarBinary
            Case "varchar"
                Return OdbcType.VarChar
            Case "date"
                Return OdbcType.Date
            Case "time"
                Return OdbcType.Time
            Case Else
                Return OdbcType.VarChar
        End Select
    End Function

    Private Function uploadToDestination(cnctStrings As connectionStrings) As Boolean
        Dim destinationFields As New ArrayList
        Dim sourceFieldOrdinals As New ArrayList
        loadConfigFromUI()
        For i As Integer = 0 To dgMapping.Rows.Count - 1
            Dim destComboBoxCell As DataGridViewComboBoxCell = DirectCast(dgMapping.Rows(i).Cells(mapping.destination), System.Windows.Forms.DataGridViewComboBoxCell)
            If destComboBoxCell.Value Is Nothing Then Continue For
            Dim destDDIndex = destComboBoxCell.Items.IndexOf(destComboBoxCell.Value)
            If destDDIndex = 0 Then Continue For
            destinationFields.Add(destComboBoxCell.Value)
            sourceFieldOrdinals.Add(i)
        Next
        If sourceFieldOrdinals.Count = 0 And Not automode Then
            Alert("You must map at least one field from the source table to the destination table.", MsgBoxStyle.OkOnly, AppName)
            Me.Cursor = Cursors.Default
            Return False
        End If
        Dim copyThread As System.Threading.Thread = New Threading.Thread(AddressOf executeUpload)
        copyThread.Start(cnfg)
        Me.Cursor = Cursors.Default
        Return True
    End Function
    Private Function executeUpload(upCnfg As config) As Boolean
        Try
            Dim strDestinationSQL As String = "INSERT INTO """ & upCnfg.destinationTable & """ ("""
            Dim destinationConnection As OdbcConnection = getODBCConnection(upCnfg.destinationConnectionString)
            Try
                strDestinationSQL &= String.Join(""", """, upCnfg.destinationFields.ToArray) & """) VALUES ("
                For Each var In upCnfg.destinationFields
                    strDestinationSQL &= "?,"
                Next
                If strDestinationSQL.Length > 0 Then
                    strDestinationSQL = strDestinationSQL.Substring(0, strDestinationSQL.Length - 1)
                End If
                strDestinationSQL &= ")"
                Using destinationCommand As OdbcCommand = New OdbcCommand(strDestinationSQL, destinationConnection)
                    Dim qdbTypes(upCnfg.destinationFields.Count) As OdbcType
                    Dim j As Integer
                    For j = 0 To upCnfg.destinationFields.Count - 1
                        qdbTypes(j) = getODBCType(destinationFieldNameToType(upCnfg.destinationFields(j)))
                        destinationCommand.Parameters.Add("@fid" & j, qdbTypes(j))
                    Next

                    Dim transaction As OdbcTransaction = Nothing
                    transaction = destinationConnection.BeginTransaction()
                    destinationCommand.Transaction = transaction
                    destinationCommand.CommandType = CommandType.Text
                    destinationCommand.CommandTimeout = 0
                    If Not automode Then
                        Dim progressThread As System.Threading.Thread = New Threading.Thread(AddressOf showProgress)
                        Volatile.Write(progressMessage, "Initializing...")
                        progressThread.Start()
                    End If
                    'we have to open up a reader on the source
                    Dim srcConnection As OdbcConnection
                    srcConnection = New OdbcConnection(upCnfg.sourceConnectionString)
                    srcConnection.Open()
                    Dim fileLineCounter As Integer = 0
                    Dim conversionErrors As String = ""
                    Using srcCmd As OdbcCommand = New OdbcCommand(upCnfg.sourceSQL, srcConnection)
                        Dim dr As OdbcDataReader
                        Try
                            Dim sourceFieldMaxIndex As Integer = upCnfg.sourceFieldOrdinals.Count - 1
                            dr = srcCmd.ExecuteReader()
                            While (dr.Read())
                                For i As Integer = 0 To sourceFieldMaxIndex
                                    If setODBCParameter(qdbTypes(i), dr.GetValue(upCnfg.sourceFieldOrdinals(i)), CStr(i), destinationCommand, conversionErrors) Then
                                        Throw New System.Exception("Could not convert '" & dr.GetValue(upCnfg.sourceFieldOrdinals(i)) & "' to " & qdbTypes(i))
                                    End If
                                Next
                                If fileLineCounter Mod 1000 = 0 And Not automode Then
                                    Volatile.Write(progressMessage, "Queuing up record " & fileLineCounter)
                                End If
                                fileLineCounter += destinationCommand.ExecuteNonQuery()
                            End While
                            If Not automode Then
                                Volatile.Write(progressMessage, "Committing " & fileLineCounter & " records")
                            End If
                            transaction.Commit()
                            If Not automode Then
                                Volatile.Write(progressMessage, "Committed " & fileLineCounter & " records")
                            End If
                            Alert(fileLineCounter & " rows were uploaded to " & destinationConnection.DataSource & ".")
                        Catch excpt As Exception
                            srcCmd.Cancel()
                            srcCmd.Dispose()
                            transaction.Rollback()
                            srcConnection.Close()
                            Throw New System.Exception("Could not get record " & fileLineCounter & " from " & upCnfg.sourceConnectionString & vbCrLf & excpt.Message)
                        End Try
                    End Using
                End Using
            Catch e As Exception
                Alert("Could not copy because " & e.Message)
            Finally

            End Try
        Catch ex As Exception
            If Not automode Then
                Volatile.Write(progressMessage, "")
            End If
            Alert("Could not copy because " & ex.Message)
        End Try
        If Not automode Then
            Volatile.Write(progressMessage, "")
        End If
        Return True

    End Function
    Function Alert(msg As String) As MsgBoxResult
        If automode Then
            If cmdLineArgs.Length > arg.logFile Then
                logFile.WriteLine(msg)
                logFile.Flush()
            End If
        Else
            Return MsgBox(msg)
        End If
        Return MsgBoxResult.Ok
    End Function
    Function Alert(msg As String, style As MsgBoxStyle) As MsgBoxResult
        If automode Then
            If cmdLineArgs.Length > arg.logFile Then
                logFile.WriteLine(msg)
                logFile.Flush()
            End If
        Else
            Me.Cursor = Cursors.Default
            Return MsgBox(msg, style)
        End If
        Return MsgBoxResult.Ok
    End Function
    Function Alert(msg As String, style As MsgBoxStyle, title As String) As MsgBoxResult
        If automode Then
            If cmdLineArgs.Length > arg.logFile Then
                logFile.WriteLine(msg)
                logFile.Flush()
            End If
        Else
            Me.Cursor = Cursors.Default
            Return MsgBox(msg, style, title)
        End If
        Return MsgBoxResult.Ok
    End Function
    Private Function getODBCConnection(connectionString As String) As OdbcConnection
        If odbcConnections.ContainsKey(connectionString) Then
            Return odbcConnections(connectionString)
        End If
        Dim odbcConn As OdbcConnection = New OdbcConnection(connectionString)
        odbcConn.Open()
        Return odbcConn
    End Function
    Private Sub btnDestination_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDestination.Click
        If cmbDestinationDSN.SelectedIndex <= 0 And rdbDestinationDSN.Checked Then
            Alert("Please choose a DSN.")
            Return
        End If
        If txtDestinationConnectionString.Text.Length = 0 And rdbDestinationConnectionString.Checked Then
            Alert("Please enter a connection string.")
            Return
        End If
        If rdbDestinationDSN.Checked Then
            txtDestinationConnectionString.Text = "DSN=" & cmbDestinationDSN.Text & ";"
        End If
        showHideControls()
        listTables(txtDestinationConnectionString.Text)
    End Sub

    Private Sub listTables(connectionString As String)
        Me.Cursor = Cursors.WaitCursor
        Try
            Dim odbcConn As OdbcConnection = getODBCConnection(connectionString)
            Dim tables As DataTable = odbcConn.GetSchema("Tables")
            Dim views As DataTable = odbcConn.GetSchema("Views")
            tables.Merge(views)
            listTablesFromGetSchema(tables)

        Catch ex As Exception
            Alert("Could not list tables because " & ex.Message)
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Sub listTablesFromGetSchema(tables As DataTable)
        frmTableChooser.tvAppsTables.BeginUpdate()
        frmTableChooser.tvAppsTables.Nodes.Clear()
        frmTableChooser.tvAppsTables.ShowNodeToolTips = True
        Dim dbName As String
        Dim applicationName As String = ""
        Dim prevAppName As String = ""


        For i = 0 To tables.Rows.Count - 1

            Application.DoEvents()
            dbName = tables.Rows(i)(2)
            applicationName = tables.Rows(i)(0)
            If applicationName <> prevAppName Then

                Dim appNode As TreeNode = frmTableChooser.tvAppsTables.Nodes.Add(applicationName)
                prevAppName = applicationName
            End If
            Dim tableName As String = dbName
            If applicationName.Length Then
                Dim tableNode As TreeNode = frmTableChooser.tvAppsTables.Nodes(frmTableChooser.tvAppsTables.Nodes.Count - 1).Nodes.Add(tableName)

            Else
                Dim tableNode As TreeNode = frmTableChooser.tvAppsTables.Nodes.Add(tableName)

            End If

        Next
        frmTableChooser.Text = "Choose a Table"
        frmTableChooser.tvAppsTables.EndUpdate()

        btnImport.Visible = True
        lblDestinationTable.Visible = True
        frmTableChooser.Show()
        Me.Cursor = Cursors.Default
    End Sub
    Function listDestinationFields(tableName As String) As Dictionary(Of String, String)
        listDestinationFields = New Dictionary(Of String, String)
        Try
            Dim connection As OdbcConnection = getODBCConnection(txtDestinationConnectionString.Text)
            If connection Is Nothing Then Exit Function
            Dim restrictions(2) As String
            restrictions(2) = tableName
            Dim columns As DataTable = connection.GetSchema("Columns", restrictions)

            'Loop through all of the fields in the schema.
            Dim dgComboBox As System.Windows.Forms.DataGridViewComboBoxColumn = DirectCast(dgMapping.Columns(mapping.destination), System.Windows.Forms.DataGridViewComboBoxColumn)
            dgComboBox.Items.Clear()
            destinationFIDToFieldName.Clear()
            destinationFieldNameToFID.Clear()
            dgComboBox.Items.Add("")
            For i As Integer = 0 To columns.Rows.Count - 1
                dgComboBox.Items.Add(columns.Rows(i)(3))
                listDestinationFields.Add(columns.Rows(i)(3), columns.Rows(i)(5))
                If connection.DataSource = "QuNect ODBC for QuickBase" Then
                    Dim fid As String = columns.Rows(i)(11).ToString()
                    fid = Regex.Replace(fid, "^.* fid", "")
                    destinationFIDToFieldName.Add(fid, columns.Rows(i)(3))
                    destinationFieldNameToFID.Add(columns.Rows(i)(3), fid)
                End If
            Next
        Catch ex As Exception
            Alert(ex.Message)
        End Try
    End Function
    Function listFields(destinationTable As String, strSourceSQL As String, guess As Boolean) As Dictionary(Of String, String)
        destinationFieldNameToType = listDestinationFields(destinationTable)
        listFields = destinationFieldNameToType
        Try
            'here we need to open the source and get the field names
            sourceFieldNames.Clear()
            Dim sourceColumns As DataTable = getColumnsDataTable(strSourceSQL, txtSourceConnectionString.Text)
            If sourceColumns Is Nothing Then
                Return destinationFieldNameToType
            End If
            dgMapping.Rows.Clear()
            Dim i As Integer = 0
            For Each columnRow As DataRow In sourceColumns.Rows
                Dim field As New odbcField(columnRow(0), columnRow(5).Name)
                dgMapping.Rows.Add(New String() {field.label})
                sourceFieldNames.Add(field.label, i)
                i += 1
            Next

            'End Using
            If guess Then
                For Each row In dgMapping.Rows
                    guessDestination(row.Cells(mapping.source).Value, row.Index)
                Next
            End If
            showHideControls()

            Me.Cursor = Cursors.Default
        Catch ex As Exception
            Alert(ex.Message)
        End Try
    End Function

    Function getColumnsDataTable(strSourceSQL As String, connectionString As String) As DataTable
        Dim srcConnection As OdbcConnection
        srcConnection = New OdbcConnection(connectionString)
        srcConnection.Open()
        Using srcCmd As OdbcCommand = New OdbcCommand(strSourceSQL, srcConnection)
            Dim dr As OdbcDataReader
            Try
                dr = srcCmd.ExecuteReader()
            Catch excpt As Exception
                Alert("Could not get field information from " & cmbSourceDSN.Text & vbCrLf & excpt.Message)
                Return Nothing
            End Try

            Try
                getColumnsDataTable = dr.GetSchemaTable()
                dr.Close()
            Catch excpt As Exception
                If excpt.Message.Contains("SS_TIME_EX") Then
                    Alert("The source contains a Time of Day field which is not supported. Please try a SQL statement that specifies only non time of day columns.")
                    Return Nothing
                End If
                Throw New System.Exception("Could not get field information for " & cmbSourceDSN.Text & " " & excpt.Message)
            End Try
        End Using
    End Function
    Sub guessDestination(sourceFieldName As String, sourceFieldOrdinal As Integer)

        For Each field In destinationFieldNameToType
            Dim fieldName = field.Key
            If (fieldName.ToLower = sourceFieldName.ToLower Or Regex.Replace(fieldName, "[^a-zA-Z]", "_").ToLower = sourceFieldName.ToLower) Then
                DirectCast(dgMapping.Rows(sourceFieldOrdinal).Cells(mapping.destination), System.Windows.Forms.DataGridViewComboBoxCell).Value = fieldName
                Exit For
            End If
        Next

    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        import()
    End Sub
    Sub import()
        Me.Cursor = Cursors.WaitCursor
        Dim cnctStrings As connectionStrings
        cnctStrings.src = txtSourceConnectionString.Text
        cnctStrings.dst = txtDestinationConnectionString.Text
        uploadToDestination(cnctStrings)

    End Sub

    Private Sub btnListFields_Click(sender As Object, e As EventArgs) Handles btnListFields.Click
        If lblDestinationTable.Text = "" Then
            Alert("Please choose a table to copy into.", MsgBoxStyle.OkOnly, AppName)
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        destinationFieldNameToType = listFields(lblDestinationTable.Text, txtSQL.Text, True)
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
            Dim msg As String = " "
            While msg.Length > 0
                msg = Volatile.Read(progressMessage)
                SetLabelProgress(msg)
                Threading.Thread.Sleep(2000)
            End While
            SetLabelProgress("")
        Catch ex As Exception
            'Alert(ex.Message)
        End Try
    End Sub
    Delegate Sub LabelDelegate(progressMessage As String)
    Private Sub SetLabelProgress(ByVal progressMessage As String)

        ' InvokeRequired required compares the thread ID of the  
        ' calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If Me.lblProgress.InvokeRequired Then
            Dim d As New LabelDelegate(AddressOf SetLabelProgress)
            Me.Invoke(d, New Object() {progressMessage})
        Else
            Me.lblProgress.Text = progressMessage
        End If
    End Sub
    Delegate Sub PBDelegate(numRecords As Integer)


    Private Sub cmbDestinationDSN_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDestinationDSN.SelectedIndexChanged
        showHideControls()


    End Sub
    Private Sub cmbSourceDSN_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSourceDSN.SelectedIndexChanged
        showHideControls()

    End Sub
    Private Sub GetDSNs()
        Dim strKeyNames() As String
        Dim intKeyCount As Integer
        Dim intCount As Integer
        Dim key As Microsoft.Win32.RegistryKey
        cmbSourceDSN.Items.Add("Please choose...")
        cmbDestinationDSN.Items.Add("Please choose...")
        Try
            key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("software\odbc\odbc.ini\odbc data sources")
            If key Is Nothing Then
                Alert("Sorry no DSNs to choose from." & vbCrLf & "Please install QuNect ODBC for QuickBase from https://qunect.com")
                Return
            End If
            strKeyNames = key.GetValueNames() 'Get an array of the key names
            intKeyCount = key.ValueCount() 'Get the number of keys
                For intCount = 0 To intKeyCount - 1
                    cmbSourceDSN.Items.Add(strKeyNames(intCount))
                    cmbDestinationDSN.Items.Add(strKeyNames(intCount))
                Next
        Catch ex As Exception
            Alert("Could not access system DSNs. " & vbCrLf & ex.Message)
        End Try

        Try
            key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("software\odbc\odbc.ini\odbc data sources")
            strKeyNames = key.GetValueNames() 'Get an array of the value names
            intKeyCount = key.ValueCount() 'Get the number of values
            For intCount = 0 To intKeyCount - 1
                cmbSourceDSN.Items.Add(strKeyNames(intCount))
                cmbDestinationDSN.Items.Add(strKeyNames(intCount))
            Next
        Catch ex As Exception
            Alert("Could not access user DSNs. " & vbCrLf & ex.Message)
        End Try

        cmbSourceDSN.SelectedIndex = 0
        cmbDestinationDSN.SelectedIndex = 0
    End Sub

    Protected Overrides Sub Finalize()
        For Each connection In odbcConnections
            connection.Value.Close()
        Next
        MyBase.Finalize()
    End Sub

    Private Sub btnSourceTable_Click(sender As Object, e As EventArgs) Handles btnSourceTable.Click
        If cmbSourceDSN.SelectedIndex <= 0 And rdbSourceDSN.Checked Then
            Alert("Please choose a DSN.")
            Return
        End If
        If txtSourceConnectionString.Text.Length = 0 And rdbSourceConnectionString.Checked Then
            Alert("Please enter a connection string.")
            Return
        End If
        If rdbSourceDSN.Checked Then
            txtSourceConnectionString.Text = "DSN=" & cmbSourceDSN.Text & ";"
        End If
        showHideControls()
        listTables(txtSourceConnectionString.Text)

    End Sub

    Private Sub txtSourceConnectionString_TextChanged(sender As Object, e As EventArgs) Handles txtSourceConnectionString.TextChanged
        SaveSetting(AppName, "Connection", "sourceConnectionString", txtSourceConnectionString.Text)
    End Sub
    Private Sub txtDestinationConnectionString_TextChanged(sender As Object, e As EventArgs) Handles txtDestinationConnectionString.TextChanged
        SaveSetting(AppName, "Connection", "destinationConnectionString", txtDestinationConnectionString.Text)
    End Sub

    Private Sub rdbDestinationDSN_CheckedChanged(sender As Object, e As EventArgs) Handles rdbDestinationDSN.CheckedChanged
        If rdbDestinationDSN.Checked Then
            txtDestinationConnectionString.Enabled = False
            cmbDestinationDSN.Enabled = True
            txtDestinationPWD.Enabled = True
            txtDestinationUID.Enabled = True
        Else
            txtDestinationConnectionString.Enabled = True
            cmbDestinationDSN.Enabled = False
            txtDestinationPWD.Enabled = False
            txtDestinationUID.Enabled = False
        End If
    End Sub
    Private Sub rdbSourceDSN_CheckedChanged(sender As Object, e As EventArgs) Handles rdbSourceDSN.CheckedChanged
        If rdbSourceDSN.Checked Then
            txtSourceConnectionString.Enabled = False
            cmbSourceDSN.Enabled = True
            txtSourcePWD.Enabled = True
            txtSourceUID.Enabled = True
        Else
            txtSourceConnectionString.Enabled = True
            cmbSourceDSN.Enabled = False
            txtSourcePWD.Enabled = False
            txtSourceUID.Enabled = False
        End If
    End Sub



    Private Sub txtSQL_TextChanged(sender As Object, e As EventArgs) Handles txtSQL.TextChanged
        SaveSetting(AppName, "config", "SQL", txtSQL.Text)
        showHideControls()
    End Sub

    Private Sub dgMapping_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles dgMapping.DataError
        'Alert("Mapping Grid error on row " & e.RowIndex & " column " & e.ColumnIndex & " " & e.Exception.Message)
    End Sub

    Private Sub txtDestinationPWD_TextChanged(sender As Object, e As EventArgs) Handles txtDestinationPWD.TextChanged
        showHideControls()
    End Sub

    Private Sub txtDestinationUID_TextChanged(sender As Object, e As EventArgs) Handles txtDestinationUID.TextChanged
        showHideControls()
    End Sub

    Private Sub txtSourcePWD_TextChanged(sender As Object, e As EventArgs) Handles txtSourcePWD.TextChanged
        showHideControls()
    End Sub

    Private Sub txtSourceUID_TextChanged(sender As Object, e As EventArgs) Handles txtSourceUID.TextChanged
        showHideControls()
    End Sub

    Private Sub btnPreview_Click(sender As Object, e As EventArgs) Handles btnPreview.Click
        Me.Cursor = Cursors.WaitCursor
        frmPreview.sql = txtSQL.Text
        frmPreview.connectionString = txtSourceConnectionString.Text
        frmPreview.ShowDialog()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btnPreviewDestination_Click(sender As Object, e As EventArgs) Handles btnPreviewDestination.Click
        If lblDestinationTable.Text = "" Then
            Alert("Please choose a destination table first.")
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        frmPreview.sql = "SELECT * FROM """ & lblDestinationTable.Text & """"
        frmPreview.connectionString = txtDestinationConnectionString.Text
        frmPreview.ShowDialog()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub btnCommandLine_Click(sender As Object, e As EventArgs) Handles btnCommandLine.Click
        loadConfigFromUI()
        Dim programScript As String = """" & cmdLineArgs(0) & """"
        saveDialog.Filter = "BATCH Files (*.bat)|*.bat"
        saveDialog.FileName = ""
        If saveDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            My.Computer.FileSystem.WriteAllText(saveDialog.FileName, programScript & " """ & lblJobFile.Text & """ """ & Replace(lblJobFile.Text, ".job", ".log", 1, 1) & """", False, Encoding.ASCII)
        End If
    End Sub
End Class


