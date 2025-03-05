Imports System.IO
Imports System.Configuration
Imports System.Net.Mail
Imports System.Net
Imports System.Text
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Runtime.CompilerServices

Module IntegrationFunctions
    Public ProgramName As String
    Public IntegrationStarts As Boolean = False
    Public IntegrationYearStarts As Integer = 1950
    Public ErrorFilePath, SuccessFilePath, FilePath, ProcessPath As String
    Public ErrorDump, SuccessDump, FileDump, ProcessDump, ErrorFileDumpPath, SuccessFileDumpPath As String
    Public SalesPath, TenderPath, ReturnPath As String
    Public ErrMsg As String
    Public ServerName, UserId, UserPassword, YearSelected, CompanySelected As String
    Public UserLogged As Boolean = False
    Public ErrNumber As Long
    Public TimerCount As Integer
    Public ReturnMessage As String
    'Public ActionReturn As Boolean = True
    Public RowChecked As Boolean = True
    Public FormObjectType As String
    Public FormKey As String
    Public DBBackupPath As String
    Public Reprocess As Boolean
    Public Today As String
    Public ftpserver, ftpuserid, ftppassword, ftpinbound, ftpoutbound, destination As String
    Public pathtoload As String
    Public dtProcess, dtError, dtSuccess, dtLoadType As DataTable
    Public formgridview As DataGridView
    Public SLayerUrl As String

    'SQL
    Public SQLConnectionString, SQLB1ConnectionString, SQLTARGETB1ConnectionString As String
    Public SQLQuery As String
    Public SQLDBName, SQLDBForQuery, SQLUser, SQLPassword, SQLServer, SQLForCheckingDBName, SQLServerType As String
    Public SQLDBExists As Boolean = False

    Public MinuteRetry As String
    Public FileLocationPath As String
    Public ReConnect As Boolean

    'SAPB1

    Public SAPB1Company As SAPbobsCOM.Company
    Public SAPB1Company_2 As SAPbobsCOM.Company

    Public SAPRetVal As Long
    Public SAPB1CompanyName As String
    Public SAPB1UseTrusted As Boolean
    Public SAPB1ReturnCode As Long = Nothing

    Public FormDataTable(9) As DataTable

    Public UserDefinedFields As DataTable = New DataTable("UDFS")


    Public FileOriginDump As String
    Public ProcessTime As String

    Public OnProcess As Boolean

    Public SAPB1UserIdtoMessage As String

    Public FMSValue As String


    Public Function getobjcode(ByVal objtype As String)

        Select Case objtype
            Case "13"
                Return "OINV"
            Case "14"
                Return "ORIN"
            Case "18"
                Return "OPCH"
            Case "19"
                Return "ORPC"
            Case Else
                Return "-1"
        End Select

        Return ""
    End Function

    Public Function ForCompany(ByVal as_segment As String, ByVal adt_branches As DataTable) As Boolean
        Dim return_val As Boolean = False

        For Each rows As DataRow In adt_branches.Rows
            If as_segment = rows("Code").ToString() Then
                return_val = True
            End If
        Next

        Return return_val
    End Function

    Public Function GetMonthDesc(ByVal month As Integer)

        Select Case month
            Case 0
                Return "January"
            Case 1
                Return "February"
            Case 2
                Return "March"
            Case 3
                Return "April"
            Case 4
                Return "May"
            Case 5
                Return "June"
            Case 6
                Return "July"
            Case 7
                Return "August"
            Case 8
                Return "September"
            Case 9
                Return "October"
            Case 10
                Return "November"
            Case 11
                Return "December"
        End Select
        Return ""
    End Function

    Public Function GetMonthNum(ByVal month As String) As String

        Select Case month
            Case "January"
                Return "01"
            Case "February"
                Return "02"
            Case "March"
                Return "03"
            Case "April"
                Return "04"
            Case "May"
                Return "05"
            Case "June"
                Return "06"
            Case "July"
                Return "07"
            Case "August"
                Return "08"
            Case "September"
                Return "09"
            Case "October"
                Return "10"
            Case "November"
                Return "11"
            Case "December"
                Return "12"
        End Select

        Return ""
    End Function

    Public Function GetGridViewSelectedCount(ByVal grid As DataGridView) As Integer
        Dim count As Integer = 0
        For i As Integer = 0 To grid.Rows.Count - 1
            If grid.Rows(i).Cells(0).Value = True Then
                count += 1
            End If
        Next
        Return count
    End Function

    Public Sub ErrorAppend(ByVal text As String, Optional ByVal Attachment As String = "")

        Dim lb_memoryleaked As Boolean = False

        'FileWrite()
        Dim ErrorLog As FileInfo = New FileInfo(ErrorFilePath)
        Dim StreamWriter As StreamWriter = ErrorLog.AppendText()

        StreamWriter.WriteLine(DateTime.Now.ToString() & "        " & text)
        StreamWriter.Close()



    End Sub
    Public Sub SuccessAppend(ByVal text As String)
        'FileWrite()
        Dim SuccessLog As FileInfo = New FileInfo(SuccessFilePath)
        Dim StreamWriter As StreamWriter = SuccessLog.AppendText()

        StreamWriter.WriteLine(DateTime.Now.ToString() & "        " & text)
        StreamWriter.Close()

    End Sub
    Public Function GetCurrentMemory() As String
        Try
            Dim c As Process = Process.GetCurrentProcess()
            Return ((c.WorkingSet64 / 1024) / 1024).ToString() & " M"
        Catch ex As Exception
            Return "-1 M[" & ex.Message() & "]"
        End Try
        Return "0 M"
    End Function
    Public Sub ExitProgram()
        GC.Collect()
        'Application.Exit()
        Process.GetCurrentProcess.Kill()
    End Sub

    Public Function ExitApp() As Boolean

        If MessageBox.Show("Do you want to exit the application?", ProgramName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            ExitProgram()
            ''''Return False
        Else
            Return True
        End If

        Return True
    End Function


    Public Sub CheckForExistingInstance()

        If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1 Then
            MessageBox.Show(ProgramName + " is already running. Please close first the existing process!")
            Process.GetCurrentProcess.Kill()
        End If
    End Sub
    Public Sub FileWrite(Optional ByVal as_programname As String = "")
        If as_programname = "" Then as_programname = ProgramName

        Dim st_successdump, st_errordump As String



        SuccessFileDumpPath = System.Windows.Forms.Application.StartupPath & "\SuccessDump"
        If Not Directory.Exists(SuccessFileDumpPath) Then
            Directory.CreateDirectory(SuccessFileDumpPath)
        End If
        SuccessFileDumpPath += "\" + Today
        If Not Directory.Exists(SuccessFileDumpPath) Then
            Directory.CreateDirectory(SuccessFileDumpPath)
        End If

        ErrorFileDumpPath = System.Windows.Forms.Application.StartupPath & "\ErrorDump"
        If Not Directory.Exists(ErrorFileDumpPath) Then
            Directory.CreateDirectory(ErrorFileDumpPath)
        End If
        ErrorFileDumpPath += "\" + Today
        If Not Directory.Exists(ErrorFileDumpPath) Then
            Directory.CreateDirectory(ErrorFileDumpPath)
        End If





        If Not Directory.Exists(System.Windows.Forms.Application.StartupPath & "\ProcessPath") Then
            Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath & "\ProcessPath")
        End If


        If Not Directory.Exists(System.Windows.Forms.Application.StartupPath & "\FileDump") Then
            Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath & "\FileDump")
        End If




        ErrorDump = System.Windows.Forms.Application.StartupPath & "\ErrorLog"
        If as_programname <> "" Then
            ErrorDump += "\" + as_programname
        End If

        If Not Directory.Exists(ErrorDump) Then
            Directory.CreateDirectory(ErrorDump)
        End If


        SuccessDump = System.Windows.Forms.Application.StartupPath & "\SuccessLog"
        If as_programname <> "" Then
            SuccessDump += "\" + as_programname
        End If

        If Not Directory.Exists(SuccessDump) Then
            Directory.CreateDirectory(SuccessDump)
        End If


        FileDump = System.Windows.Forms.Application.StartupPath & "\FileDump"

        If as_programname <> "" Then
            FileDump += "\" + as_programname
        End If

        If Not Directory.Exists(FileDump) Then
            Directory.CreateDirectory(FileDump)
        End If

        'SalesPath = FileDump + "\SALES"
        'TenderPath = FileDump + "\TENDER"
        'ReturnPath = FileDump + "\RETURN"


        'If Not Directory.Exists(SalesPath) Then
        '    Directory.CreateDirectory(SalesPath)
        'End If

        'If Not Directory.Exists(TenderPath) Then
        '    Directory.CreateDirectory(TenderPath)
        'End If

        'If Not Directory.Exists(ReturnPath) Then
        '    Directory.CreateDirectory(ReturnPath)
        'End If



        ProcessDump = System.Windows.Forms.Application.StartupPath & "\ProcessPath"

        If as_programname <> "" Then
            ProcessDump += "\" + as_programname
        End If

        If Not Directory.Exists(ProcessDump) Then
            Directory.CreateDirectory(ProcessDump)
        End If




        ErrorFilePath = ErrorDump + "\" + Today + ".txt"
        SuccessFilePath = SuccessDump + "\" + Today + ".txt"
        ProcessPath = ProcessDump + "\" + Today + ".txt"

        'FileDump += "\" + Today


        'If Not Directory.Exists(FileDump) Then
        '    Directory.CreateDirectory(FileDump)
        'End If



        If Not File.Exists(ErrorFilePath) Then

            Dim ErrorLog As New FileInfo(ErrorFilePath)
            Dim StreamWriter As StreamWriter = ErrorLog.CreateText()

            StreamWriter.WriteLine("Error Log")
            StreamWriter.Close()
        End If
        If Not File.Exists(SuccessFilePath) Then

            Dim SuccessLog As New FileInfo(SuccessFilePath)
            Dim StreamWriter As StreamWriter = SuccessLog.CreateText()

            StreamWriter.WriteLine("Success Log")
            StreamWriter.Close()
        End If

        If Not File.Exists(ProcessPath) Then

            Dim ProcessLog As New FileInfo(ProcessPath)
            Dim StreamWriter As StreamWriter = ProcessLog.CreateText()

            StreamWriter.WriteLine("Process")
            StreamWriter.Close()
        End If

        GC.Collect()
    End Sub


    Public Sub ErrorAppendNew(ByVal as_message As String, Optional ByVal as_programname As String = "")
        If as_programname = "" Then as_programname = ProgramName


        Dim Log As FileInfo = New FileInfo(ErrorFilePath)
        Dim StreamWriter As StreamWriter = Log.AppendText()

        Dim ndate As DateTime = DateTime.Now


        StreamWriter.WriteLine(ndate.ToString() & vbTab & as_message)
        StreamWriter.Close()


        'If pathtoload.Contains("Error") Then
        '    Dim nrow As DataRow = dtProcess.NewRow

        '    nrow("DateAndTime") = ndate.ToString()
        '    nrow("Detail") = as_message

        '    dtProcess.Rows.Add(nrow)

        '    formgridview.Refresh()
        'End If

    End Sub


    Public Sub SuccessAppendNew(ByVal as_message As String, Optional ByVal as_programname As String = "")
        If as_programname = "" Then as_programname = ProgramName


        Dim Log As FileInfo = New FileInfo(SuccessFilePath)
        Dim StreamWriter As StreamWriter = Log.AppendText()

        Dim ndate As DateTime = DateTime.Now


        StreamWriter.WriteLine(ndate.ToString() & vbTab & as_message)
        StreamWriter.Close()


        'If pathtoload.Contains("Success") Then
        '    Dim nrow As DataRow = dtProcess.NewRow

        '    nrow("DateAndTime") = ndate.ToString()
        '    nrow("Detail") = as_message

        '    dtProcess.Rows.Add(nrow)
        '    formgridview.Refresh()
        'End If


    End Sub

    Public Sub ProcessAppendNew(ByVal as_message As String, Optional ByVal as_programname As String = "")
        If as_programname = "" Then as_programname = ProgramName


        Dim Log As FileInfo = New FileInfo(ProcessPath)
        Dim StreamWriter As StreamWriter = Log.AppendText()

        Dim ndate As DateTime = DateTime.Now


        StreamWriter.WriteLine(ndate.ToString() & vbTab & as_message)
        StreamWriter.Close()

        'If pathtoload.Contains("Process") Then
        '    Dim nrow As DataRow = dtProcess.NewRow

        '    nrow("DateAndTime") = ndate.ToString()
        '    nrow("Detail") = as_message

        '    dtProcess.Rows.Add(nrow)
        '    formgridview.Refresh()
        'End If
    End Sub


    Public Function GenerateSQLConnectionString(ByVal as_databasename As String) As String
        Return "Data Source=" + SQLServer + "; Initial Catalog=" + as_databasename + "; User ID=" + SQLUser + "; Password=" + SQLPassword + "; MultipleActiveResultSets=True;"
    End Function


    Public Function b1fs(ByVal as_sql As String) As String
        Dim actionReturn As String = ""
        Try

            Dim oRs As SAPbobsCOM.Recordset
            oRs = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRs.DoQuery(as_sql)
            If oRs.RecordCount > 0 Then
                oRs.MoveFirst()
                actionReturn = oRs.Fields.Item(0).Value.ToString()
            End If

            oRs = Nothing
            GC.Collect()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return actionReturn
    End Function

    Public Function FormattedSearch(ByVal sql As String, Optional ByVal as_databasename As String = "") As String
        Dim ls_return_value As String = ""
        Dim lsql_connection As New SqlConnection()
        Dim lsql_command As New SqlCommand
        Dim lsql_reader As SqlDataReader

        If as_databasename = "" Then
            as_databasename = SQLDBForQuery
        End If

        Try
            lsql_connection = New SqlConnection(GenerateSQLConnectionString(as_databasename))
            lsql_connection.Open()
            lsql_command = New SqlCommand(sql, lsql_connection)
            lsql_command.CommandTimeout = 10000
            lsql_reader = lsql_command.ExecuteReader()

            If lsql_reader.HasRows Then
                lsql_reader.Read()
                ls_return_value = lsql_reader.Item(0).ToString()
            End If

            lsql_reader.Close()
            lsql_command.Dispose()
            lsql_connection.Close()
            GC.Collect()
        Catch ex As Exception
            ls_return_value = "Error! ErrDesc: " + ex.Message
        End Try

        Return ls_return_value
    End Function


    Public Function FormattedSearchDTManual(ByVal sql As String, Optional ByVal as_databasename As String = "") As DataTable
        Dim ldt_return_table As New DataTable
        Dim lsql_connection As New SqlConnection()
        Dim lsql_command As New SqlCommand
        Dim lsql_reader As SqlDataReader

        If as_databasename = "" Then
            as_databasename = SQLDBForQuery
        End If

        Try
            lsql_connection = New SqlConnection(GenerateSQLConnectionString(as_databasename))
            lsql_connection.Open()
            lsql_command.CommandTimeout = 10000
            lsql_command = New SqlCommand(sql, lsql_connection)
            lsql_reader = lsql_command.ExecuteReader()

            If lsql_reader.HasRows Then

                'Dim columns = New List(Of String)()

                For i As Integer = 0 To lsql_reader.FieldCount - 1
                    ldt_return_table.Columns.Add(lsql_reader.GetName(i))
                Next

                While lsql_reader.Read
                    Dim row As DataRow = ldt_return_table.NewRow

                    For i As Integer = 0 To lsql_reader.FieldCount - 1

                        If Not IsDBNull(lsql_reader.GetValue(i)) Then

                            row(lsql_reader.GetName(i)) = lsql_reader.GetValue(i)
                        Else
                            row(lsql_reader.GetName(i)) = ""
                        End If
                    Next

                    ldt_return_table.Rows.Add(row)
                End While

            End If

            lsql_reader.Close()
            lsql_command.Dispose()
            lsql_connection.Close()
            GC.Collect()
            lsql_reader = Nothing
            lsql_command = Nothing
            lsql_connection = Nothing
            GC.Collect()
        Catch ex As Exception

            ' MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Throw New Exception(ex.Message)
            ldt_return_table = Nothing
            'MessageBox.Show("FMS DT: " + ex.Message, gs_programname, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return ldt_return_table
    End Function



    Public Function ExecuteNonQuery(ByVal sql As String, Optional ByVal as_databasename As String = "") As String
        Dim ldt_return_table As New DataTable
        Dim lsql_connection As New SqlConnection()
        Dim lsql_command As New SqlCommand
        'Dim lsql_reader As SqlDataReader

        If as_databasename = "" Then
            as_databasename = SQLDBForQuery
        End If

        Try
            lsql_connection = New SqlConnection(GenerateSQLConnectionString(as_databasename))
            lsql_connection.Open()
            lsql_command.CommandTimeout = 10000
            lsql_command = New SqlCommand(sql, lsql_connection)

            lsql_command.ExecuteNonQuery()


            lsql_command.Dispose()
            lsql_connection.Close()
            GC.Collect()

            lsql_command = Nothing
            lsql_connection = Nothing
            GC.Collect()

        Catch ex As Exception
            Return ex.Message
        End Try

        Return ""
    End Function

    Public Function ExecuteReader(ByVal sql As String, Optional ByVal as_databasename As String = "") As DataTable
        Dim ldt_return_table As New DataTable
        Dim lsql_connection As New SqlConnection()
        Dim lsql_command As New SqlCommand
        Dim lsql_reader As SqlDataReader

        If as_databasename = "" Then
            as_databasename = SQLDBForQuery
        End If

        Try
            lsql_connection = New SqlConnection(GenerateSQLConnectionString(as_databasename))
            lsql_connection.Open()
            lsql_command.CommandTimeout = 10000
            lsql_command = New SqlCommand(sql, lsql_connection)
            lsql_reader = lsql_command.ExecuteReader()

            If lsql_reader.HasRows Then

                'Dim columns = New List(Of String)()

                'For i As Integer = 0 To lsql_reader.FieldCount - 1
                '    ldt_return_table.Columns.Add(lsql_reader.GetName(i))
                'Next

                'While lsql_reader.Read
                '    Dim row As DataRow = ldt_return_table.NewRow

                '    For i As Integer = 0 To lsql_reader.FieldCount - 1

                '        If Not IsDBNull(lsql_reader.GetValue(i)) Then

                '            row(lsql_reader.GetName(i)) = lsql_reader.GetValue(i)
                '        Else
                '            row(lsql_reader.GetName(i)) = ""
                '        End If
                '    Next

                '    ldt_return_table.Rows.Add(row)
                'End While

                ldt_return_table.Load(lsql_reader)

            End If

            lsql_reader.Close()
            lsql_command.Dispose()
            lsql_connection.Close()
            GC.Collect()
            lsql_reader = Nothing
            lsql_command = Nothing
            lsql_connection = Nothing
            GC.Collect()
        Catch ex As Exception
            ' MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Throw New Exception(ex.Message)
            ldt_return_table = Nothing
        End Try
        Return ldt_return_table
    End Function
    Private Sub TransferFile(ByVal as_destination As String, ByVal as_filename As String, ByVal as_filepath As String)
        Dim ls_filename As String
        Dim ls_filenamepart() As String
        Dim ls_addedtofilename As String
        Try

            ls_filename = as_filename
            'Check if file exists 
            If File.Exists(as_destination + "\\" + as_filename) Then
                'Split by . the file name
                If ls_filename.Contains(".") Then
                    ls_filenamepart = ls_filename.Split(".")

                    'Remove if the filename has already a timestamp.
                    If (ls_filenamepart(0).Contains("_T")) Then
                        Dim ls_removetimestamp() As String = ls_filenamepart(0).Split("_T")
                        ls_filenamepart(0) = ls_removetimestamp(0)
                    End If

                    ls_addedtofilename = DateTime.Now.Hour.ToString(+DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + DateTime.Now.Millisecond.ToString())
                    ls_filename = ls_filenamepart(0) + "_T" + ls_addedtofilename + "." + ls_filenamepart(1)
                End If
            End If
            File.Move(as_filepath + "\\" + as_filename, as_destination + "\\" + ls_filename)
        Catch ex As Exception
            Throw New Exception(ex.Message + " FileName: " + as_filename)
        End Try
    End Sub


    Public Function FileTransfer(ByVal srcFile As String, ByVal desFile As String) As Boolean

        ' MessageBox.Show(srcFile + " " + desFile)
        ErrMsg = ""
        Dim actionReturn As Boolean = True
        Try
            Dim TempPath As String = Path.GetTempPath()
            If Not File.Exists(desFile) Then
                File.Copy(srcFile, desFile)
            Else
                File.Delete(desFile)
                File.Copy(srcFile, desFile)
            End If

        Catch ex As Exception
            actionReturn = False
            ErrMsg = ex.Message
        End Try
        Return actionReturn
    End Function

    Public Function FileIsDone(ByVal path As String) As Boolean
        Try

            Using File.Open(path, FileMode.Open, FileAccess.Read, FileShare.None)
            End Using

        Catch ex As Exception
            'ErrMsg = ex.Message
            GC.Collect()
            Return False
        End Try

        Return True
    End Function


    Public Sub ReleaseObject(ByVal as_obj)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(as_obj)
        Catch ex As Exception
            Throw New Exception("Releasing Object Error: " + ex.Message)
        End Try
    End Sub

End Module

'Public Module MyExtensions
'    <Extension()>
'    Public Sub Add(Of T)(ByRef arr As T(), item As T)
'        Array.Resize(arr, arr.Length + 1)
'        arr(arr.Length - 1) = item
'    End Sub

'    Public Function B1AlphaNumeric(Of T)() As Long
'        Return SAPbobsCOM.BoFieldTypes.db_Alpha
'    End Function

'End Module
