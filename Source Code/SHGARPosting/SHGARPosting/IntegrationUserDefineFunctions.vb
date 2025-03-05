Imports System.IO
Imports System.Configuration
Imports System.Net.Mail
Imports System.Net
Imports System.Text
Imports System.Data.Sql
Imports System.Data.SqlClient


Module IntegrationUserDefineFunctions

    Public Function isDBSQLExists(ByVal as_databasename As String) As Boolean

        Dim exists As String = FormattedSearch("SELECT name FROM master.dbo.sysdatabases WHERE name = N'" + as_databasename + "'")
        If exists = "" Then
            Return False
        End If
        Return True
    End Function

    Public Function isDBSQLTableExists(ByVal as_tablename As String) As Boolean

        Dim exists As String = FormattedSearch("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = N'" + as_tablename + "'")
        If exists = "" Then
            Return False
        End If
        Return True
    End Function

    Public Function isDBSQLTableColumnExists(ByVal as_tablename As String, ByVal as_columnname As String) As Boolean

        as_columnname = "U_" + as_columnname

        Dim exists As String = FormattedSearch("SELECT column_id FROM sys.columns  WHERE Name = N'" + as_columnname + "' AND Object_ID = Object_ID(N'" + as_tablename + "')")
        If exists = "" Then
            Return False
        End If
        Return True
    End Function

    Public Function createDBSQL(ByVal as_databasename As String) As String

        Dim ldt_return_table As New DataTable
        Dim lsql_connection As New SqlConnection()
        Dim lsql_command As New SqlCommand
        'Dim lsql_reader As SqlDataReader

        Try
            Threading.Thread.Sleep(1000)
            lsql_connection = New SqlConnection(GenerateSQLConnectionString(SQLDBForQuery))
            lsql_connection.Open()
            lsql_command = New SqlCommand("CREATE DATABASE " + as_databasename, lsql_connection)
            lsql_command.ExecuteNonQuery()
            lsql_command.Dispose()
            lsql_connection.Close()
            GC.Collect()
        Catch ex As Exception
            Return ex.Message
        End Try

        Return ""
    End Function

    Public Function createUDTSQL(ByVal as_tablename As String) As String
        Dim ldt_return_table As New DataTable
        Dim lsql_connection As New SqlConnection()
        Dim lsql_command As New SqlCommand
        'Dim lsql_reader As SqlDataReader

        Try
            Threading.Thread.Sleep(1000)
            lsql_connection = New SqlConnection(SQLConnectionString)
            lsql_connection.Open()
            SQLQuery = "CREATE TABLE [dbo].[" + as_tablename + "] " +
                                "([DocEntry] [int] IDENTITY(1,1) Not NULL, " +
                                "[DocNum] [varchar] (50) Not NULL, " +
                                "[LineId] [int] Not NULL, " +
                                "[Object] [varchar] (10) Not NULL, " +
                                "[Status] [varchar] (2) Not NULL DEFAULT N'A', " +
                                "[CreateDate] [datetime] Not NULL, " +
                                "[CreatedBy] [varchar] (50) Not NULL " +
                                ",[UpdateDate] [datetime] Not NULL, " +
                                "[UpdatedBy] [varchar] (50) Not NULL, " +
                                "CONSTRAINT [PK_" + as_tablename + "] PRIMARY KEY CLUSTERED ([DocEntry] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON)) ON [PRIMARY]"
            lsql_command = New SqlCommand(SQLQuery, lsql_connection)

            lsql_command.ExecuteNonQuery()

            lsql_command.Dispose()
            lsql_connection.Close()
            GC.Collect()

        Catch ex As Exception
            Return ex.Message
        End Try

        Return ""
    End Function



    Public Function createUDFSQL(ByVal as_tablename As String, ByVal as_columnname As String, ByVal as_datatype As String, ByVal as_defaultval As String, Optional ByVal as_datalength As String = "0") As String
        Dim ldt_return_table As New DataTable
        Dim lsql_connection As New SqlConnection()
        Dim lsql_command As New SqlCommand
        Try
            Threading.Thread.Sleep(1000)
            lsql_connection = New SqlConnection(SQLConnectionString)
            lsql_connection.Open()

            Dim length As String = ""
            If as_datalength <> "0" Then length = "(" + as_datalength.ToString() + ")"

            SQLQuery = "ALTER TABLE [dbo].[" + as_tablename + "] ADD U_" + as_columnname + " [" + as_datatype + "] " + length

            If as_defaultval <> "" Then
                SQLQuery += " DEFAULT N'" + as_defaultval + "'"
            End If

            lsql_command = New SqlCommand(SQLQuery, lsql_connection)

            lsql_command.ExecuteNonQuery()

            lsql_command.Dispose()
            lsql_connection.Close()
            GC.Collect()
        Catch ex As Exception
            Return ex.Message
        End Try
        Return ""
    End Function



    ''B1

    Public Function B1Connect(ByVal as_databasename As String) As SAPbobsCOM.Company

        Try
            If Not ReConnect Then
                Return SAPB1Company
            Else
                If SAPB1Company.Connected Then
                    ProcessAppendNew("Re-Connecting to SAP B1 Database")
                    SAPB1Company.Disconnect()
                    ReleaseObject(SAPB1Company)
                End If
            End If


            Dim DBType As String = ""

            SAPB1Company = New SAPbobsCOM.Company
            ProcessAppendNew("Connecting to " + as_databasename + " via DI API...")

            If Not SAPB1Company.Connected Then
                Select Case SQLServerType
                    Case "MSSQL2000"
                        DBType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL
                    Case "MSSQL2005"
                        DBType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
                    Case "MSSQL2008"
                        DBType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
                    Case "MSSQL2012"
                        DBType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
                    Case "HANADB"
                        DBType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
                End Select

                SAPB1Company.Server = ServerName
                SAPB1Company.DbServerType = DBType
                SAPB1Company.UseTrusted = SAPB1UseTrusted
                SAPB1Company.DbUserName = SQLUser
                SAPB1Company.DbPassword = SQLPassword

                SAPB1Company.CompanyDB = as_databasename
                SAPB1Company.UserName = UserId
                SAPB1Company.Password = UserPassword


                SAPB1Company.language = SAPbobsCOM.BoSuppLangs.ln_English

                If Not SAPB1Company.Connected Then
                    SAPB1ReturnCode = SAPB1Company.Connect()
                End If

                If SAPB1ReturnCode <> 0 Then

                    SAPB1Company.GetLastError(ErrNumber, ErrMsg)
                    Dim string_test As String = ""
                    'string_test = String.Format(" Server:{0}, DBType: {1}, UserTrusted: {2}, DBUserId: {3}, DBUserPassword: {4}, CompanyDB: {5}, UserId: {6}, Password: {7}", ServerName, DBType, SAPB1UseTrusted.ToString, SQLUser, SQLPassword, as_databasename, UserId, UserPassword)
                    Throw New Exception(ErrNumber.ToString() + " " + ErrMsg + string_test)
                    Return Nothing
                Else
                    ReConnect = False
                End If
            End If

            GC.Collect()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try


        Return SAPB1Company

    End Function

    Public Function ConnectToCompany() As Boolean


        SAPB1Company = New SAPbobsCOM.Company

        Dim DBType As SAPbobsCOM.BoDataServerTypes
        Dim lretCode, lErrCode As Integer
        Dim sErrMsg, Server, DbServerType, UseTrusted, DbUserName, DbPassword, SBOCompanyDB, SBOUserName, SBOPassword, ReprocessTime, DBName As String
        Dim sr As StreamReader
        sErrMsg = ""
        Try
            FilePath = System.Windows.Forms.Application.StartupPath + "\\FTB1IntegrationConnectionSettings.ini"
            sr = New StreamReader(FilePath)
        Catch ex As Exception
            sr = Nothing
            Throw New Exception("Connecting to Company: FTB1IntegrationConnectionSettings.ini was not found or not configured properly.")

            Return False
        End Try


        Server = sr.ReadLine()
        DbServerType = sr.ReadLine()
        UseTrusted = sr.ReadLine()
        DBName = sr.ReadLine()
        SBOCompanyDB = sr.ReadLine()
        DbUserName = sr.ReadLine()
        DbPassword = sr.ReadLine()

        ' SBOUserName = sr.ReadLine()
        'SBOPassword = sr.ReadLine()
        'ReprocessTime = sr.ReadLine()
        ProgramName = sr.ReadLine()
        'FilePath = sr.ReadLine()


        Server = Server.Substring(Server.IndexOf("=") + 1)
        DbServerType = DbServerType.Substring(DbServerType.IndexOf("=") + 1)
        UseTrusted = UseTrusted.Substring(UseTrusted.IndexOf("=") + 1)
        DBName = DBName.Substring(DBName.IndexOf("=") + 1)
        DbUserName = DbUserName.Substring(DbUserName.IndexOf("=") + 1)
        DbPassword = DbPassword.Substring(DbPassword.IndexOf("=") + 1)
        SBOCompanyDB = SBOCompanyDB.Substring(SBOCompanyDB.IndexOf("=") + 1)
        'SBOUserName = SBOUserName.Substring(SBOUserName.IndexOf("=") + 1)
        ' SBOPassword = SBOPassword.Substring(SBOPassword.IndexOf("=") + 1)
        'ReprocessTime = ReprocessTime.Substring(ReprocessTime.IndexOf("=") + 1)
        ProgramName = ProgramName.Substring(ProgramName.IndexOf("=") + 1)
        'FilePath = FilePath.Substring(FilePath.IndexOf("=") + 1)


        If Server = "[server_name]" Or String.IsNullOrEmpty(Server) = True Then
            Throw New Exception("Connecting to Company: Please configure the connection setting properly.")
        End If



        'If String.IsNullOrEmpty(FilePath) = True Then
        '    Throw New Exception("Connecting to Company: FilePath must not be blank please configure the connection setting properly.")
        'ElseIf Directory.Exists(FilePath) = False Then
        '    Throw New Exception("Connecting to Company: FilePath [" + FilePath + "] does not exists please configure the connection setting properly.")
        'End If

        'li_minresync = Convert.ToInt32(ReprocessTime)

        If Not SAPB1Company.Connected Then
            Select Case DbServerType
                Case "SQL Server 2000"
                    DBType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL
                Case "SQL Server 2005"
                    DBType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
                Case "SQL Server 2008"
                    DBType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
                Case "SQL Server 2012"
                    DBType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            End Select

            SAPB1Company.Server = Server
            SAPB1Company.DbServerType = DBType
            SAPB1Company.UseTrusted = UseTrusted
            SAPB1Company.DbUserName = DbUserName
            SAPB1Company.DbPassword = DbPassword

            SAPB1Company.CompanyDB = SBOCompanyDB
            SAPB1Company.UserName = UserId
            SAPB1Company.Password = UserPassword


            SAPB1Company.language = SAPbobsCOM.BoSuppLangs.ln_English

            If Not SAPB1Company.Connected Then
                lretCode = SAPB1Company.Connect()
            End If

            If lretCode <> 0 Then
                SAPB1Company.GetLastError(lErrCode, sErrMsg)
                Throw New Exception(lErrCode.ToString() + " " + sErrMsg)
                Return False
            Else

                SQLDBName = SBOCompanyDB

                SQLServer = Server
                SQLUser = DbUserName
                SQLPassword = DbPassword
                SQLDBName = DBName

                If SQLDBExists Then
                    SQLForCheckingDBName = DBName
                Else
                    SQLForCheckingDBName = SBOCompanyDB
                End If

                SQLConnectionString = "Data Source=" + SQLServer + "; Initial Catalog=" + SQLForCheckingDBName + "; User ID=" + SQLUser + "; Password=" + SQLPassword + "; MultipleActiveResultSets=True;"


            End If
        End If
        sr.Dispose()
        sr.Close()
        GC.Collect()
        Return True
    End Function


    Public Function isUDTExist(ByVal as_tablename As String) As Boolean

        Dim oRecordSet As SAPbobsCOM.Recordset

        oRecordSet = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            oRecordSet.DoQuery("select ""TableName"" from OUTB where ""TableName"" ='" + as_tablename + "'")
            If oRecordSet.RecordCount = 0 Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                oRecordSet = Nothing
                GC.Collect()
                Return False
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()
            Return True

        Catch ex As Exception
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()
            Return False
        End Try
    End Function


    Public Function isUDFExist(ByVal as_tablename As String, ByVal as_fieldname As String) As Boolean

        Dim oRecordSet As SAPbobsCOM.Recordset

        oRecordSet = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try


            oRecordSet.DoQuery("Select ""AliasID"" from CUFD where ""TableID"" ='" + as_tablename + "' and ""AliasID"" ='" + as_fieldname + "'")
            If oRecordSet.RecordCount = 0 Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                oRecordSet = Nothing
                GC.Collect()
                Return False
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()
            Return True

        Catch ex As Exception
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()
            Return False
        End Try
    End Function


    Public Function createUDF(ByVal as_tablename As String, ByVal as_name As String, ByVal as_description As String, ByVal al_type As Long, ByVal ai_size As Integer, ByVal as_default As String, ByVal as_options As String) As Boolean

        SAPRetVal = 0
        Dim UserFieldsMD As SAPbobsCOM.UserFieldsMD

        Dim li_index As Integer

        Try
            UserFieldsMD = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            UserFieldsMD.Description = as_description
            UserFieldsMD.Name = as_name
            UserFieldsMD.TableName = as_tablename

            Select Case al_type
                Case 0
                    UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
                    If ai_size > 0 Then
                        UserFieldsMD.EditSize = ai_size
                    End If

                    If Not as_default = "" Then
                        UserFieldsMD.DefaultValue = as_default
                    End If

                Case 1
                    UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Memo

                Case 2
                    UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Numeric
                    If (ai_size > 0) Then
                        UserFieldsMD.EditSize = ai_size
                    End If

                    If Not as_default = "" Then
                        UserFieldsMD.DefaultValue = as_default
                    End If

                Case 3
                    UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
                    If Not as_default = "" Then
                        UserFieldsMD.DefaultValue = as_default
                    End If

                Case 4
                    UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
                    If Not as_default = "" Then
                        UserFieldsMD.DefaultValue = as_default
                    End If

                Case 77
                    UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
                    UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Measurement
                    If Not as_default = "" Then
                        UserFieldsMD.DefaultValue = as_default
                    End If

                Case 37
                    UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
                    UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Percentage
                    If Not as_default = "" Then
                        UserFieldsMD.DefaultValue = as_default
                    End If

                Case 80
                    UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
                    UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
                    If Not as_default = "" Then
                        UserFieldsMD.DefaultValue = as_default
                    End If

                Case 81

                    UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
                    UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Quantity
                    If Not as_default = "" Then
                        UserFieldsMD.DefaultValue = as_default
                    End If

                Case 82
                    UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
                    UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Rate
                    If Not as_default = "" Then
                        UserFieldsMD.DefaultValue = as_default
                    End If

                Case 83
                    UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
                    UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Sum
                    If Not as_default = "" Then
                        UserFieldsMD.DefaultValue = as_default
                    End If

                Case 84
                    UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
                    UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Time
                    If Not as_default = "" Then
                        UserFieldsMD.DefaultValue = as_default
                    End If

                Case 35
                    'SAPbobsCOM.BoFldSubTypes.st_Phone

                Case 63
                    'SAPbobsCOM.BoFldSubTypes.st_Address

                Case 73
                    'SAPbobsCOM.BoFldSubTypes.st_Image

                Case 66
                    'SAPbobsCOM.BoFldSubTypes.st_Link

                Case Else

            End Select

            Dim sOptions() As String
            Dim sValue() As String

            li_index = 0
            If as_options.Trim <> "" Then
                sOptions = as_options.Split(",")

                For i_cnt = 0 To sOptions.Length - 1
                    sValue = sOptions(i_cnt).Split("-")
                    If sValue(0).Trim <> "" Then
                        If (li_index > 0) Then
                            UserFieldsMD.ValidValues.Add()
                            UserFieldsMD.ValidValues.SetCurrentLine(li_index)
                        End If

                        UserFieldsMD.ValidValues.Value = sValue(0).Trim
                        UserFieldsMD.ValidValues.Description = sValue(1).Trim

                        li_index = li_index + 1
                    End If
                Next
            End If

            li_index = 0
            SAPRetVal = UserFieldsMD.Add()
            If Not SAPRetVal = 0 Then
                ErrNumber = SAPB1Company.GetLastErrorCode()
                ErrMsg = SAPB1Company.GetLastErrorDescription()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(UserFieldsMD)
                UserFieldsMD = Nothing
                GC.Collect()
                Throw New Exception("Error Code: " + ErrNumber.ToString() + " Error Description: " + ErrMsg)
                'Return False
            Else
                System.Runtime.InteropServices.Marshal.ReleaseComObject(UserFieldsMD)
                UserFieldsMD = Nothing
                GC.Collect()
                'Return True
            End If

            UserFieldsMD = Nothing


        Catch ex As Exception
            'MessageBox.Show(ex.Message, Me.Text)
            ErrMsg = ex.Message
            UserFieldsMD = Nothing
            GC.Collect()
            Return False
        End Try
        Return True
    End Function




    Public Function createUDT(ByVal as_tablename As String, ByVal as_tabledescription As String, ByVal aole_tabletype As SAPbobsCOM.BoUTBTableType) As Boolean

        Dim UserTablesMD As SAPbobsCOM.UserTablesMD

        SAPRetVal = 0

        Try
            UserTablesMD = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)

            If UserTablesMD.GetByKey(as_tablename) = False Then

                UserTablesMD.TableName = as_tablename
                UserTablesMD.TableDescription = as_tabledescription
                UserTablesMD.TableType = aole_tabletype

                SAPRetVal = UserTablesMD.Add()

                If Not SAPRetVal = 0 Then

                    ErrNumber = SAPB1Company.GetLastErrorCode()
                    ErrMsg = SAPB1Company.GetLastErrorDescription()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(UserTablesMD)
                    UserTablesMD = Nothing
                    GC.Collect()

                    Throw New Exception("Error Code: " + ErrNumber.ToString() + " Error Description: " + ErrMsg)
                Else

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(UserTablesMD)
                    UserTablesMD = Nothing
                    GC.Collect()

                    'Return True
                End If

            End If

            'System.Runtime.InteropServices.Marshal.ReleaseComObject(UserTablesMD)
            'UserTablesMD = Nothing
            'GC.Collect()

            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            UserTablesMD = Nothing
            GC.Collect()
            Return False
        End Try


        Return True
    End Function


    Public Function createUDO(ByVal as_code As String, ByVal as_name As String, ByVal al_objecttype As SAPbobsCOM.BoUDOObjType, ByVal as_tablename As String, ByVal as_childtables As String, ByVal as_findcolumns As String, ByVal ab_manageseries As Boolean, ByVal ab_cancel As Boolean, ByVal ab_close As Boolean, ByVal ab_delete As Boolean) As Boolean

        Try

            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
            Dim ErrNumber As Long
            Dim sChild() As String
            Dim sValue As String
            Dim li_index As Integer
            oUserObjectMD = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

            If oUserObjectMD.GetByKey(as_code) = False Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.FindColumns.ColumnAlias = "Code"
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.LogTableName = ""
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ExtensionName = ""
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.Code = as_code
                oUserObjectMD.Name = as_name
                oUserObjectMD.ObjectType = al_objecttype
                oUserObjectMD.TableName = as_tablename


                If as_childtables.Trim <> "" Then
                    sChild = as_childtables.Split(",")

                    For i_cnt = 0 To sChild.Length - 1
                        sValue = sChild(i_cnt).Trim
                        If sValue <> "" Then
                            If (li_index > 0) Then
                                oUserObjectMD.ChildTables.Add()
                            End If
                            oUserObjectMD.ChildTables.SetCurrentLine(li_index)
                            oUserObjectMD.ChildTables.TableName = sValue

                            li_index = li_index + 1
                        End If
                    Next
                End If

                If Not oUserObjectMD.Add() = 0 Then
                    ErrNumber = SAPB1Company.GetLastErrorCode()
                    ErrMsg = SAPB1Company.GetLastErrorDescription()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
                    oUserObjectMD = Nothing
                    GC.Collect()
                    Return False
                End If

                GC.Collect()
                Return True

            End If

            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function
End Module
