Imports System.Data.Sql
Imports System.Data.SqlClient


Public Class SQLClass

    Private dbserver As String
    Private dbdatabasename As String
    Private dbuserid As String
    Private dbpassword As String

    Private sqlconnection As SqlConnection
    Private sqlcommand As SqlCommand
    Private sqltransaction As SqlTransaction
    Private sqlreader As SqlDataReader

    Private sqlerrormsg As String

    Private insertid As String

    Private stServerName As String
    Private stDatabaseUserId As String
    Private stDatabasePassword As String


    Private sqlconnectionstring As String


    Public Sub New(ByVal as_server As String, ByVal as_databasename As String, ByVal as_userid As String, ByVal as_password As String)
        dbserver = as_server
        dbdatabasename = as_databasename
        dbuserid = as_userid
        dbpassword = as_password

        sqlconnectionstring = "Data Source=" + as_server + "; Initial Catalog=" + as_databasename + "; User ID=" + as_userid + "; Password=" + as_password + "; MultipleActiveResultSets=True;"

        ConnectionOpen()
    End Sub

    Public Sub ConnectionOpen()
        Try
            sqlconnection = New SqlConnection(sqlconnectionstring)
            sqlconnection.Open()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub


    Public Function GetSqlError() As String
        Return sqlerrormsg
    End Function

    Public Function IsConnectionOpen() As Boolean
        sqlerrormsg = ""
        Try
            If sqlconnection.State = ConnectionState.Closed Then
                Return False
            End If
        Catch ex As Exception
            sqlerrormsg = ex.Message
            Return False
        End Try

        Return True
    End Function

    Public Sub ConnectionClose()
        Try
            sqlconnection.Close()
            GC.Collect()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub


    Public Sub BeginTransaction()
        Try
            sqltransaction = sqlconnection.BeginTransaction
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub Commit()
        Try
            sqltransaction.Commit()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub RollBack()

        Try
            sqltransaction.Rollback()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub


    Public Function FormattedSearch(ByVal as_sql As String) As String
        sqlerrormsg = ""
        Try

            sqlcommand = New SqlCommand(as_sql, sqlconnection)
            sqlcommand.CommandTimeout = 10
            sqlreader = sqlcommand.ExecuteReader()

            If sqlreader.HasRows Then
                sqlreader.Read()
                Return sqlreader.Item(0).ToString()
            End If

            GC.Collect()

        Catch ex As Exception
            sqlerrormsg = ex.Message
            Return "ERROR"
        End Try
        Return ""
    End Function


    Public Function ExecuteReaderToDataTable(ByVal as_sql As String) As DataTable
        Dim dt As DataTable = New DataTable
        sqlerrormsg = ""
        Try
            sqlcommand = New SqlCommand(as_sql, sqlconnection)
            sqlcommand.CommandTimeout = 10
            sqlreader = sqlcommand.ExecuteReader()

            If sqlreader.HasRows Then
                dt.Load(sqlreader)
            End If

            GC.Collect()

        Catch ex As Exception
            sqlerrormsg = ex.Message
            dt = Nothing
        End Try

        Return dt
    End Function


    Public Function ExecuteReader(ByVal as_sql As String) As SqlDataReader
        sqlerrormsg = ""
        Try
            sqlcommand = New SqlCommand(as_sql, sqlconnection)
            sqlcommand.CommandTimeout = 10
            sqlreader = sqlcommand.ExecuteReader()

            Return sqlreader


        Catch ex As Exception
            sqlerrormsg = ex.Message
            Return Nothing
        End Try

        Return Nothing
    End Function

    Public Function ExecuteNonQuery(ByVal as_sql As String) As Boolean
        sqlerrormsg = ""
        Try
            sqlcommand = New SqlCommand(as_sql, sqlconnection)
            sqlcommand.CommandTimeout = 10
            sqlcommand.ExecuteNonQuery()
            GC.Collect()
        Catch ex As Exception
            sqlerrormsg = ex.Message
            Return False
        End Try

        Return True
    End Function


    Public Function ExecuteTransactionalNonQuery(ByVal as_sql As String) As Boolean
        sqlerrormsg = ""
        Try
            sqlcommand = New SqlCommand(as_sql, sqlconnection, sqltransaction)
            sqlcommand.CommandTimeout = 10
            sqlcommand.ExecuteNonQuery()
            GC.Collect()
        Catch ex As Exception
            sqlerrormsg = ex.Message
            Return False
        End Try

        Return True
    End Function

End Class
