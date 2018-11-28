Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Runtime.CompilerServices
Imports System.Reflection
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Linq
Imports System.Linq
Namespace SQLClients
    ''' <summary>
    ''' SQLClient Helper class for more methods on the SQL Server database usage
    ''' </summary>
    ''' <remarks></remarks>
    Public Module SQLClientHelper
        ''' <summary>
        ''' Checks if a sql database is reachable
        ''' </summary>
        ''' <param name="DBConnectionString">Connectionstring to database</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function PingDB(DBConnectionString As String) As Boolean
            Dim Conn As String = DBConnectionString : Dim Connection As New SqlConnection(Conn)
            Try
                Connection.Open()
                Return True
            Catch ex As Exception
                Throw ex
                Return False
            Finally
                If Connection.State = ConnectionState.Open Then
                    Connection.Close()
                End If
            End Try
        End Function

        ''' <summary>
        ''' Derives array of SQL parameter from the properties of an object
        ''' </summary>
        ''' <param name="RequiredObject">Object to be examined for available properties</param>
        ''' <param name="NeededKeys">Optional array of string to list the required properties. If not specified, all properties would be used</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSqlParameterFromProperties(ByVal RequiredObject As Object, Optional NeededKeys As String() = Nothing) As SqlParameter()
            Dim param As New ArrayList
            Dim _type As Type = RequiredObject.GetType()
            Dim properties() As PropertyInfo = _type.GetProperties()

            For Each _property As PropertyInfo In properties
                'Console.WriteLine("Name: " + _property.Name + ", Value: " + _property.GetValue(RequiredObject, Nothing)) 
                If NeededKeys IsNot Nothing AndAlso NeededKeys.Length > 0 Then
                    If NeededKeys.Select(Function(x) x.Contains(_property.Name)).Count() > 0 Then
                        param.Add(New SqlParameter(_property.Name, _property.GetValue(RequiredObject, Nothing))) ' filtered records 
                    End If
                Else
                    param.Add(New SqlParameter(_property.Name, _property.GetValue(RequiredObject, Nothing))) ' all records
                End If
            Next
            Return param.ToArray(GetType(SqlParameter))
        End Function

        ''' <summary>
        ''' Translates exception to friendly message
        ''' </summary>
        ''' <param name="ex">Exception variable</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function CustomError(ex As Exception) As String
            Return New CustomSQLException(ex).GetMessage
        End Function
    End Module

    ''' <summary>
    ''' A class to customize SQL server message from default 'throw exception' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Class CustomSQLException
        ''' <summary>
        ''' Represents the error code returned from stored procedure when entity could not be found.
        ''' </summary>
        ''' <remarks></remarks>
        Private Const SQL_ERROR_CODE_ENTITY_NOT_FOUND As Integer = 50001

        ''' <summary>
        ''' Represents the error code returned from stored procedure when entity to be updated has time mismatch
        ''' </summary>
        ''' <remarks></remarks>
        Private Const SQL_ERROR_CODE_TIME_MISMATCH As Integer = 50002

        ''' <summary>
        ''' Represents the error code returned from stored procedure when a persistence exception occurs
        ''' </summary>
        ''' <remarks></remarks> 
        Private Const SQL_ERROR_CODE_PERSISTENCE_ERROR As Integer = 50003

        ''' <summary>
        ''' Represents the error code returned when dead lock occurs
        ''' </summary>
        ''' <remarks></remarks>
        Private Const SQL_DEADLOCK_ERROR As Integer = 1205

        ''' <summary>
        ''' Represents the error code returned when timeout occurs
        ''' </summary>
        ''' <remarks></remarks>
        Private Const SQL_TIMEOUT_ERROR As Integer = -2

        Private Const DefaultSQLMessage As String = "Error found in database processing. Contact administrator"
        Private Message As String = ""
        Public Sub New(ex As Exception)
            If ex.InnerException Is GetType(SqlException) Then 'handle SQL Exception here
                Dim SQLExcep = CType(ex.InnerException, SqlException)
                Select Case SQLExcep.Number
                    Case SQL_ERROR_CODE_ENTITY_NOT_FOUND
                        Message = "Invalid database transaction was sent"
                    Case SQL_ERROR_CODE_PERSISTENCE_ERROR
                        Message = "Unending database process was detected"
                    Case SQL_ERROR_CODE_TIME_MISMATCH
                        Message = "Date sent is in incorrect format"
                    Case SQL_DEADLOCK_ERROR
                        Message = "The processing could not be completed due to unending transaction"
                    Case SQL_TIMEOUT_ERROR
                        Message = "Database server is unreachable"
                    Case Else
                        Message = DefaultSQLMessage
                End Select
            ElseIf ex.GetType Is GetType(OutOfMemoryException) Then 'handle local memory exception here
                Message = "No enough memory to complete process. Release more memory by closing some application"
            Else 'other exceptions
                Message = ex.Message
            End If
        End Sub

        Public Function GetMessage() As String
            Return Message
        End Function
    End Class

End Namespace
