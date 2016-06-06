Imports System.Data.SqlClient

' These functions call Stored Procedures using IN and/or OUT parameters.
' If successful, any OUT parameters are available to the calling procedure as 
' returned by the database.

Partial Public Class DBHelper

    ''' <summary>
    ''' Executes a Stored Procedure on the database.
    ''' </summary>
    ''' <param name="spName">The name of the Stored Procedure to execute.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="rowsAffected">For UPDATE, INSERT, and DELETE statements, stores the number of rows affected by the command.</param>
    ''' <returns>True if the Stored Procedure ran successfully. Otherwise, false.</returns>
    Public Function SPRunCommand(spName As String,
                                 Optional ByRef parameter As SqlParameter = Nothing,
                                 Optional ByRef rowsAffected As Integer = 0
                                 ) As Boolean
        rowsAffected = 0
        Dim parameterArray As SqlParameter() = Nothing
        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If
        Dim result As Boolean = SPRunCommand(spName, parameterArray, rowsAffected)
        If result AndAlso parameterArray IsNot Nothing Then
            parameter = parameterArray(0)
        End If
        Return result
    End Function

    ''' <summary>
    ''' Executes a Stored Procedure on the database.
    ''' </summary>
    ''' <param name="spName">The name of the Stored Procedure to execute.</param>
    ''' <param name="parameterArray">An SqlParameter array to send.</param>
    ''' <returns>True if the Stored Procedure ran successfully. Otherwise, false.</returns>
    Public Function SPRunCommand(spName As String,
                                 ByRef parameterArray As SqlParameter(),
                                 Optional ByRef rowsAffected As Integer = 0
                                 ) As Boolean
        Dim success As Boolean = False

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(spName, connection)
                command.CommandType = CommandType.StoredProcedure
                If parameterArray IsNot Nothing Then
                    command.Parameters.AddRange(parameterArray)
                End If
                command.Connection.Open()
                rowsAffected = command.ExecuteNonQuery()
                command.Connection.Close()
                If parameterArray IsNot Nothing Then
                    command.Parameters.CopyTo(parameterArray, 0)
                End If
                command.Parameters.Clear()
                success = True
            End Using
        End Using

        Return success
    End Function

End Class
