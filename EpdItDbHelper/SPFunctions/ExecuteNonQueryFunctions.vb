Imports System.Data.SqlClient

Partial Public Class DB

#Region " Stored Procedures that run commands (ExecuteNonQuery) "

    ' These functions call Oracle Stored Procedures using IN and/or OUT parameters.
    ' If successful, any OUT parameters are available to the calling procedure as 
    ' returned by the Oracle database.

    ''' <summary>
    ''' Executes a Stored Procedure on the database.
    ''' </summary>
    ''' <param name="spName">The name of the Stored Procedure to execute.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <returns>True if the Stored Procedure ran successfully. Otherwise, false.</returns>
    Public Function SPRunCommand(ByVal spName As String, Optional ByRef parameter As SqlParameter = Nothing) As Boolean
        Dim parameterArray As SqlParameter() = {parameter}
        Dim result As Boolean = SPRunCommand(spName, parameterArray)
        If result Then
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
    Public Function SPRunCommand(ByVal spName As String, ByRef parameterArray As SqlParameter()) As Boolean
        Dim success As Boolean = True

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(spName, connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddRange(parameterArray)
                Try
                    command.Connection.Open()
                    command.ExecuteNonQuery()
                    command.Connection.Close()
                    command.Parameters.CopyTo(parameterArray, 0)
                Catch ee As SqlException
                    success = False
                End Try
            End Using
        End Using

        Return success
    End Function

#End Region

End Class
