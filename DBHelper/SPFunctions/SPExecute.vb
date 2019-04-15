Imports System.Data.SqlClient
Imports EpdIt.DBUtilities

Partial Public Class DBHelper

    ''' <summary>
    ''' Executes a non-query stored procedure on the database.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameterArray">An array of SqlParameter values. The array may be modified by the stored produre if it includes output parameters.</param>
    ''' <param name="returnValue">Output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>The number of rows affected.</returns>
    Private Function SPExecuteNonQuery(spName As String, parameterArray As SqlParameter(), ByRef returnValue As Integer) As Integer
        If String.IsNullOrEmpty(spName) Then Throw New ArgumentException("The name of the stored procedure must be specified.", "spName")

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(spName, connection)

                ' Setup
                command.CommandType = CommandType.StoredProcedure

                If parameterArray IsNot Nothing AndAlso parameterArray.Count > 0 Then
                    DBNullifyParameters(parameterArray)
                    command.Parameters.AddRange(parameterArray)
                End If

                Dim returnParameter As SqlParameter = ReturnValueParameter()
                command.Parameters.Add(returnParameter)

                ' Run
                command.Connection.Open()
                Dim rowsAffected As Integer = command.ExecuteNonQuery()
                command.Connection.Close()

                ' Cleanup
                returnValue = returnParameter.Value
                command.Parameters.Remove(returnParameter)

                If parameterArray IsNot Nothing AndAlso parameterArray.Count > 0 Then
                    Dim newArray(command.Parameters.Count) As SqlParameter
                    command.Parameters.CopyTo(newArray, 0)
                    Array.Copy(newArray, parameterArray, parameterArray.Length)
                End If

                command.Parameters.Clear()

                ' Return
                Return rowsAffected

            End Using
        End Using
    End Function

    ''' <summary>
    ''' Retrieves a DataSet containing one or more DataTables selected from the database by calling a stored procedure.
    ''' (Adds the necessary columns and primary key information to complete the schema.)
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameterArray">An array of SqlParameter values. The array may be modified by the stored produre if it includes output parameters.</param>
    ''' <param name="returnValue">Output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>A DataSet.</returns>
    Private Function SPFillDataSet(spName As String, parameterArray As SqlParameter(), ByRef returnValue As Integer) As DataSet
        If String.IsNullOrEmpty(spName) Then Throw New ArgumentException("The name of the stored procedure must be specified.", "spName")

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(spName, connection)

                ' Setup
                command.CommandType = CommandType.StoredProcedure

                If parameterArray IsNot Nothing AndAlso parameterArray.Count > 0 Then
                    DBNullifyParameters(parameterArray)
                    command.Parameters.AddRange(parameterArray)
                End If

                Dim returnParameter As SqlParameter = ReturnValueParameter()
                command.Parameters.Add(returnParameter)

                ' Run
                Dim dataSet As New DataSet
                Using adapter As New SqlDataAdapter(command)
                    adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
                    adapter.Fill(dataSet)
                End Using

                ' Cleanup
                returnValue = returnParameter.Value
                command.Parameters.Remove(returnParameter)

                If parameterArray IsNot Nothing AndAlso parameterArray.Count > 0 Then
                    Dim newArray(command.Parameters.Count) As SqlParameter
                    command.Parameters.CopyTo(newArray, 0)
                    Array.Copy(newArray, parameterArray, parameterArray.Length)
                End If

                command.Parameters.Clear()

                ' Return
                Return dataSet

            End Using
        End Using
    End Function

    ''' <summary>
    ''' Retrieves a single value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameterArray">An array of SqlParameter values. The array may be modified by the stored produre if it includes output parameters.</param>
    ''' <param name="returnValue">Output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>The first column of the first row in the result set, or a null reference (Nothing
    ''' in Visual Basic) if the result set is empty.</returns>
    Private Function SPExecuteScalar(spName As String, parameterArray As SqlParameter(), ByRef returnValue As Integer) As Object
        If String.IsNullOrEmpty(spName) Then Throw New ArgumentException("The name of the stored procedure must be specified.", "spName")

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(spName, connection)

                ' Setup
                command.CommandType = CommandType.StoredProcedure

                If parameterArray IsNot Nothing AndAlso parameterArray.Count > 0 Then
                    DBNullifyParameters(parameterArray)
                    command.Parameters.AddRange(parameterArray)
                End If

                Dim returnParameter As SqlParameter = ReturnValueParameter()
                command.Parameters.Add(returnParameter)

                ' Run
                command.Connection.Open()
                Dim result As Object = command.ExecuteScalar()
                command.Connection.Close()

                ' Cleanup
                returnValue = returnParameter.Value
                command.Parameters.Remove(returnParameter)

                If parameterArray IsNot Nothing AndAlso parameterArray.Count > 0 Then
                    Dim newArray(command.Parameters.Count) As SqlParameter
                    command.Parameters.CopyTo(newArray, 0)
                    Array.Copy(newArray, parameterArray, parameterArray.Length)
                End If

                command.Parameters.Clear()

                ' Return
                Return result

            End Using
        End Using
    End Function

End Class
