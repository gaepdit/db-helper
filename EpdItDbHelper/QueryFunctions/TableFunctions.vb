﻿Imports System.Data.SqlClient

Partial Public Class DB

#Region " Read (Lookup Dictionary) "

    ''' <summary>
    ''' Retrieves a dictionary of (integer -> string) values from the database
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <returns>A lookup dictionary.</returns>
    Public Function GetLookupDictionary(ByVal query As String, Optional ByVal parameter As SqlParameter = Nothing) _
        As Dictionary(Of Integer, String)
        Dim d As New Dictionary(Of Integer, String)

        Dim dataTable As DataTable = GetDataTable(query, parameter)

        For Each row As DataRow In dataTable.Rows
            d.Add(row.Item(0), row.Item(1))
        Next

        Return d
    End Function

#End Region

#Region " Read (DataRow) "

    ''' <summary>
    ''' Retrieves a single row of values from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <returns>A DataRow of values.</returns>
    Public Function GetDataRow(ByVal query As String, Optional ByVal parameter As SqlParameter = Nothing) As DataRow
        Dim parameterArray As SqlParameter() = {parameter}
        Return GetDataRow(query, parameterArray)
    End Function

    ''' <summary>
    ''' Retrieves a single row of values from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <returns>A DataRow of values.</returns>
    Public Function GetDataRow(ByVal query As String, ByVal parameterArray As SqlParameter()) As DataRow
        Dim resultTable As DataTable = GetDataTable(query, parameterArray)
        If resultTable IsNot Nothing And resultTable.Rows.Count = 1 Then
            Return resultTable.Rows(0)
        Else
            Return Nothing
        End If
    End Function

#End Region

#Region " Read (DataTable) "

    ''' <summary>
    ''' Retrieves a DataTable of values from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <returns>A DataTable of values.</returns>
    Public Function GetDataTable(ByVal query As String, Optional ByVal parameter As SqlParameter = Nothing) As DataTable
        Dim parameterArray As SqlParameter() = {parameter}
        Return GetDataTable(query, parameterArray)
    End Function

    ''' <summary>
    ''' Retrieves a DataTable of values from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameterArray">An SqlParameter array to send.</param>
    ''' <returns>A DataTable of values.</returns>
    Public Function GetDataTable(ByVal query As String, ByVal parameterArray As SqlParameter()) As DataTable
        Dim table As New DataTable

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(query, connection)
                command.CommandType = CommandType.Text
                command.Parameters.AddRange(parameterArray)
                Using adapter As New SqlDataAdapter(command)
                    command.Connection.Open()
                    adapter.Fill(table)
                    command.Connection.Close()
                End Using
            End Using
        End Using

        Return table
    End Function

#End Region

End Class
