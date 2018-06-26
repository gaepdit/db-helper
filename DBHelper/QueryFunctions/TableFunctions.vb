Imports System.Data.SqlClient

Partial Public Class DBHelper

#Region " DataRow "

    ''' <summary>
    ''' Retrieves a single row of values from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <returns>A DataRow of values.</returns>
    Public Function GetDataRow(query As String, Optional parameter As SqlParameter = Nothing) As DataRow
        Dim parameterArray As SqlParameter() = Nothing

        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If

        Return GetDataRow(query, parameterArray)
    End Function

    ''' <summary>
    ''' Retrieves a single row of values from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameterArray">An array of SqlParameter values to send.</param>
    ''' <returns>A DataRow of values.</returns>
    Public Function GetDataRow(query As String, parameterArray As SqlParameter()) As DataRow
        Dim dataTable As DataTable = GetDataTable(query, parameterArray)

        If dataTable Is Nothing OrElse dataTable.Rows.Count = 0 Then
            Return Nothing
        ElseIf dataTable.Rows.Count > 1 Then
            Throw New TooManyRecordsException()
        End If

        Return dataTable.Rows(0)
    End Function

#End Region

#Region " DataTable "

    ''' <summary>
    ''' Retrieves a DataTable of values from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <returns>A DataTable of values.</returns>
    Public Function GetDataTable(query As String, Optional parameter As SqlParameter = Nothing) As DataTable
        Dim parameterArray As SqlParameter() = Nothing

        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If

        Return GetDataTable(query, parameterArray)
    End Function

    ''' <summary>
    ''' Retrieves a DataTable of values from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameterArray">An array of SqlParameter values to send.</param>
    ''' <returns>A DataTable of values.</returns>
    Public Function GetDataTable(query As String, parameterArray As SqlParameter()) As DataTable
        Return QFillDataTable(query, parameterArray)
    End Function

#End Region

#Region " Lookup Dictionary "

    ''' <summary>
    ''' Retrieves a dictionary of (integer -> string) values from the database
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <returns>A lookup dictionary.</returns>
    Public Function GetLookupDictionary(query As String,
                                        Optional parameter As SqlParameter = Nothing
                                        ) As Dictionary(Of Integer, String)

        Dim parameterArray As SqlParameter() = Nothing

        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If

        Return GetLookupDictionary(query, parameterArray)
    End Function

    ''' <summary>
    ''' Retrieves a dictionary of (integer -> string) values from the database
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameterArray">An array of SqlParameter values to send.</param>
    ''' <returns>A lookup dictionary.</returns>
    Public Function GetLookupDictionary(query As String,
                                        parameterArray As SqlParameter()
                                        ) As Dictionary(Of Integer, String)
        Dim d As New Dictionary(Of Integer, String)

        Dim dataTable As DataTable = GetDataTable(query, parameterArray)

        For Each row As DataRow In dataTable.Rows
            d.Add(row.Item(0), row.Item(1))
        Next

        Return d
    End Function

#End Region

End Class
