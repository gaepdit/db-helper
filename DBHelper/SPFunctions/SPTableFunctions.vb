Imports System.Data.SqlClient

Partial Public Class DBHelper

#Region " DataRow "

    ''' <summary>
    ''' Retrieves a single row of values from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameter">An optional SqlParameter value. The value may be modified by the stored produre if it is an output parameter.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>A DataRow.</returns>
    Public Function SPGetDataRow(spName As String,
                                 ByRef Optional parameter As SqlParameter = Nothing,
                                 Optional ByRef returnValue As Integer = Nothing
                                 ) As DataRow

        Dim parameterArray As SqlParameter() = Nothing

        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If

        Dim dataRow As DataRow = SPGetDataRow(spName, parameterArray, returnValue)

        If parameterArray IsNot Nothing AndAlso parameterArray.Count > 0 Then
            parameter = parameterArray(0)
        End If

        Return dataRow
    End Function

    ''' <summary>
    ''' Retrieves a single row of values from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameterArray">An array of SqlParameter values. The array may be modified by the stored produre if it includes output parameters.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>A DataRow.</returns>
    Public Function SPGetDataRow(spName As String,
                                 ByRef parameterArray As SqlParameter(),
                                 Optional ByRef returnValue As Integer = Nothing
                                 ) As DataRow

        Dim dataTable As DataTable = SPGetDataTable(spName, parameterArray, returnValue)

        If dataTable Is Nothing Then
            Return Nothing
        ElseIf dataTable.Rows.Count > 1 Then
            Throw New TooManyRecordsException(spName)
        End If

        Return dataTable.Rows(0)
    End Function

#End Region

#Region " DataTable "

    ''' <summary>
    ''' Retrieves a DataTable of values from the database.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameter">An optional SqlParameter value. The value may be modified by the stored produre if it is an output parameter.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>A DataTable.</returns>
    Public Function SPGetDataTable(spName As String,
                                   Optional ByRef parameter As SqlParameter = Nothing,
                                   Optional ByRef returnValue As Integer = Nothing
                                   ) As DataTable

        Dim parameterArray As SqlParameter() = Nothing

        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If

        Dim table As DataTable = SPGetDataTable(spName, parameterArray, returnValue)

        If parameterArray IsNot Nothing AndAlso parameterArray.Count > 0 Then
            parameter = parameterArray(0)
        End If

        Return table
    End Function

    ''' <summary>
    ''' Retrieves a DataTable of values from the database.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameterArray">An array of SqlParameter values. The array may be modified by the stored produre if it includes output parameters.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>A DataTable</returns>
    Public Function SPGetDataTable(spName As String,
                                   ByRef parameterArray As SqlParameter(),
                                   Optional ByRef returnValue As Integer = Nothing
                                   ) As DataTable

        Dim dataSet As DataSet = SPGetDataSet(spName, parameterArray, returnValue)

        If dataSet Is Nothing Then
            Return Nothing
        ElseIf dataSet.Tables.Count > 1 Then
            Throw New TooManyRecordsException(spName)
        End If

        Return dataSet.Tables(0)
    End Function

#End Region

#Region " DataSet "

    ''' <summary>
    ''' Retrieves a DataSet containing one or more DataTables selected from the database.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameter">An optional SqlParameter value. The value may be modified by the stored produre if it is an output parameter.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>A DataSet.</returns>
    Public Function SPGetDataSet(spName As String,
                                 Optional ByRef parameter As SqlParameter = Nothing,
                                 Optional ByRef returnValue As Integer = Nothing
                                 ) As DataSet

        Dim parameterArray As SqlParameter() = Nothing

        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If

        Dim dataSet As DataSet = SPGetDataSet(spName, parameterArray, returnValue)

        If parameterArray IsNot Nothing AndAlso parameterArray.Count > 0 Then
            parameter = parameterArray(0)
        End If

        Return dataSet
    End Function

    ''' <summary>
    ''' Retrieves a DataSet containing one or more DataTables selected from the database.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameterArray">An array of SqlParameter values. The array may be modified by the stored produre if it includes output parameters.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>A DataSet</returns>
    Public Function SPGetDataSet(spName As String,
                                 ByRef parameterArray As SqlParameter(),
                                 Optional ByRef returnValue As Integer = Nothing
                                 ) As DataSet

        Return SPFillDataSet(spName, parameterArray, returnValue)

    End Function

#End Region

#Region " Lookup Dictionary "

    ''' <summary>
    ''' Retrieves a dictionary of (integer -> string) values from the database
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameter">A SqlParameter value. The value may be modified by the stored produre if it is an output parameter.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>A lookup dictionary.</returns>
    Public Function SPGetLookupDictionary(spName As String,
                                          ByRef Optional parameter As SqlParameter = Nothing,
                                          Optional ByRef returnValue As Integer = Nothing
                                          ) As Dictionary(Of Integer, String)

        Dim parameterArray As SqlParameter() = Nothing

        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If

        Dim d As Dictionary(Of Integer, String) = SPGetLookupDictionary(spName, parameterArray, returnValue)

        If parameterArray IsNot Nothing AndAlso parameterArray.Count > 0 Then
            parameter = parameterArray(0)
        End If

        Return d
    End Function

    ''' <summary>
    ''' Retrieves a dictionary of (integer -> string) values from the database
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameterArray">An array of SqlParameter values. The array may be modified by the stored produre if it includes output parameters.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>A lookup dictionary.</returns>
    Public Function SPGetLookupDictionary(spName As String,
                                          ByRef parameterArray As SqlParameter(),
                                          Optional ByRef returnValue As Integer = Nothing
                                          ) As Dictionary(Of Integer, String)

        Dim d As New Dictionary(Of Integer, String)

        Dim dataTable As DataTable = SPGetDataTable(spName, parameterArray, returnValue)

        For Each row As DataRow In dataTable.Rows
            d.Add(row.Item(0), row.Item(1))
        Next

        Return d
    End Function

#End Region

End Class
