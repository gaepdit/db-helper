Imports System.Data.SqlClient

Partial Public Class DB

    ' These functions call Oracle FUNCTIONs that return an Oracle SYS_REFCURSOR.

#Region " DataRow "

    ''' <summary>
    ''' Retrieves a single row of values from the database.
    ''' </summary>
    ''' <param name="spName">The Oracle Stored Procedure to call (SP must be a function that returns a REFCURSOR)</param>
    ''' <param name="parameter">An optional Oracle Parameter to send.</param>
    ''' <returns>A DataRow.</returns>
    Public Function SPGetDataRow(ByVal spName As String, Optional ByVal parameter As SqlParameter = Nothing) As DataRow
        Dim parameterArray As SqlParameter() = {parameter}
        Return SPGetDataRow(spName, parameterArray)
    End Function

    ''' <summary>
    ''' Retrieves a single row of values from the database.
    ''' </summary>
    ''' <param name="spName">The Oracle Stored Procedure to call (SP must be a function that returns a REFCURSOR)</param>
    ''' <param name="parameterArray">An optional Oracle Parameter to send.</param>
    ''' <returns>A DataRow</returns>
    Public Function SPGetDataRow(ByVal spName As String, ByVal parameterArray As SqlParameter()) As DataRow
        Dim resultTable As DataTable = SPGetDataTable(spName, parameterArray)
        If resultTable IsNot Nothing And resultTable.Rows.Count = 1 Then
            Return resultTable.Rows(0)
        Else
            Return Nothing
        End If
    End Function

#End Region

#Region " DataTable "

    ''' <summary>
    ''' Retrieves a DataTable of values from the database.
    ''' </summary>
    ''' <param name="spName">The Oracle Stored Procedure to call (SP must be a function that returns a REFCURSOR)</param>
    ''' <param name="parameter">An optional Oracle Parameter to send.</param>
    ''' <returns>A DataTable</returns>
    Public Function SPGetDataTable(ByVal spName As String, Optional ByVal parameter As SqlParameter = Nothing) As DataTable
        Dim parameterArray As SqlParameter() = {parameter}
        Return SPGetDataTable(spName, parameterArray)
    End Function

    ''' <summary>
    ''' Retrieves a DataTable of values from the database.
    ''' </summary>
    ''' <param name="spName">The Oracle Stored Procedure to call (SP must be a function that returns a REFCURSOR)</param>
    ''' <param name="parameterArray">An Oracle Parameter array to send.</param>
    ''' <returns>A DataTable</returns>
    Public Function SPGetDataTable(ByVal spName As String, ByVal parameterArray As SqlParameter()) As DataTable
        If String.IsNullOrEmpty(spName) Then
            Return Nothing
        End If

        AddRefCursorParameter(parameterArray)

        Dim table As New DataTable
        Dim success As Boolean = True
        Dim startTime As Date = Date.UtcNow

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(spName, connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddRange(parameterArray)
                Using adapter As New SqlDataAdapter(command)
                    Try
                        command.Connection.Open()
                        adapter.Fill(table)
                        command.Connection.Close()
                    Catch ee As SqlException
                        success = False
                        table = Nothing
                    End Try
                End Using
            End Using
        End Using

        Return table
    End Function

#End Region

#Region " Lists, Keys & Values "

    ''' <summary>
    ''' Calls an Oracle Stored Procedure and returns a List of KeyValuePairs with Integer keys and 
    ''' String values. Useful for creating DropDownList ComboBoxes.
    ''' </summary>
    ''' <param name="spName">The Oracle Stored Procedure to call (SP must be a function that returns a REFCURSOR)</param>
    ''' <param name="parameter">A single Oracle Parameter to pass in</param>
    ''' <returns>List of Integer keys and String value pairs</returns>
    ''' <remarks>Use List returned with ComboBox.BindToKeyValuePairs</remarks>
    Public Function SPGetListOfKeyValuePair(ByVal spName As String, Optional ByVal parameter As SqlParameter = Nothing) _
        As List(Of KeyValuePair(Of Integer, String))
        Dim l As New List(Of KeyValuePair(Of Integer, String))
        Dim dt As DataTable = SPGetDataTable(spName, parameter)

        For Each r As DataRow In dt.Rows
            l.Add(New KeyValuePair(Of Integer, String)(r.Item(0), DBUtilities.GetNullable(Of String)(r.Item(1))))
        Next

        Return l
    End Function

#End Region

#Region " SYS_REFCURSOR Utility "

    Private Sub AddRefCursorParameter(ByRef parameterArray As SqlParameter())
        Throw New NotImplementedException()

        ' TODO: SQL Server migration

        'Dim pRefCursor As New SqlParameter
        'pRefCursor.Direction = ParameterDirection.ReturnValue
        'pRefCursor.SqlDbType = OracleDbType.RefCursor

        'If parameterArray Is Nothing Then
        '    Array.Resize(parameterArray, 1)
        'ElseIf parameterArray(0) IsNot Nothing Then
        '    Array.Resize(parameterArray, parameterArray.Length + 1)
        'End If

        'parameterArray(parameterArray.GetUpperBound(0)) = pRefCursor
    End Sub

#End Region

End Class
