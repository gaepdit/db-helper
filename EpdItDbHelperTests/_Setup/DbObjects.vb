Module DbObjects

    Public Property UDF_TEST_VALUE As Integer = 13

    Public Property UdfObjects As New List(Of String) From {
        "CREATE FUNCTION dbo.IntFunctionNoParam () RETURNS INT AS BEGIN RETURN " & UDF_TEST_VALUE & " END",
        "CREATE FUNCTION dbo.IntFunctionWithParam (@param1 int, @param2 int) RETURNS INT AS BEGIN RETURN @param1 + @param2 END"
    }

End Module
