Module DbObjects

    ' A
    Public Property DB_TEST_FUNCTION_A As String = "IntFunctionNoParam"
    Public Property DB_TEST_INT_A As Integer = 13

    ' B
    Public Property DB_TEST_FUNCTION_B As String = "IntFunctionWithParam"

    ' C
    Public Property DB_TEST_PROCEDURE_C As String = "InsertSpNoParam"
    Public Property DB_TEST_INT_C As Integer = 100
    Public Property DB_TEST_STRING_C As String = "test" & DB_TEST_INT_C

    ' D
    Public Property DB_TEST_PROCEDURE_D As String = "InsertSpWithParam"

    Public Property DbObjects As New List(Of String) From {
        "CREATE FUNCTION dbo." & DB_TEST_FUNCTION_A & " () RETURNS INT AS BEGIN RETURN " & DB_TEST_INT_A & " END",
        "CREATE FUNCTION dbo." & DB_TEST_FUNCTION_B & " (@param1 int, @param2 int) RETURNS INT AS BEGIN RETURN @param1 + @param2 END",
        "CREATE PROCEDURE " & DB_TEST_PROCEDURE_C & " AS BEGIN " &
        "    INSERT INTO dbo.Things (Name, Status, MandatoryInteger, MandatoryDate) VALUES ('" & DB_TEST_STRING_C & "', 1, " & DB_TEST_INT_C & ",  N'2016-04-01 00:00:00'); END;",
        "CREATE PROCEDURE " & DB_TEST_PROCEDURE_D & " @name varchar(max), @status bit, @mandatoryInteger int, @mandatoryDate date AS BEGIN " &
        "    INSERT INTO dbo.Things (Name, Status, MandatoryInteger, MandatoryDate) VALUES (@name, @status, @mandatoryInteger, @mandatoryDate); END;"
    }
End Module
