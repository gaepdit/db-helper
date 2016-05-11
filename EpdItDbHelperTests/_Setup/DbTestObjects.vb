Module DbTestObjects

    ' UDF no parameters
    Public Property DBO_UDF_NoParameters_NAME As String = "IntFunctionNoParam"
    Public Property DBO_UDF_NoParameters_VALUE As Integer = 300

    ' UDF with parameters
    Public Property DBO_UDF_WithParameters_NAME As String = "IntFunctionWithParam"

    ' SP insert no parameters
    Public Property DBO_SP_INSERT_NoParameters_NAME As String = "InsertSpNoParam"
    Public Property DBO_SP_INSERT_NoParameters_VALUE_int As Integer = 100
    Public Property DBO_SP_INSERT_NoParameters_VALUE_str As String = "test" & DBO_SP_INSERT_NoParameters_VALUE_int

    ' SP insert single parameter
    Public Property DBO_SP_INSERT_OneParameter_NAME As String = "InsertSpWith1Param"
    Public Property DBO_SP_INSERT_OneParameter_VALUE_int As Integer = 200
    Public Property DBO_SP_INSERT_OneParameter_VALUE_str As String = "test" & DBO_SP_INSERT_OneParameter_VALUE_int

    ' SP insert multiple parameters
    Public Property DBO_SP_INSERT_WithParameters_NAME As String = "InsertSpWithParam"

    ' SP table no parameters
    Public Property DBO_SP_Table_NoParameters_NAME As String = "TableSpNoParam"

    ' SP table one parameter
    Public Property DBO_SP_Table_OneParameter_NAME As String = "TableSpOneParam"

    ' SP row no parameters
    Public Property DBO_SP_Row_NoParameters_NAME As String = "RowSpNoParam"
    Public Property DBO_SP_Row_NoParameters_VALUE_int As Integer = 2

    ' SP row one parameter
    Public Property DBO_SP_Row_OneParameter_NAME As String = "RowSpOneParam"
    Public Property DBO_SP_Row_OneParameter_VALUE_int As Integer = 2

    ' SP dict no param
    Public Property DBO_SP_Dictionary_NoParameters_NAME As String = "DictSpNoParam"

    ' SP dict one parameter
    Public Property DBO_SP_Dictionary_OneParameter_NAME As String = "DictSpOneParam"


    Public Property DbTestObjectStrings As New List(Of String) From {
        "CREATE FUNCTION dbo." & DBO_UDF_NoParameters_NAME & " () RETURNS INT AS BEGIN RETURN " & DBO_UDF_NoParameters_VALUE & " END",
        "CREATE FUNCTION dbo." & DBO_UDF_WithParameters_NAME & " (@param1 int, @param2 int) RETURNS INT AS BEGIN RETURN @param1 + @param2 END",
        "CREATE PROCEDURE " & DBO_SP_INSERT_NoParameters_NAME & " AS BEGIN " &
        "     INSERT INTO dbo.Things (Name, Status, MandatoryInteger, MandatoryDate) VALUES ('" & DBO_SP_INSERT_NoParameters_VALUE_str & "', 1, " & DBO_SP_INSERT_NoParameters_VALUE_int & ",  N'2016-04-01 00:00:00'); END;",
        "CREATE PROCEDURE " & DBO_SP_INSERT_OneParameter_NAME & " @mandatoryInteger int AS BEGIN " &
        "     INSERT INTO dbo.Things (Name, Status, MandatoryInteger, MandatoryDate) VALUES ('" & DBO_SP_INSERT_OneParameter_VALUE_str & "', 1, @mandatoryInteger,  N'2016-04-01 00:00:00'); END;",
        "CREATE PROCEDURE " & DBO_SP_INSERT_WithParameters_NAME & " @name varchar(max), @status bit, @mandatoryInteger int, @mandatoryDate date AS BEGIN " &
        "     INSERT INTO dbo.Things (Name, Status, MandatoryInteger, MandatoryDate) VALUES (@name, @status, @mandatoryInteger, @mandatoryDate); END;",
        "CREATE PROCEDURE " & DBO_SP_Table_NoParameters_NAME & " AS BEGIN " &
        "     SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE MandatoryInteger = " & ThingData.ThingSelectionKey & " ORDER BY [ID]; END;",
        "CREATE PROCEDURE " & DBO_SP_Table_OneParameter_NAME & " @mandatoryInteger int AS BEGIN " &
        "     SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE MandatoryInteger = @mandatoryInteger ORDER BY [ID]; END;",
        "CREATE PROCEDURE " & DBO_SP_Row_NoParameters_NAME & " AS BEGIN " &
        "     SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE ID = " & DBO_SP_Row_NoParameters_VALUE_int & "; END;",
        "CREATE PROCEDURE " & DBO_SP_Row_OneParameter_NAME & " @id int AS BEGIN " &
        "     SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE ID = @id; END;",
        "CREATE PROCEDURE " & DBO_SP_Dictionary_NoParameters_NAME & " AS BEGIN " &
        "     SELECT ID, Name FROM dbo.Things WHERE MandatoryInteger = " & ThingData.ThingSelectionKey & " ORDER BY [ID]; END;",
        "CREATE PROCEDURE " & DBO_SP_Dictionary_OneParameter_NAME & " @mandatoryInteger int AS BEGIN " &
        "     SELECT ID, Name FROM dbo.Things WHERE MandatoryInteger = @mandatoryInteger ORDER BY [ID]; END;"
    }
End Module
