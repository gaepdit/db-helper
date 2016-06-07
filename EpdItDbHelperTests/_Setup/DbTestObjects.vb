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

    ' SP with OUT parameter
    Public Property DBO_SP_Output_Parameter_NAME As String = "SpOutParam"
    Public Property DBO_SP_Output_Parameter_VALUE_int As Integer = 2

    ' SP with return_value_argument as string
    Public Property DBO_SP_Return_String_NAME As String = "SpOutReturnString"
    Public Property DBO_SP_Return_String_VALUE As Integer = 2

    ' SP with return_value_argument as integer
    Public Property DBO_SP_Return_Int_NAME As String = "SpOutReturnIntParam"
    Public Property DBO_SP_Return_Int_VALUE As Integer = 2

    ' SP with return_value_argument as boolean
    Public Property DBO_SP_Return_Bool_NAME As String = "SpOutReturnBoolParam"
    Public Property DBO_SP_Return_Bool_VALUE As Integer = 2

    ' SP with return_value_argument as nullable date
    Public Property DBO_SP_Return_NullDate_NAME As String = "SpOutReturnNullDateParam"
    Public Property DBO_SP_Return_NullDate_VALUE As Integer = 2

    ' SP get scalar boolean
    Public Property DBO_SP_Get_Boolean As String = "SpGetBoolean"

    ' SP get scalar string
    Public Property DBO_SP_Get_String As String = "SpGetString"

    ' SP get scalar int
    Public Property DBO_SP_Get_Int As String = "SpGetInt"

    ' SP get scalar nullable int
    Public Property DBO_SP_Get_Nullable_Int As String = "SpGetNullableInt"

    Public Property DbTestObjectStrings As New List(Of String) From {
        "CREATE FUNCTION dbo." & DBO_UDF_NoParameters_NAME & " () RETURNS INT AS BEGIN RETURN " & DBO_UDF_NoParameters_VALUE & " END",
        "CREATE FUNCTION dbo." & DBO_UDF_WithParameters_NAME & " (@param1 int, @param2 int) RETURNS INT AS BEGIN RETURN @param1 + @param2 END",
        "CREATE PROCEDURE " & DBO_SP_INSERT_NoParameters_NAME & " AS BEGIN 
             INSERT INTO dbo.Things (Name, Status, MandatoryInteger, MandatoryDate) VALUES ('" & DBO_SP_INSERT_NoParameters_VALUE_str & "', 1, " & DBO_SP_INSERT_NoParameters_VALUE_int & ",  N'2016-04-01 00:00:00'); END;
        ",
        "CREATE PROCEDURE " & DBO_SP_INSERT_OneParameter_NAME & " @mandatoryInteger int AS BEGIN 
             INSERT INTO dbo.Things (Name, Status, MandatoryInteger, MandatoryDate) VALUES ('" & DBO_SP_INSERT_OneParameter_VALUE_str & "', 1, @mandatoryInteger,  N'2016-04-01 00:00:00'); END;
        ",
        "CREATE PROCEDURE " & DBO_SP_INSERT_WithParameters_NAME & " @name varchar(max), @status bit, @mandatoryInteger int, @mandatoryDate date AS BEGIN 
             INSERT INTO dbo.Things (Name, Status, MandatoryInteger, MandatoryDate) VALUES (@name, @status, @mandatoryInteger, @mandatoryDate); END;
        ",
        "CREATE PROCEDURE " & DBO_SP_Table_NoParameters_NAME & " AS BEGIN 
             SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE MandatoryInteger = " & ThingData.ThingSelectionKey & " ORDER BY [ID]; END;
        ",
        "CREATE PROCEDURE " & DBO_SP_Table_OneParameter_NAME & " @mandatoryInteger int AS BEGIN 
             SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE MandatoryInteger = @mandatoryInteger ORDER BY [ID]; END;
        ",
        "CREATE PROCEDURE " & DBO_SP_Row_NoParameters_NAME & " AS BEGIN 
             SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE ID = " & DBO_SP_Row_NoParameters_VALUE_int & "; END;",
        "CREATE PROCEDURE " & DBO_SP_Row_OneParameter_NAME & " @id int AS BEGIN 
             SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE ID = @id; END;
        ",
        "CREATE PROCEDURE " & DBO_SP_Dictionary_NoParameters_NAME & " AS BEGIN 
             SELECT ID, Name FROM dbo.Things WHERE MandatoryInteger = " & ThingData.ThingSelectionKey & " ORDER BY [ID]; END;
        ",
        "CREATE PROCEDURE " & DBO_SP_Dictionary_OneParameter_NAME & " @mandatoryInteger int AS BEGIN 
             SELECT ID, Name FROM dbo.Things WHERE MandatoryInteger = @mandatoryInteger ORDER BY [ID]; END;
        ",
        "CREATE PROCEDURE " & DBO_SP_Output_Parameter_NAME & " @id int, @name varchar(max) OUTPUT AS BEGIN 
             SET NOCOUNT ON; 
             SELECT @name = Name FROM dbo.Things WHERE id = @id; END;
        ",
        "CREATE PROCEDURE " & DBO_SP_Return_String_NAME & " @id int, @return_value_argument varchar(max) OUTPUT AS BEGIN 
             SET NOCOUNT ON; 
             SELECT @return_value_argument = Name FROM dbo.Things WHERE id = @id; END;
        ",
        "CREATE PROCEDURE " & DBO_SP_Return_Int_NAME & " @id int, @return_value_argument int OUTPUT AS BEGIN 
             SET NOCOUNT ON; 
             SELECT @return_value_argument = MandatoryInteger FROM dbo.Things WHERE id = @id; END;
        ",
        "CREATE PROCEDURE " & DBO_SP_Return_Bool_NAME & " @id int, @return_value_argument bit OUTPUT AS BEGIN 
             SET NOCOUNT ON; 
             SELECT @return_value_argument = Status FROM dbo.Things WHERE id = @id; END;
        ",
        "CREATE PROCEDURE " & DBO_SP_Return_NullDate_NAME & " @id int, @return_value_argument date OUTPUT AS BEGIN 
             SET NOCOUNT ON; 
             SELECT @return_value_argument = OptionalDate FROM dbo.Things WHERE id = @id; END;
        ",
        "CREATE PROCEDURE " & DBO_SP_Get_Boolean & " @select BIT AS BEGIN 
             IF @select = 1 SELECT 1; ELSE SELECT 0; END;
        ",
        "CREATE PROCEDURE " & DBO_SP_Get_String & " @select int AS BEGIN 
             SELECT CASE WHEN @select = 0 THEN NULL WHEN @select = 1 THEN '' WHEN @select = 2 THEN 'abc' END; END;
        ",
        "CREATE PROCEDURE " & DBO_SP_Get_Int & " @select int AS BEGIN 
             SELECT CASE WHEN @select = 0 THEN 0 WHEN @select = 1 THEN 1 WHEN @select = 2 THEN -1 END; END;
        ",
        "CREATE PROCEDURE " & DBO_SP_Get_Nullable_Int & " @select BIT AS BEGIN 
             IF @select = 1 SELECT 1; ELSE SELECT NULL; END;
        "
    }
End Module
