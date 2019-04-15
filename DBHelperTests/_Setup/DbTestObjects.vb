Module DbTestObjects

    Public Property DbTestObjectStrings As New List(Of String) From {
        "CREATE FUNCTION dbo.IntFunctionNoParam () RETURNS INT AS BEGIN RETURN 300 END",
        "CREATE FUNCTION dbo.IntFunctionWithParam (@param1 int, @param2 int) RETURNS INT AS BEGIN RETURN @param1 + @param2 END",
        "CREATE PROCEDURE InsertSpNoParam AS BEGIN 
             INSERT INTO dbo.Things (Name, Status, MandatoryInteger, MandatoryDate) VALUES ('test100', 1, " & 100 & ",  N'2016-04-01 00:00:00'); END;
        ",
        "CREATE PROCEDURE InsertSpNoParamRV AS BEGIN 
             INSERT INTO dbo.Things (Name, Status, MandatoryInteger, MandatoryDate) VALUES ('test100', 1, 100,  N'2016-04-01 00:00:00'); return 2; END;
        ",
        "CREATE PROCEDURE InsertSpWith1Param @mandatoryInteger int AS BEGIN 
             INSERT INTO dbo.Things (Name, Status, MandatoryInteger, MandatoryDate) VALUES ('test200', 1, @mandatoryInteger,  N'2016-04-01 00:00:00'); END;
        ",
        "CREATE PROCEDURE InsertSpWithParam @name varchar(max), @status bit, @mandatoryInteger int, @mandatoryDate date AS BEGIN 
             INSERT INTO dbo.Things (Name, Status, MandatoryInteger, MandatoryDate) VALUES (@name, @status, @mandatoryInteger, @mandatoryDate); END;
        ",
        "CREATE PROCEDURE TableSpNoParam AS BEGIN 
             SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE MandatoryInteger = " & ThingData.ThingSelectionKey & " ORDER BY [ID]; END;
        ",
        "CREATE PROCEDURE TableSpOneParam @mandatoryInteger int AS BEGIN 
             SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE MandatoryInteger = @mandatoryInteger ORDER BY [ID]; END;
        ",
        "CREATE PROCEDURE RowSpNoParam AS BEGIN 
             SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE ID = 2; END;",
        "CREATE PROCEDURE RowSpOneParam @id int AS BEGIN 
             SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE ID = @id; END;
        ",
        "CREATE PROCEDURE DictSpNoParam AS BEGIN 
             SELECT ID, Name FROM dbo.Things WHERE MandatoryInteger = " & ThingData.ThingSelectionKey & " ORDER BY [ID]; END;
        ",
        "CREATE PROCEDURE DictSpOneParam @mandatoryInteger int AS BEGIN 
             SELECT ID, Name FROM dbo.Things WHERE MandatoryInteger = @mandatoryInteger ORDER BY [ID]; END;
        ",
        "CREATE PROCEDURE DataSetSpNoParam AS BEGIN 
             SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE MandatoryInteger = " & ThingData.ThingSelectionKey & " ORDER BY [ID]; 
             SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE MandatoryInteger <> " & ThingData.ThingSelectionKey & " ORDER BY [ID]; 
             END;
        ",
        "CREATE PROCEDURE SpOutParam @id int, @name varchar(max) OUTPUT AS BEGIN 
             SET NOCOUNT ON; 
             SELECT @name = Name FROM dbo.Things WHERE id = @id; END;
        ",
        "CREATE PROCEDURE SpOutReturnString @id int, @return_value_argument varchar(max) OUTPUT AS BEGIN 
             SET NOCOUNT ON; 
             SELECT @return_value_argument = Name FROM dbo.Things WHERE id = @id; END;
        ",
        "CREATE PROCEDURE SpOutReturnIntParam @id int, @return_value_argument int OUTPUT AS BEGIN 
             SET NOCOUNT ON; 
             SELECT @return_value_argument = MandatoryInteger FROM dbo.Things WHERE id = @id; END;
        ",
        "CREATE PROCEDURE SpOutReturnBoolParam @id int, @return_value_argument bit OUTPUT AS BEGIN 
             SET NOCOUNT ON; 
             SELECT @return_value_argument = Status FROM dbo.Things WHERE id = @id; END;
        ",
        "CREATE PROCEDURE SpOutReturnNullDateParam @id int, @return_value_argument date OUTPUT AS BEGIN 
             SET NOCOUNT ON; 
             SELECT @return_value_argument = OptionalDate FROM dbo.Things WHERE id = @id; END;
        ",
        "CREATE PROCEDURE SpGetBoolean @select BIT AS BEGIN 
             IF @select = 1 SELECT 1; ELSE SELECT 0; END;
        ",
        "CREATE PROCEDURE SpGetString @select int AS BEGIN 
             SELECT CASE WHEN @select = 0 THEN NULL WHEN @select = 1 THEN '' WHEN @select = 2 THEN 'abc' END; END;
        ",
        "CREATE PROCEDURE SpGetInt @select int AS BEGIN 
             SELECT CASE WHEN @select = 0 THEN 0 WHEN @select = 1 THEN 1 WHEN @select = 2 THEN -1 END; END;
        ",
        "CREATE PROCEDURE SpGetNullableInt @select BIT AS BEGIN 
             IF @select = 1 SELECT 1; ELSE SELECT NULL; END;
        ",
        "CREATE PROCEDURE SpGetBooleanAndReturn @select BIT AS BEGIN 
             IF @select = 1 SELECT 1; ELSE SELECT 0; return 2; END;
        "
    }
End Module
