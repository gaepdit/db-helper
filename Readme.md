# EPD-IT DB Helper Class Library

The purpose of this library is to simplify interactions with a SQL Server database. 

## What is this?

This library was originally written for working with an Oracle database. It is being migrated to work with SQL Server. **This is still a work in progress.** Some parts are working and some are not. Please contribute and help fix the parts that don't work yet.

See the status of each method at the bottom of this page. 

## How do I get set up?

* Clone this repository to your local machine.
* In Visual Studio, open YOUR solution where you want to use this library. 
* Choose File → Add → Existing Project...
* Find and select `EpdItDbHelper\EpdItDbHelper.vbproj` from this library.
* In Solution Explorer, select YOUR project, and choose Project → Add Reference...
* In the Reference Manager that opens, select Projects and check the box for the EpdItDbHelper project.

## How do I use the library?

There are two classes available:

### `DB`

The DB class must be instantiated with a connection string. (In the future, there may be more options for instantiating.) 

Example:

```
Friend DB As New EpdIt.DBHelper(ConnectionString)
```

The DB object has many functions available for interacting with the database. Some functions which require a SQL query or command and optional SQL parameters.

Example:

```
Dim userId as Integer = 3
Dim query as String = "select UserName from UserTable where UserId = @id"
Dim parameter As SqlParameter = New SqlParameter("id", userId)
Dim userName as String = DB.GetSingleValue(Of String)(query, parameter)
```

Other functions require a name of a Stored Procedure instead of a SQL query. These functions all start with "SP" in the name.

Example:

```
Dim spName as String = "RetrieveFacilitiesByCounty"
Dim parameterArray As SqlParameter() = {
    New SqlParameter("state", "GA"),
    New SqlParameter("county", "Fulton")
}
Dim facilities as DataTable = DB.SPGetDataTable(spName, parameterArray)
```

### `DBUtilities`

This class does not need to be instantiated. It includes several useful functions for working with database data. The most useful is `DBUtilities.GetNullable(Of T)` which helps handle data retrieved from nullable columns in the database (handling DBNull appropriately).

## How do I help make it better?

* See the TO-DO list below. 
* Write unit tests for any functions not yet covered. 
* Run the unit tests and fix any code that is broken.
* Fix the [XML documentation](https://msdn.microsoft.com/en-us/library/ms172652.aspx) for public functions.

### TO-DO

* AddRefCursorParameter(SqlParameter())
* GetByteArrayFromBlob(String, SqlParameter()) As Byte()
* SaveBinaryFileFromDB(String, String, SqlParameter()) As Boolean
* SaveBinaryFileFromDB(String, String, SqlParameter) As Boolean
* SPGetBoolean(String, SqlParameter(), Boolean) As Boolean
* SPGetBoolean(String, SqlParameter, Boolean) As Boolean
* SPGetDataRow(String, SqlParameter()) As DataRow
* SPGetDataRow(String, SqlParameter) As DataRow
* SPGetDataTable(String, SqlParameter()) As DataTable
* SPGetDataTable(String, SqlParameter) As DataTable
* SPGetInteger(String) As Integer
* SPGetListOfKeyValuePair(String, SqlParameter) As List(Of KeyValuePair(Of Integer, String))
* SPGetSingleValue(Of T)(String, SqlParameter(), Boolean) As T
* SPGetSingleValue(Of T)(String, SqlParameter, Boolean) As T

### DONE

* GetBoolean(String, SqlParameter(), Boolean) As Boolean
* GetBoolean(String, SqlParameter, Boolean) As Boolean
* GetSingleValue(Of T)(String, SqlParameter(), Boolean) As T
* GetSingleValue(Of T)(String, SqlParameter, Boolean) As T
* ValueExists(String, SqlParameter()) As Boolean
* ValueExists(String, SqlParameter) As Boolean
* RunCommand(String, SqlParameter(), Integer, Boolean) As Boolean
* RunCommand(String, SqlParameter, Integer, Boolean) As Boolean
* RunCommand(List(Of String), List(Of SqlParameter()), List(Of Integer), Boolean) As Boolean
* RunCommand(List(Of String), List(Of SqlParameter()), List(Of Integer), Boolean) As Boolean
* SPRunCommand(String, SqlParameter()) As Boolean
* SPRunCommand(String, SqlParameter) As Boolean
* GetDataRow(String, SqlParameter()) As DataRow
* GetDataRow(String, SqlParameter) As DataRow
* GetDataTable(String, SqlParameter()) As DataTable
* GetDataTable(String, SqlParameter) As DataTable
* GetLookupDictionary(String, SqlParameter) As Dictionary(Of Integer, String)
