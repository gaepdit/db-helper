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

See the automatically-generated documentation in the documentation folder.

There are two classes available:

### `DB` Class

The DB class must be instantiated with a connection string. (In the future, there may be more options for instantiating. Let me know if this is a priority for you.) 

Example:

```
Friend DB As New EpdIt.DBHelper(ConnectionString)
```

The DB object has many functions available for interacting with the database. They differ based on whether they require a SQL query string as input, or the name of a Stored Procedure. 

#### Query string functions

These functions require a SQL query or command plus optional SQL parameters.

**Example 1:** No SQL parameters

```
Dim query as String = "select States from StatesTable"
Dim states as DataTable = DB.GetDataTable(query)
```

**Example 2:** One SQL parameter

```
Dim query as String = "select UserName from UserTable where UserId = @id"
Dim parameter As New SqlParameter("@id", MyUserId)
Dim userName as String = DB.GetSingleValue(Of String)(query, parameter)
```

**Example 3:** Multiple SQL parameters

```
Dim query as String = "update FacilityTable set Name = @name where FacilityId = @id"
Dim parameterArray As SqlParameter() = {
    New SqlParameter("@name", MyNewFacilityName),
    New SqlParameter("@id", MyFacilityId)
}
Dim userName as String = DB.GetDataTable(query, parameterArray)
```

#### Stored Procedure functions

These functions require the name of a Stored Procedure instead of a SQL query. These functions all start with "SP" in the name. 

Any combination of INPUT or OUTPUT parameters can be used, and the OUTPUT parameters will contain their values as sent by the database.

**Example 1:** Specifying INPUT and OUTPUT parameters

```
Dim spName as String = "RetrieveFacilitiesByCounty"
Dim parameterArray As SqlParameter() = {
    New SqlParameter("@county", MyCounty),
    New SqlParameter("@total", SqlDbType.Int) With {
        .Direction = ParameterDirection.Output
    }
}
Dim facilities as DataTable = DB.SPGetDataTable(spName, parameterArray)
```

Some convenience functions are available that return a single value from a stored procedure. These require the stored procedure to be written with a single OUTPUT parameter named `@return_value_argument`. The OUTPUT parameter does not need to be specified when calling the function.

**Example 2:** Single implied OUTPUT parameter

```
Dim spName as String = "RetrieveUserStatus" 
' RetrieveUserStatus must have an OUTPUT bit parameter named @return_value_argument
Dim parameter As New SqlParameter("id", userId)
Dim status as Boolean = DB.SPGetBoolean(spName, parameter)
```

### `DBUtilities` Class

This class does not need to be instantiated. It includes several useful shared functions for working with database data. The most useful is `DBUtilities.GetNullable(Of T)` which helps handle data retrieved from nullable columns in the database (handling DBNull appropriately).

## How can I help make it better?

* See the TO-DO list below for functions that have not been completed or tested.
* Add new functions as needed.
* Write unit tests for any functions not yet covered. 
* Run the unit tests and fix any code that is broken.
* Review the [XML documentation](https://msdn.microsoft.com/en-us/library/ms172652.aspx) for public functions.

### TO-DO

* GetByteArrayFromBlob(String, SqlParameter()) As Byte()
* SaveBinaryFileFromDB(String, String, SqlParameter()) As Boolean
* SaveBinaryFileFromDB(String, String, SqlParameter) As Boolean

### Missing Unit Tests

* SP Table Functions with OUTPUT parameters (SPGetDataTable, etc.)

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
* SPGetDataRow(String, SqlParameter()) As DataRow
* SPGetDataRow(String, SqlParameter) As DataRow
* SPGetDataTable(String, SqlParameter()) As DataTable
* SPGetDataTable(String, SqlParameter) As DataTable
* SPGetLookupDictionary(String, SqlParameter) As Dictionary(Of Integer, String)
* SPGetBoolean(String, SqlParameter(), Boolean) As Boolean
* SPGetBoolean(String, SqlParameter, Boolean) As Boolean
* SPGetSingleValue(Of T)(String, SqlParameter(), Boolean) As T
* SPGetSingleValue(Of T)(String, SqlParameter, Boolean) As T
