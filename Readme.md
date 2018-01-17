# EPD-IT Simple DB Helper

The purpose of this library is to simplify interactions with a SQL Server database. 

[![NuGet](https://img.shields.io/nuget/v/EpdIt.DbHelper.svg?maxAge=86400)](https://www.nuget.org/packages/EpdIt.DbHelper/) 

## What is this?

This library was originally written for working with an Oracle database. It is being migrated to work with SQL Server. **This is still a work in progress.** Some parts are working and some are not. Please contribute and help fix the parts that don't work yet.

See the status of each method at the bottom of this page. 

## How do I use the library?

To install Simple DB Helper, run the following command in the [Package Manager Console](https://docs.nuget.org/consume/package-manager-console):

`PM> Install-Package EpdIt.DbHelper`

There are two classes available:

### The `DB` Class

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

### The `DBUtilities` Class

This class does not need to be instantiated and only includes shared functions:

* `GetNullable(Of T)` converts a database value to a generic, useable .NET value, handling DBNull appropriately
* `GetNullableDateTime` converts a database value to a nullable DateTime object, handling DBNull appropriately
* `TvpSqlParameter(Of T)` converts an IEnumerable of type T to a structured, table-valued SqlParameter

## Breaking change in version 2.0.0!

*The `forceAddNullableParameters` parameter now defaults to `true`.* 

If this parameter is not set (or is manually set to `true`), then `DBNull.Value` will be sent for `SqlParameter`'s that evaluate to `Nothing`. To return to the default behavior of dropping such parameters, you must manually set `forceAddNullableParameters` to `false`.

## How can I help make it better?

* Write documentation for all available functions. 
* Add new functions as needed.
* Write unit tests for any functions not yet covered. 
* Run the unit tests and fix any code that is broken.
* Review the [XML documentation](https://msdn.microsoft.com/en-us/library/ms172652.aspx) for public functions.

Note to self: To push changes to NuGet.org, build a Release version and run `nuget push EpdIt.DbHelper.x.x.x.x.nupkg -Source https://www.nuget.org/api/v2/package`

### TO-DO

* Add unit tests for SP Table Functions with OUTPUT parameters (SPGetDataTable, etc.)
* Add unit tests for the `forceAddNullableParameters` parameter
* Add unit tests for table-valued parameters (`TvpSqlParameter`)
