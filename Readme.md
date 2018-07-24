# EPD-IT DB Helper

The purpose of this library is to simplify interactions with a SQL Server database. 

[![NuGet](https://img.shields.io/nuget/v/EpdIt.DbHelper.svg?maxAge=86400)](https://www.nuget.org/packages/EpdIt.DbHelper/) 

## What is this?

There are two classes available:

* The `DBHelper` class has many functions available for interacting with a SQL Server database. These functions take care of the tedious parts of querying data or running commands on the database, such as handling connection objects and transactions. There are separate sets of functions depending on whether a SQL query string is used as input or the name of a stored procedure.

* There is also a `DBUtilities` class with some static functions that simplify working with DBNull values and table-valued SQL parameters.

## How do I install it?

To install DB Helper, search for "EpdIt.DbHelper" in the NuGet package manager or run the following command in the [Package Manager Console](https://docs.nuget.org/consume/package-manager-console):

`PM> Install-Package EpdIt.DbHelper`

## Using the `DBHelper` class

The `DBHelper` class must be instantiated with a connection string:

```
Public DB As New EpdIt.DBHelper(connectionString)
```

### Query string functions

These functions require a SQL query or command plus an optional SQL parameter or array of SQL parameters. When running a non-query command, such as `INSERT` or `UPDATE`, an optional output parameter will contain the number of rows affected.

**Example 1:** Simple query

```
Dim query as String = "select States from StatesTable"
Dim states as DataTable = DB.GetDataTable(query)
```

**Example 2:** Query with one SQL parameter

```
Dim query as String = "select UserName from UserTable where UserId = @id"
Dim parameter As New SqlParameter("@id", MyUserId)
Dim userName as String = DB.GetSingleValue(Of String)(query, parameter)
```

**Example 3:** Command with multiple SQL parameters

```
Dim query as String = "update FacilityTable set Name = @name where FacilityId = @id"
Dim parameterArray As SqlParameter() = {
    New SqlParameter("@name", MyNewFacilityName),
    New SqlParameter("@id", MyFacilityId)
}
Dim result as Boolean = DB.RunCommand(query, parameterArray)
```

**Example 4:** Count rows affected by command

```
Dim query as String = "update CompanyTable set Status = @status where State = @state"
Dim parameterArray As SqlParameter() = {
    New SqlParameter("@status", MyNewStatus),
    New SqlParameter("@state", AffectedState)
}
Dim rowsAffected as Integer
Dim result as Boolean = DB.RunCommand(query, parameterArray, rowsAffected)
Console.WriteLine(rowsAffected & " rows affected.")
```

### Stored procedure functions

These functions require the name of a stored procedure instead of a SQL query, but otherwise are very similar to the query functions. These functions all start with "SP" in the name. 

An optional output parameter will contain the integer RETURN value of the stored procedure. (Often, a return value of 0 indicates success and a nonzero value indicates failure, but this depends on the particular stored procedure.)

**Example 1:** Specifying INPUT and OUTPUT SQL parameters

```
Dim spName as String = "RetrieveFacilitiesByCounty"
Dim returnParam As New SqlParameter("@total", SqlDbType.Int) With {
    .Direction = ParameterDirection.Output
}
Dim parameterArray As SqlParameter() = {
    New SqlParameter("@county", MyCounty),
    returnParam
}
Dim facilities as DataTable = DB.SPGetDataTable(spName, parameterArray)
Dim total As Integer = returnParam.Value
```

**Example 2:** Querying for a DataSet and using RETURN value

```
Dim spName As String = "GetMyData"
Dim returnValue As Integer
Dim dataSet As DataSet = DB.SPGetDataSet(spName, returnValue:=returnValue)
```

## Using the `DBUtilities` class

This class does not need to be instantiated and only includes shared functions:

* `GetNullable(Of T)` converts a database value to a generic, useable .NET value, handling DBNull appropriately
* `GetNullableString` converts a database value to a string, handling DBNull appropriately
* `GetNullableDateTime` converts a database value to a nullable DateTime, handling DBNull appropriately
* `TvpSqlParameter(Of T)` converts an IEnumerable of type T to a structured, table-valued SqlParameter

## Breaking changes in version 3

* The `forceAddNullableParameters` parameter has been removed. `DBNull.Value` will be sent for `SqlParameter`'s that evaluate to null (`Nothing` in VB.NET).
* The output parameter convenience functions have been removed. If you need an output SQL parameter, just add it as you would normally add any other parameter.

## Breaking changes in version 2

* The `forceAddNullableParameters` parameter now defaults to `true`. If this parameter is not set (or is manually set to `true`), then `DBNull.Value` will be sent for `SqlParameter`'s that evaluate to `Nothing`. To return to the default behavior of dropping such parameters, you must manually set `forceAddNullableParameters` to `false`.

## Development

Note to self: To push changes to NuGet.org, build a Release version and run 
`nuget push EpdIt.DbHelper.x.x.x.nupkg -Source https://www.nuget.org/api/v2/package`

### TO-DO

* Add unit tests for table-valued parameters (`TvpSqlParameter`)
