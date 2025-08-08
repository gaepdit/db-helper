# DB Helper

The purpose of this library is to simplify interactions with a SQL Server database. 

[![NuGet](https://img.shields.io/nuget/v/GaEpd.DbHelper.svg?maxAge=86400)](https://www.nuget.org/packages/GaEpd.DbHelper/) 

[![.NET Test](https://github.com/gaepdit/db-helper/actions/workflows/dotnet-test.yml/badge.svg)](https://github.com/gaepdit/enforcement-orders/actions/workflows/dotnet.yml)
[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=gaepdit_db-helper&metric=alert_status)](https://sonarcloud.io/summary/new_code?id=gaepdit_db-helper)
[![Lines of Code](https://sonarcloud.io/api/project_badges/measure?project=gaepdit_db-helper&metric=ncloc)](https://sonarcloud.io/summary/new_code?id=gaepdit_db-helper)

## What is this?

There are two classes available:

* The `DBHelper` class has many functions available for interacting with a SQL Server database. These functions take care of the tedious parts of querying data or running commands on the database, such as handling connection objects and transactions. There are separate sets of functions depending on whether a SQL query string is used as input or the name of a stored procedure.

* There is also a `DBUtilities` class with some static functions that simplify working with DBNull values and table-valued SQL parameters.

## How do I install it?

To install DB Helper, search for "GaEpd.DbHelper" in the NuGet package manager or run the following command in the [Package Manager Console](https://docs.nuget.org/consume/package-manager-console):

`PM> Install-Package GaEpd.DbHelper`

## Using the `DBHelper` class

The `DBHelper` class must be instantiated with a connection string:

```vb
Public DB As New GaEpd.DBHelper(connectionString)
```

Optionally, you can configure [connection retry logic](https://learn.microsoft.com/en-us/sql/connect/ado-net/configurable-retry-logic-sqlclient-introduction) and provide it to the `DBHelper` constructor:

```vb
Dim options As New SqlRetryLogicOption() With {
    .NumberOfTries = 3,
    .DeltaTime = TimeSpan.FromSeconds(5),
    .MaxTimeInterval = TimeSpan.FromSeconds(15),
    .AuthorizedSqlCondition = Function(x) String.IsNullOrEmpty(x) OrElse Regex.IsMatch(x, "\b(SELECT)\b", RegexOptions.IgnoreCase)
}

connectionRetryProvider = SqlConfigurableRetryFactory.CreateFixedRetryProvider(options)

DB = New GaEpd.DBHelper(connectionString, connectionRetryProvider)
```

### Query string functions

These functions require a SQL query or command plus an optional SQL parameter or array of SQL parameters. When running a non-query command, such as `INSERT` or `UPDATE`, an optional output parameter will contain the number of rows affected.

**Example 1:** Simple query

```vb
Dim query as String = "select States from StatesTable"
Dim states as DataTable = DB.GetDataTable(query)
```

**Example 2:** Query with one SQL parameter

```vb
Dim query as String = "select UserName from UserTable where UserId = @id"
Dim parameter As New SqlParameter("@id", MyUserId)
Dim userName as String = DB.GetSingleValue(Of String)(query, parameter)
```

**Example 3:** Command with multiple SQL parameters

```vb
Dim query as String = "update FacilityTable set Name = @name where FacilityId = @id"
Dim parameterArray As SqlParameter() = {
    New SqlParameter("@name", MyNewFacilityName),
    New SqlParameter("@id", MyFacilityId)
}
Dim result as Boolean = DB.RunCommand(query, parameterArray)
```

**Example 4:** Count rows affected by command

```vb
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

```vb
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

```vb
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

## Changelog

### Breaking changes

#### Version 5.1

The SqlClient package was migrated from System.Data.SqlClient to the new [Microsoft.Data.SqlClient](https://www.nuget.org/packages/Microsoft.Data.SqlClient/). This update introduces some breaking changes. 

Please see the [Microsoft announcement](https://github.com/dotnet/announcements/issues/322) and the [porting cheat sheet](https://github.com/dotnet/SqlClient/blob/main/porting-cheat-sheet.md) for more information and guidance on updating your code.

#### Version 5

The namespace/prefix was changed from `EpdIt` to `GaEpd`, which is a reserved prefix on nuget.org. This required publishing as a new package and deprecating the old package.

#### Version 3

The `forceAddNullableParameters` parameter has been removed. `DBNull.Value` will be sent for `SqlParameter`'s that evaluate to null (`Nothing` in VB.NET).

The output parameter convenience functions have been removed. If you need an output SQL parameter, just add it as you would normally add any other parameter.

#### Version 2

The `forceAddNullableParameters` parameter now defaults to `true`. If this parameter is not set (or is manually set to `true`), then `DBNull.Value` will be sent for `SqlParameter`'s that evaluate to `Nothing`. To return to the default behavior of dropping such parameters, you must manually set `forceAddNullableParameters` to `false`.

## Developer Notes

To publish a package update to NuGet.org, build a Release version, navigate to the project folder, and run:

```
nuget push GaEpd.DbHelper.x.x.x.nupkg -Source https://api.nuget.org/v3/index.json
```
