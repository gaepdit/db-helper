# EPD-IT DB Helper Class Library

The purpose of this library is to simplify interactions with a SQL Server database. 

## Status

This library was originally written for working with an Oracle database. It is being migrated to work with SQL Server.

This is still a work in progress. Some parts are working and some are not. Please contribute and help fix the parts that don't work yet. 

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
Friend DB As New EpdItDbHelper.DB(ConnectionString)
```

The DB object has many functions available for interacting with the database. Some functions which require a SQL query or command and optional SQL parameters.

Example:

```
Dim userId as Integer = 3
Dim query as String = "select UserName from UserTable where UserId = @id"
Dim parameter As SqlParameter = New SqlParameter("id", userId)
Dim userName as String = DB.GetSingleValue(Of String)(query, parameter)
```

Other functions require a name of a Stored Procedure or User-defined Function instead of a SQL query. These functions all start with "SP" in the name.

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

Start by writing Unit Tests for any functions not yet covered. Then run the tests and fix any code that is broken.