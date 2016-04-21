# Oracle Excel CRUD Interface

This is a small piece to allow you to insert a table directly into an Oracle database. It uses the column headers as the columns in the table. Primary keys are on the left hand side for matching and presently if you want a numeric value to be a primary key it needs to be saved as a text field(With the Number Stored as Text Tag)

Include This As a Button To Initialize A Session
```VBA
Private Sub BtnLogin_Click()
    frmLogin.Show
End Sub
```

This is an example use of the insert command
```VBA
Private Sub btnInsert_Click()
    If ValidSession = True Then
        ECSession.Insert "SWBWORD", Range("swbword[#All]"), True
    End If
End Sub
```

## Global Module API
#### ECSession:				ECSession

#### ValidSession	 => 		Boolean			
```
	Summary: This evaluates the current state of the application and returns a   
  Boolean Of whether or not you have a validated ECSession Login. Almost Every  
  Operation Requires This So it is a Recommended If Block On Any Command.  
```

#### ClearTable
```
	SheetName:		String
	TableName:			String

	Summary: This Clears a Table Object. It removes any background coloring and 
  resets the range to 1 beneath the headers
```

## ECSession Class API

##### Username : 			String
##### DSN:					String
##### Validated:				Boolean

#### Initialize
```
	Username:			String
	Password:			String
	DSN:				String
	
	Summary: This validates your connection ensuring that you can login and 
  operate in the banner environment. It stores your Username Password, 
  and DSN value for reuse. This is your only way to establish a connection.
  All values cannot contain single quotes or SQL Comment Strings.
```
#### Reset_Password
```
	txtNewPassword1: 	String
	txtNewPassword2:	String

	Summary: Utilizing an already validated session this will change the users 
  password. Requires that both new password are the same. Passwords cannot 
  contain single quotes or SQL Comment Strings.
```
#### Insert
```
	TableName:			String
	Selection:			Range
	ColorResults:		Boolean [Default = False]

	Summary: This subprocess takes a TableName as the table it will insert the 
  values into. It uses the first row as the columns that it will insert the 
  corresponding values into. ColorResults Changes whether or not to show 
  success or failure by coloring the far left cell.
```
#### Update
```
	TableName : 		String
	Selection: 			Range
	PrimaryKeys:		Integer [Default = 1]
	ColorResults:		Boolean [Default = False]

	Summary: The subprocess will update a table.

	Explanation: This subprocess takes a TableName as the table it will update 
  that table in our Banner Database. The first row of the range populates the
  column names for the table. These must match for the command to succeed. It
  then loops through all data in the table and generates a Set Clause, and 
  depending on the Number of Primary Keys Entered it will generate that many
  statement in the where clause, always starting from the leftmost column. 
  ColorResults Changes whether or not to show success or failure by coloring 
  the far left cell.
```
#### Delete
```
	TableName: 		String
	Selection:			Range
	PrimaryKeys:		Integer [Default = 1]
	ColorResults:		Boolean [Default = False]

	Summary: This subprocess will removed the entries in it from a table.

	Explanation: This subprocess takes a TableName as the table it will delete  
  entries from that table in our Banner Database. The where clause is populated  
  based on how many primary keys are indicated by values on the left side.   
  ColorResults Changes whether or not to show success or failure by coloring  
  the far left cell.  
```
#### StoredProcedure
```
	ProcedureName:	String
	Arguments:			String [Default = “”]

	Summary: This will execute a Stored procedure and has a package body so that 
  you can bring in other elements to execute the procedure if you would like 
  as well.
```
#### UpdateHardCode
```
	TableName:			String
	SetStatement:		String
	WhereStatement:	String

	Summary: This subprocess updates a table based on an arbitrary update command. 

	Explanation: 
```
#### DeleteHardCode
```
	TableName: 		String
	WhereStatement:	String

	Summary: This subprocess deletes entries from a table based on an arbitrary 
  Where statement.
```