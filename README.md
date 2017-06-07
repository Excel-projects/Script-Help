<h1 align="center">
  <img src="Images/toolbar.png" alt="MyToolbar" />
</h1>

# ScriptHelp
This is an Excel 2010 VSTO Addin written in Visual Studio 2013 C#. It allows the user to use an Excel table to create different SQL scripts.



<h1 align="center">
  <img src="Images/script_toolbar.png" alt="MyToolbar" />
</h1>

*	T-SQL Insert Values – This menu item will format the script column to use individual insert statements
*	T-SQL Merge Values – This menu item will format the script column to use a merge statement with a select values
*	T-SQL Select Values – This menu item will format the script column to be used in insert statements 
*	T-SQL Select Union – This menu item will format the script column to be used in an update statement 
*	T-SQL Update Values – This menu item will format the script column to use individual update statements
*	PL/SQL Insert Values – This menu item will format the script column to use individual insert statements
*	PL/SQL Select Union – This menu item will format the script column to be used in an update statement 
*	PL/SQL Update Values – This menu item will format the script column to use individual update statements
*	DQL Append – This menu item will format the script column to be used in an append statement for Documentum (this is used for repeating values)
*	DQL Append/Locked – This menu item will format the script column to be used in an append statement for Documentum (this is used for repeating values) and unlocks and then locks the record
*	DQL Create – This menu item will format the script column to be used in an create statement for Documentum
*	DQL Truncate/Append – This menu item will format the script column to be used in an truncate and then append statement for Documentum (this is used for repeating values)
*	DQL Update – This menu item will format the script column to be used in an update statement for Documentum
*	DQL Update/Locked – This menu item will format the script column to be used in an update statement for Documentum and unlocks and then locks the record
*	Add “WHERE” before the column name in the header you want to use as criteria.
*	The “Table Alias” is used as the update table name
This window will pop-up on “Add Script Column” click. “Save to File” from the toolbar will save the text to a .dql file.
