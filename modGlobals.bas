Attribute VB_Name = "modGlobals"
                '''''''''''''''''''''''''''''''''
                '    ADO Beginner's Tutorial    '
                '        Derek Torrence         '
                '     Fork501@netscape.net      '
                '''''''''''''''''''''''''''''''''
Public myConnection As ADODB.Connection 'This object will hold your connection
Public myRecords As ADODB.Recordset 'This object will hold your recordset

Public Sub LoadDB(TableName As String)
Dim strConnect As String 'This string holds our connection string.
Dim strDbLocation As String 'This string holds the location of our database.

'Before we can begin opening the database, we must first tell our
'Program what we will use to actually open the database.

'This is where we define myConnection and myRecords objects

    Set myConnection = New ADODB.Connection

    Set myRecords = New ADODB.Recordset

    strDbLocation = App.Path & "\dbTutorial.mdb"

    strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDbLocation
    'Don't get scared by this!  If you can't remember (I can't, either) what to
    'type for all of this, just visit http://www.connectionstrings.com/ and click
    'on "Access" followed by "OLE DB, OleDbConnection (.NET)" it will be the
    'section called "Standard Security" .. just take off the last part, which
    'says, ";Password=;"

'After we have defined our objects and strings, we must now tell our
'Program that it needs to open our database.

    myConnection.Open strConnect

    myRecords.Open "SELECT * FROM " & TableName, myConnection, adOpenKeyset, adLockPessimistic
    'A brief explanation about this statement:
    '
    '"SELECT * FROM " & TableName <~This basically is telling our program to use
    'an SQL statement to find all of the records in the table name, which we
    'specify in our code.  Nothing to be scared of; it is the most basic SQL statement
    'in use today.
    '
    'myConnection <~This is telling our recordset which connection to access our table
    'through.  Just think of it as a network cable connecting a computer to a switch.
    '
    'adOpenKeyset <~This enables you to move forward and backward through your records.
    'This is a very important option.
    '
    'adLockPessimistic <~Theoretically, this lets you update records.  However, I use
    'SQL to update all of my records, so it is useless to me.  I include it there just
    'for good nature.

End Sub


'There you have it!  A simple way to connect to your database and retrieve your
'information!  If you have any questions, comments or concerns, please don't
'hesitate to contact me!  I'm always here to help!
'
'                                                 ~Derek

