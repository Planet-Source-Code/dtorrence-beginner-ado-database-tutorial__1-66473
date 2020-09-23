VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Beginner DB Tutorial"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMove 
      Caption         =   ">|"
      Height          =   375
      Index           =   3
      Left            =   2880
      TabIndex        =   11
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   ">"
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   10
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "<"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   9
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "|<"
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtState 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtCity 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtLast 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtFirst 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "State"
      Height          =   495
      Index           =   3
      Left            =   2400
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "City"
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Last Name"
      Height          =   495
      Index           =   1
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "First Name"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                '''''''''''''''''''''''''''''''''
                '    ADO Beginner's Tutorial    '
                '        Derek Torrence         '
                '     Fork501@netscape.net      '
                '''''''''''''''''''''''''''''''''
'Before you begin, make sure you go up to Project-References
'and make sure that you have the two options selected for:
'1) Microsoft ActiveX Data Objects 2.8 Library
'2) Microsoft ActiveX Data Objects Recordset 2.8 Library

Private Sub UpdateFields()
'This function is set in place to update all of the textbox
'fields in our Program.  It will also be used as input
'validation to ensure that the user cannot move to a record
'in the wrong direction if they are at the beginning or end
'of the recordset.

    txtFirst = myRecords("FirstName")

    txtLast = myRecords("LastName")

    txtCity = myRecords("City")

    txtState = myRecords("State")

'Input Validation: Enable/Disable navigation buttons for user.

    If myRecords.AbsolutePosition = 1 Then

        cmdMove(0).Enabled = False

        cmdMove(1).Enabled = False

        cmdMove(2).Enabled = True

        cmdMove(3).Enabled = True

    ElseIf myRecords.AbsolutePosition = myRecords.RecordCount Then

        cmdMove(0).Enabled = True

        cmdMove(1).Enabled = True

        cmdMove(2).Enabled = False

        cmdMove(3).Enabled = False

    ElseIf myRecords.BOF = True And myRecords.EOF = True Then

    'This particular section will only arise if there is no inforamtion in
    'the user's database.

        MsgBox "ERROR: There is no information in your database!", vbCritical

        cmdMove(0).Enabled = False

        cmdMove(1).Enabled = False

        cmdMove(2).Enabled = False

        cmdMove(3).Enabled = False

    Else

        cmdMove(0).Enabled = True

        cmdMove(1).Enabled = True

        cmdMove(2).Enabled = True

        cmdMove(3).Enabled = True

    End If

'Please note: I am well aware that there is a more efficient way of handling this,
'but since this is only a beginner's tutorial, I want everything to be seen as
'clearly as possible!

End Sub

Private Sub cmdMove_Click(Index As Integer)
'I decided to use the command buttons as an indexed
'case function.  This makes things a little easier
'for learning.  There are obviously more ways to handle
'this, but I really like using this way.

    Select Case Index

        Case 0 '|<

            myRecords.MoveFirst

        Case 1 '<

            myRecords.MovePrevious

        Case 2 '>

            myRecords.MoveNext

        Case 3 '|<

            myRecords.MoveLast

    End Select

    UpdateFields 'This function is defined at the top of this page.

End Sub

Private Sub Form_Load()

    LoadDB ("People") 'This function is defined in modGlobals

    UpdateFields 'This function is defined at the top of this page.

End Sub


'There you have it!  A simple way to connect to your database and retrieve your
'information!  If you have any questions, comments or concerns, please don't
'hesitate to contact me!  I'm always here to help!
'
'                                                 ~Derek
