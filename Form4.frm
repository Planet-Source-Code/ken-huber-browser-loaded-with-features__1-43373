VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "KenScape Navigator Login"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3405
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   2175
   ScaleWidth      =   3405
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Delete User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "NEW USER NAME"
      Top             =   600
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Log-In"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   2895
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnFavorites As ADODB.Connection
Dim rsCats As ADODB.Recordset

Private Sub Combo1_Click()
    If Combo1.Text = "New User" Then
        Text1.Text = ""
        Text1.Enabled = True
        Text1.SetFocus
    Else
        Text1.Text = "NEW USER NAME"
        Text1.Enabled = False
    End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    Combo1.ListIndex = 0
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    Combo1.ListIndex = 0
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    Combo1.ListIndex = 0
End Sub

Private Sub Command1_Click()
    If Text1.Enabled = False Then
        UserNames = Combo1.Text
    Else
        Set rsCats = New ADODB.Recordset
        Dim strSQL As String
        strSQL = "INSERT INTO Options (UserName, HistorySaves) VALUES ('" & Text1.Text & "', 10) "
        rsCats.Open strSQL, _
            cnFavorites, adOpenStatic, _
            adLockOptimistic
        UserNames = Text1.Text
    End If
    Form4.Hide
    Unload Form6
    Form1.Show
End Sub

Private Sub Command2_Click()
    If Text1.Enabled = False Then
        Dim answer As Integer
        answer = MsgBox("Are you sure you wish to delete this user?", vbYesNo)
        If answer = vbYes Then
            Set rsCats = New ADODB.Recordset
            Dim strSQL As String
            strSQL = "DELETE FROM Options WHERE UserName Like '" & Combo1.Text & "' "
            rsCats.Open strSQL, _
                cnFavorites, adOpenStatic, _
                adLockOptimistic
            Call Form_Load
        End If
    Else
        MsgBox "Not possible to delete someone that does not exist yet!", vbCritical
    End If
End Sub

Private Sub Form_Load()
    Combo1.Clear
    Set cnFavorites = New ADODB.Connection
    Set rsCats = New ADODB.Recordset
    Set cnFavorites = New ADODB.Connection
    cnFavorites.Provider = "Microsoft.Jet.OLEDB.4.0;"
    cnFavorites.ConnectionString = "Persist Security Info = False;" _
       & "Data Source = Components/db1.mdb"
    cnFavorites.Open
    Dim strSQL As String
    strSQL = "SELECT * FROM Options"
        rsCats.Open strSQL, _
            cnFavorites, adOpenStatic, _
            adLockOptimistic
    Do Until rsCats.EOF
        Combo1.AddItem rsCats!UserName
        rsCats.MoveNext
    Loop
    Combo1.AddItem "New User"
    Combo1.ListIndex = 0
End Sub

Private Sub Text1_Change()
    Dim temp As Boolean
    temp = Check_Apos(Text1.Text)
    If temp = True Then
        Text1.Text = ""
    End If
End Sub
