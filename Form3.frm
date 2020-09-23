VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Adding New Favorite"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   3165
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   2520
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   280
      TabIndex        =   4
      Top             =   2520
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
      Left            =   300
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1920
      Width           =   4000
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
      Left            =   300
      TabIndex        =   1
      Text            =   "Type New Category Here"
      Top             =   960
      Width           =   4000
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
      ItemData        =   "Form3.frx":0000
      Left            =   300
      List            =   "Form3.frx":0002
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   480
      Width           =   4000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Title of Category for Listing:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   300
      TabIndex        =   6
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Title for Favorite Listing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   300
      TabIndex        =   3
      Top             =   1560
      Width           =   3975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnFavorites As ADODB.Connection
Dim rsCats As ADODB.Recordset
Dim rsAddNew As ADODB.Recordset
Dim Loading As Boolean
Dim rsAddNew2 As ADODB.Recordset

Private Sub Combo1_Click()
    If Combo1.Text = "Other" Then
        Text1.Enabled = True
        Text1.Text = ""
        On Error Resume Next
        Text1.SetFocus
    Else
        Text1.Enabled = False
        Text1.Text = "Type New Category Here"
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
    Set rsAddNew = New ADODB.Recordset
    Set rsAddNew2 = New ADODB.Recordset
    Dim strSQLAdd As String
    If Text2.Text = "" Then
        MsgBox "Title field must be completed!", vbCritical
        Exit Sub
    End If
    If Text1.Text = "Type New Category Here" And Combo1.Text = "Other" Then
        MsgBox "Very Funny!  Type a real category name!", vbCritical
        Exit Sub
    End If
    If Text1.Text = "Type New Category Here" And Combo1.Text <> "Other" Then
        strSQLAdd = "INSERT INTO Favorites (URLAddress, Category, Title) VALUES ('" & Form1.Combo1.Text & "', '" & Combo1.Text & "', '" & Text2.Text & "') "
        rsAddNew.Open strSQLAdd, _
                cnFavorites, adOpenStatic, _
                adLockOptimistic
        Unload Me
    Else
        If Text1.Text <> "" Then
            strSQLAdd = "INSERT INTO Categories (Category, UserName) VALUES ('" & Text1.Text & "', '" & UserNames & "') "
            rsAddNew.Open strSQLAdd, _
                cnFavorites, adOpenStatic, _
                adLockOptimistic
            strSQLAdd = "INSERT INTO Favorites (URLAddress, Category, Title, UserName) VALUES ('" & Form1.Combo1.Text & "', '" & Text1.Text & "', '" & Text2.Text & "', '" & UserNames & "') "
            rsAddNew2.Open strSQLAdd, _
                cnFavorites, adOpenStatic, _
                adLockOptimistic
            Unload Me
        Else
            MsgBox "New Category Name Must Be Entered!", vbCritical
            Exit Sub
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Loading = False
End Sub

Private Sub Form_Load()
    Loading = True
    'Dim DescriptionParsed() As String
    'Dim x As Integer
    'Dim y As Integer
    'DescriptionParsed = Split(Form1.combo1.Text, ".")
    'y = UBound(DescriptionParsed)
    'If y >= 1 Then
    '    Text2.Text = DescriptionParsed(1)
    'Else
    '    Text2.Text = ""
    'End If
    Text2.Text = Form1.WebBrowser1.LocationName
    Set cnFavorites = New ADODB.Connection
    cnFavorites.Provider = "Microsoft.Jet.OLEDB.4.0;"
    cnFavorites.ConnectionString = "Persist Security Info = False;" _
       & "Data Source = Components/db1.mdb"
    cnFavorites.Open
    Set rsCats = New ADODB.Recordset
    Dim strSQLRetrieve As String
    strSQLRetrieve = "SELECT * FROM Categories WHERE Username Like '" & UserNames & "' ORDER BY Category"
    rsCats.Open strSQLRetrieve, _
            cnFavorites, adOpenStatic, _
            adLockOptimistic
    Do Until rsCats.EOF
        Combo1.AddItem rsCats!Category
        rsCats.MoveNext
    Loop
    Combo1.AddItem "Other"
    Combo1.ListIndex = 0
    DoEvents
End Sub

Private Sub Text1_Change()
   Dim temp As Boolean
    temp = Check_Apos(Text1.Text)
    If temp = True Then
        Text1.Text = ""
    End If
End Sub

Private Sub Text2_Change()
    Dim temp As Boolean
    temp = Check_Apos(Text2.Text)
    If temp = True Then
        Text2.Text = ""
        If Loading = True Then
            MsgBox "KenScape attempted to fill in the title of this site for you" & vbCrLf & "However, the site's title includes an invalid character so you will have to fill it in yourself." & vbCrLf & "The URL will be automatically filled in for you still, you just have to give it a name you will remember it by."
        End If
    End If
End Sub
