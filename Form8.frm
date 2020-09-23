VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "History Options"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   1080
      Picture         =   "Form8.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   3375
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   1080
      Picture         =   "Form8.frx":9520
      ScaleHeight     =   495
      ScaleWidth      =   3375
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit Changes"
      Default         =   -1  'True
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
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      ItemData        =   "Form8.frx":12990
      Left            =   2040
      List            =   "Form8.frx":129A3
      TabIndex        =   0
      Text            =   "10"
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "How many links would you like to remember?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnFavorites As ADODB.Connection
Dim rsCats As ADODB.Recordset

Private Sub Combo1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Picture1(0).Visible = True
    Picture1(1).Visible = False
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
    Set cnFavorites = New ADODB.Connection
    Set rsCats = New ADODB.Recordset
    Set cnFavorites = New ADODB.Connection
    cnFavorites.Provider = "Microsoft.Jet.OLEDB.4.0;"
    cnFavorites.ConnectionString = "Persist Security Info = False;" _
       & "Data Source = Components/db1.mdb"
    cnFavorites.Open
    Dim strSQL As String
    strSQL = "UPDATE Options SET HistorySaves = '" & Combo1.Text & "' WHERE UserName Like '" & UserNames & "' "
    rsCats.Open strSQL, _
        cnFavorites, adOpenStatic, _
        adLockOptimistic
    HistorySave = Combo1.Text
    Call Form1.Save_State
    Form1.Combo1.Clear
    Call Form1.Form_Load
    Unload Me
End Sub

Private Sub Form_Load()
    Combo1.Text = HistorySave
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1(0).Visible = True
    Picture1(1).Visible = False
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1(0).Visible = True
    Picture1(1).Visible = False
End Sub

Private Sub Picture1_Click(Index As Integer)
    If Index = 1 Then
        Call Command1_Click
    End If
End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Picture1(0).Visible = False
        Picture1(1).Visible = True
    End If
End Sub
