VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Homepage Options"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   7
      Left            =   2520
      Picture         =   "Form5.frx":0000
      ScaleHeight     =   735
      ScaleWidth      =   2775
      TabIndex        =   12
      Top             =   600
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   6
      Left            =   2520
      Picture         =   "Form5.frx":9630
      ScaleHeight     =   735
      ScaleWidth      =   2775
      TabIndex        =   11
      Top             =   600
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   5
      Left            =   2520
      Picture         =   "Form5.frx":12C60
      ScaleHeight     =   735
      ScaleWidth      =   2775
      TabIndex        =   10
      Top             =   1440
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   4
      Left            =   2520
      Picture         =   "Form5.frx":1C22C
      ScaleHeight     =   735
      ScaleWidth      =   2775
      TabIndex        =   9
      Top             =   1440
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   3
      Left            =   600
      Picture         =   "Form5.frx":257F8
      ScaleHeight     =   615
      ScaleWidth      =   1335
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   2
      Left            =   600
      Picture         =   "Form5.frx":2D604
      ScaleHeight     =   615
      ScaleWidth      =   1335
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   1
      Left            =   120
      Picture         =   "Form5.frx":35410
      ScaleHeight     =   735
      ScaleWidth      =   2055
      TabIndex        =   6
      Top             =   600
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   0
      Left            =   120
      Picture         =   "Form5.frx":3D592
      ScaleHeight     =   735
      ScaleWidth      =   2055
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Last Visited Page"
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   1800
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
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
      Left            =   50
      TabIndex        =   3
      Top             =   120
      Width           =   5415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Use Current Page"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   2760
      Width           =   1005
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Random"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   1200
      Width           =   645
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set to Input"
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   2760
      Width           =   1600
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnFavorites As ADODB.Connection
Dim rsCats As ADODB.Recordset

Private Sub Save_Home()
    Set cnFavorites = New ADODB.Connection
    Set rsCats = New ADODB.Recordset
    Set cnFavorites = New ADODB.Connection
    cnFavorites.Provider = "Microsoft.Jet.OLEDB.4.0;"
    cnFavorites.ConnectionString = "Persist Security Info = False;" _
       & "Data Source = Components/db1.mdb"
    cnFavorites.Open
    Dim strSQL As String
    strSQL = "UPDATE Options SET Homepage = '" & HomePage & "' WHERE UserName Like '" & UserNames & "' "
    rsCats.Open strSQL, _
        cnFavorites, adOpenStatic, _
        adLockOptimistic
    Call Form1.Save_State
    Unload Me
End Sub

Private Sub Command1_Click()
    HomePage = Text1.Text
    Call Save_Home
End Sub

Private Sub Command2_Click()
    HomePage = "Random"
    Call Save_Home
End Sub

Private Sub Command3_Click()
    HomePage = Form1.Combo1.Text
    Call Save_Home
End Sub

Private Sub Command4_Click()
    HomePage = "Last"
    Call Save_Home
End Sub

Private Sub Form_Load()
    Picture1(1).Visible = False
    Picture1(3).Visible = False
    Picture1(5).Visible = False
    Picture1(7).Visible = False
    Picture1(0).Visible = True
    Picture1(2).Visible = True
    Picture1(4).Visible = True
    Picture1(6).Visible = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_Load
End Sub

Private Sub Picture1_Click(Index As Integer)
    If Index = 1 Then
        Call Command1_Click
    ElseIf Index = 3 Then
        Call Command2_Click
    ElseIf Index = 5 Then
        Call Command3_Click
    ElseIf Index = 7 Then
        Call Command4_Click
    End If
End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index Mod 2 = 0 Then
        Picture1(Index).Visible = False
        Picture1(Index + 1).Visible = True
    End If
End Sub

Private Sub Text1_Change()
    Dim temp As Boolean
    temp = Check_Apos(Text1.Text)
    If temp = True Then
        Text1.Text = ""
    End If
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_Load
End Sub
