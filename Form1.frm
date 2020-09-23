VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "KenScape Navigator 2.0"
   ClientHeight    =   7620
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10830
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0442
   ScaleHeight     =   7620
   ScaleWidth      =   10830
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   50
      TabIndex        =   18
      Top             =   50
      Width           =   7095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Home"
      Height          =   255
      Left            =   4680
      TabIndex        =   17
      Top             =   650
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      Height          =   950
      Left            =   9840
      Picture         =   "Form1.frx":05F4
      ScaleHeight     =   885
      ScaleWidth      =   795
      TabIndex        =   16
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   950
      Index           =   3
      Left            =   9840
      Picture         =   "Form1.frx":E4BE
      ScaleHeight     =   885
      ScaleWidth      =   795
      TabIndex        =   15
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   950
      Index           =   2
      Left            =   9840
      Picture         =   "Form1.frx":19CDA
      ScaleHeight     =   885
      ScaleWidth      =   795
      TabIndex        =   14
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   950
      Index           =   1
      Left            =   9840
      Picture         =   "Form1.frx":254F6
      ScaleHeight     =   885
      ScaleWidth      =   795
      TabIndex        =   13
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Height          =   250
      Index           =   3
      Left            =   50
      Picture         =   "Form1.frx":2BDE6
      ScaleHeight     =   195
      ScaleWidth      =   2955
      TabIndex        =   12
      Top             =   1025
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   250
      Index           =   2
      Left            =   50
      Picture         =   "Form1.frx":2E0E0
      ScaleHeight     =   195
      ScaleWidth      =   2955
      TabIndex        =   11
      Top             =   1025
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   250
      Index           =   1
      Left            =   50
      Picture         =   "Form1.frx":303DA
      ScaleHeight     =   195
      ScaleWidth      =   2955
      TabIndex        =   10
      Top             =   1025
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   250
      Index           =   0
      Left            =   50
      Picture         =   "Form1.frx":32620
      ScaleHeight     =   195
      ScaleWidth      =   2955
      TabIndex        =   9
      Top             =   1025
      Width           =   3015
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   50
      TabIndex        =   8
      Top             =   50
      Width           =   9735
   End
   Begin VB.PictureBox Picture1 
      Height          =   950
      Index           =   0
      Left            =   9840
      Picture         =   "Form1.frx":348A2
      ScaleHeight     =   885
      ScaleWidth      =   795
      TabIndex        =   7
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   650
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   650
      Width           =   975
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   "Forward"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   650
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   650
      Width           =   975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO!"
      Default         =   -1  'True
      Height          =   255
      Left            =   50
      TabIndex        =   0
      Top             =   650
      Width           =   855
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6615
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   10815
      ExtentX         =   19076
      ExtentY         =   11668
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KenScape Navigator"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   50
      TabIndex        =   1
      Top             =   360
      Width           =   9735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "New Window"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsAllow 
         Caption         =   "Pop-Ups Not Allowed"
      End
      Begin VB.Menu mnuFavOpt 
         Caption         =   "Favorites Start Expanded"
      End
      Begin VB.Menu mnuHomeOpt 
         Caption         =   "Homepage Options"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "History Options"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuAutoFill 
         Caption         =   "Automatically Add .com"
      End
   End
   Begin VB.Menu mnuFavorites 
      Caption         =   "Favorites"
      Begin VB.Menu mnuFav 
         Caption         =   "View Favorites"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add to Favorites"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuInstructions 
         Caption         =   "Instructions"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About KenScape"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'                            Coded by: Ken Huber
'***************************************************************************
Option Explicit
Dim z As Integer
Dim cnFavorites As ADODB.Connection
Dim rsCats As ADODB.Recordset
Dim rsCats2 As ADODB.Recordset
Dim rsCats3 As ADODB.Recordset
Dim rsCats4 As ADODB.Recordset
Dim rsCats5 As ADODB.Recordset
Dim rsCats6 As ADODB.Recordset
Dim rsCats7 As ADODB.Recordset
Dim rsCats8 As ADODB.Recordset
Public AllowPopup As Boolean 'This is for Pop-up windows

Private Sub cmdBack_Click()
'Go back one page
WebBrowser1.GoBack
End Sub

Private Sub cmdForward_Click()
    'go forward one page
    On Error Resume Next
    WebBrowser1.GoForward
End Sub

Public Sub cmdGo_Click()
    'Go to web page
    If mnuAutoFill.Caption = "Automatically Add .com" Then
        Dim temp As Boolean
        temp = CheckStr(Combo1.Text, ".")
        If temp = False Then
            Combo1.Text = Combo1.Text & ".com"
        End If
    End If
    On Error Resume Next
    WebBrowser1.Navigate Combo1.Text
    lblStatus.Caption = "Going to: " & Combo1.Text
End Sub

Private Sub cmdRefresh_Click()
'Refresh page
WebBrowser1.Refresh
End Sub

Private Sub cmdStop_Click()
'Stop loading
WebBrowser1.Stop
End Sub

Private Sub Combo1_Change()
    Dim temp As Boolean
    temp = Check_Apos(Combo1.Text)
    If temp = True Then
        Combo1.ListIndex = 0
    End If
End Sub

Private Sub Combo1_Click()
    If Combo1.ListIndex >= 0 Then
        Call cmdGo_Click
    End If
End Sub

Private Sub Command1_Click()
    Call Save_State
    Combo1.Clear
    Call Form_Load
End Sub

Public Sub Form_Load()
    Set rsCats4 = New ADODB.Recordset
    Set rsCats7 = New ADODB.Recordset
    Set cnFavorites = New ADODB.Connection
    Set rsCats = New ADODB.Recordset
    Set cnFavorites = New ADODB.Connection
    cnFavorites.Provider = "Microsoft.Jet.OLEDB.4.0;"
    cnFavorites.ConnectionString = "Persist Security Info = False;" _
       & "Data Source = Components/db1.mdb"
    cnFavorites.Open
    Dim strSQL As String
    strSQL = "SELECT * FROM Options Where UserName Like '" & UserNames & "' "
    rsCats.Open strSQL, _
        cnFavorites, adOpenStatic, _
        adLockOptimistic
    If rsCats!favoritesStart = "Condensed" Then
        Call mnuFavOpt_Click
    End If
    If rsCats!popups = "Allowed" Then
        Call mnuOptionsAllow_Click
    End If
    If rsCats!HomePage = "Random" Then
        Call RandomCall
    ElseIf rsCats!HomePage = "" Or IsNull(rsCats!HomePage) = True Then
        Combo1.Text = "www.geocities.com/khuber1"
    ElseIf rsCats!HomePage <> "Last" Then
        Combo1.Text = rsCats!HomePage
    End If
    If IsNull(rsCats!AddCom) = True Then
        mnuAutoFill.Caption = "Automatically Add .com"
    Else
        mnuAutoFill.Caption = rsCats!AddCom
    End If
'Resize and place objects
With WebBrowser1
    .Width = Form1.Width - 200
    .Left = 50
    .Height = Form1.Height - 200
End With
With lblStatus
    .FontBold = True
    .Width = WebBrowser1.Width
End With
Combo1.Left = 50
    strSQL = "SELECT * FROM HistoryOptions Where UserName Like '" & UserNames & "' "
    
    rsCats4.Open strSQL, _
        cnFavorites, adOpenStatic, _
        adLockOptimistic
    If rsCats!HomePage = "Last" Then
        Combo1.Text = rsCats4!Link
    End If
    strSQL = "SELECT * FROM Options Where UserName Like '" & UserNames & "' "
    rsCats7.Open strSQL, _
        cnFavorites, adOpenStatic, _
        adLockOptimistic
    HistorySave = rsCats7!HistorySaves
    Dim xxx As Integer
    Do Until xxx >= HistorySave Or rsCats4.EOF
        Combo1.AddItem rsCats4!Link, 0
        rsCats4.MoveNext
        xxx = xxx + 1
    Loop
    
    
    Call cmdGo_Click
End Sub

Private Sub Form_Resize()
'Resizes everything to fit to the form
With WebBrowser1
    If Form1.Width > 200 Then
       .Width = Form1.Width - 200
    Else
        .Width = 0
    End If
    If Form1.Height > 1700 Then
        .Height = Form1.Height - 1700
    Else
        .Height = 0
    End If
End With
With lblStatus
    If Form1.Width > 1200 Then
        Combo1.Width = Form1.Width - 1200
        .Width = Form1.Width - 1200
    Else
        .Width = 0
        Combo1.Width = 0
    End If
End With
Dim X As Integer
For X = 0 To 3
    Picture1(X).Left = lblStatus.Left + lblStatus.Width + 50
    Picture3.Left = Picture1(X).Left
Next X

If Form1.Width > 1200 Then
    cmdGo.Left = 50
    cmdGo.Width = (Form1.Width - 1200) / 6
    cmdBack.Left = cmdGo.Left + cmdGo.Width
    cmdBack.Width = (Form1.Width - 1200) / 6
    cmdForward.Width = (Form1.Width - 1200) / 6
    cmdForward.Left = cmdBack.Left + cmdBack.Width
    cmdStop.Width = (Form1.Width - 1200) / 6
    cmdStop.Left = cmdForward.Left + cmdForward.Width
    cmdRefresh.Width = (Form1.Width - 1200) / 6
    cmdRefresh.Left = cmdStop.Left + cmdStop.Width
    Command1.Width = (Form1.Width - 1200) / 6
    Command1.Left = cmdRefresh.Left + cmdRefresh.Width
Else
    cmdGo.Width = 0
    cmdBack.Width = 0
    cmdForward = 0
    cmdStop = 0
    cmdRefresh = 0
End If
End Sub
Private Sub RandomCall()
        Set rsCats2 = New ADODB.Recordset
        Dim strSQL As String
        strSQL = "SELECT * FROM Favorites WHERE Username Like '" & UserNames & "' "
        rsCats2.Open strSQL, _
            cnFavorites, adOpenStatic, _
            adLockOptimistic
        Dim temp As Long
        If rsCats2.RecordCount > 1 Then
            Randomize
            temp = Rnd() * (rsCats2.RecordCount - 1)
        Else
            MsgBox "You need 2 or more favorites to have your homepage set to random!", vbCritical
        End If
        Dim X As Integer
        Dim Y As Integer
        Y = temp
        Do Until X >= Y
            X = X + 1
            rsCats2.MoveNext
        Loop
        Combo1.Text = rsCats2!URLAddress
End Sub
Private Sub Form_Terminate()
    Call mnuexit_Click
End Sub

Private Sub mnuAbout_Click()
MsgBox "KenScape Navigator 2.0" & vbCrLf & "Coded by: Ken Huber", vbOKOnly, "About KenScape Navigator"
End Sub

Private Sub mnuAdd_Click()
    Form3.Show (vbModal)
End Sub

Private Sub mnuAutoFill_Click()
    If mnuAutoFill.Caption = "Automatically Add .com" Then
        mnuAutoFill.Caption = "Do not add .com"
    Else
        mnuAutoFill.Caption = "Automatically Add .com"
    End If
    Call Save_State
End Sub

Private Sub mnuexit_Click()
    Call Save_State
    DoEvents
    End
End Sub

Public Sub Save_State()
    Set rsCats8 = New ADODB.Recordset
    Set rsCats3 = New ADODB.Recordset
    Set rsCats5 = New ADODB.Recordset
    Dim strSQL As String
    strSQL = "UPDATE Options SET AddCom = '" & mnuAutoFill.Caption & "' WHERE UserName Like '" & UserNames & "' "
    rsCats8.Open strSQL, _
        cnFavorites, adOpenStatic, _
        adLockOptimistic
    If mnuFavOpt.Caption = "Favorites Start Condensed" Then
        strSQL = "UPDATE Options SET FavoritesStart = 'Condensed' WHERE UserName Like '" & UserNames & "' "
    Else
        strSQL = "UPDATE Options SET FavoritesStart = 'Expanded' WHERE UserName Like '" & UserNames & "' "
    End If
    rsCats3.Open strSQL, _
        cnFavorites, adOpenStatic, _
        adLockOptimistic
    
    Set rsCats3 = New ADODB.Recordset
    If mnuOptionsAllow.Caption = "Pop-Ups Allowed" Then
        strSQL = "UPDATE Options SET Popups = 'Allowed' WHERE UserName Like '" & UserNames & "' "
    Else
        strSQL = "UPDATE Options SET Popups = 'NotAllowed' WHERE UserName Like '" & UserNames & "' "
    End If
    rsCats3.Open strSQL, _
        cnFavorites, adOpenStatic, _
        adLockOptimistic
    strSQL = "DELETE FROM HistoryOptions WHERE UserName Like '" & UserNames & "' "
    rsCats5.Open strSQL, _
        cnFavorites, adOpenStatic, _
        adLockOptimistic
    Dim X As Integer
    If Combo1.ListCount > 0 Then
    Do Until X = (Combo1.ListCount - 1)
        Set rsCats6 = New ADODB.Recordset
        Combo1.ListIndex = X
        strSQL = "INSERT INTO HistoryOptions (Link, UserName) VALUES ('" & Combo1.Text & "', '" & UserNames & "') "
        rsCats6.Open strSQL, _
            cnFavorites, adOpenStatic, _
            adLockOptimistic
        X = X + 1
    Loop
    End If
    Exit Sub
End Sub
Private Sub mnuFav_Click()
    If mnuFavOpt.Caption = "Favorites Start Condensed" Then
        Expanded = False
    Else
        Expanded = True
    End If
    Form2.Show (vbModal)
End Sub

Private Sub mnuFavOpt_Click()
If mnuFavOpt.Caption = "Favorites Start Condensed" Then
    mnuFavOpt.Caption = "Favorites Start Expanded"
    lblStatus.Caption = "KenScape Navigator - Favorites Start Expanded"
    Expanded = True
ElseIf mnuFavOpt.Caption = "Favorites Start Expanded" Then
    mnuFavOpt.Caption = "Favorites Start Condensed"
    lblStatus.Caption = "KenScape Navigator - Favorites Start Condensed"
    Expanded = False
End If
Call Save_State
End Sub

Private Sub mnuHistory_Click()
    Form8.Show (vbModal)
End Sub

Private Sub mnuHomeOpt_Click()
    Form5.Show (vbModal)
End Sub

Private Sub mnuInstructions_Click()
    Combo1.Text = "C:\Program Files\KenScape\Components\Instructions.html"
    Call cmdGo_Click
End Sub

Private Sub mnuOptionsAllow_Click()
'Turn on/off pop-up windows
If mnuOptionsAllow.Caption = "Pop-Ups Allowed" Then
    mnuOptionsAllow.Caption = "Pop-Ups Not Allowed"
    lblStatus.Caption = "KenScape Navigator - Blocking Pop-Up windows"
    AllowPopup = False
ElseIf mnuOptionsAllow.Caption = "Pop-Ups Not Allowed" Then
    mnuOptionsAllow.Caption = "Pop-Ups Allowed"
    lblStatus.Caption = "KenScape Navigator - Allowing Pop-Up windows"
    AllowPopup = True
End If
Call Save_State
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    'shows done in the status bar
    lblStatus.Caption = "Done Loading"
    Form1.Caption = "KenScape Navigator - " & WebBrowser1.LocationName
    Combo1.Text = WebBrowser1.LocationURL
    Combo1.AddItem Combo1.Text, 0
    If Combo1.ListCount > HistorySave Then
        Combo1.RemoveItem (Combo1.ListCount - 1)
    End If
End Sub

Private Sub WebBrowser1_DownloadBegin()
    'Starting download
    lblStatus.Caption = "Starting Download"
    
End Sub

Private Sub WebBrowser1_DownloadComplete()
    'Done downloading
    
    lblStatus.Caption = "Download Done!"
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    'Loaded page
    lblStatus.Caption = "Done Loading!"
    'Combo1.Text = WebBrowser1.LocationURL
    'Combo1.AddItem Combo1.Text, 0
    'If Combo1.ListCount > HistorySave Then
    '    Combo1.RemoveItem (Combo1.ListCount - 1)
    'End If
    Form1.Caption = "KenScape Navigator - " & WebBrowser1.LocationName  'Shows webpage in title bar
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
'This will allow a pop-up window to load or to be blocked!
If AllowPopup = True Then
    Cancel = False
    DoEvents
ElseIf AllowPopup = False Then
    Cancel = True
End If
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
'Shows progress in status bar
lblStatus.Caption = "Reading " & Progress & "  of  " & ProgressMax
Dim X As Integer
Dim Y As Long
If z = 3 Then
    z = 0
Else
    z = z + 1
End If
If ProgressMax > 0 And Progress > 0 Then
    
    Picture2(z).Visible = True
    Picture1(z).Visible = True
    Picture3.Visible = False
    If z = 0 Then
        Picture2(3).Visible = False
        Picture1(3).Visible = False
    Else
        Picture2(z - 1).Visible = False
        Picture1(z - 1).Visible = False
    End If
    Y = (Form1.Width - 200) / ProgressMax
For X = 0 To 3
    Picture2(X).Width = Progress * Y
Next X
Else
    Picture3.Visible = True
    Picture2(0).Visible = False
    Picture2(1).Visible = False
    Picture2(2).Visible = False
    Picture2(3).Visible = False
    Picture1(0).Visible = False
    Picture1(1).Visible = False
    Picture1(2).Visible = False
    Picture1(3).Visible = False
End If
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
'shows new text in status bar
lblStatus.Caption = Text
End Sub

Function FileExist(vFile As String) As Boolean
    On Error Resume Next
    FileExist = False
    If Dir$(vFile) <> "" Then: FileExist = True
End Function
