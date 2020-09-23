VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Favorites"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6720
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete from Favorites"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go!"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7858
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   16761024
      BackColorFixed  =   16761024
      BackColorSel    =   16761024
      BackColorBkg    =   16761024
      GridColor       =   16761024
      GridColorFixed  =   16761024
      ScrollBars      =   2
      BorderStyle     =   0
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsFavs As ADODB.Recordset
Dim rsFavs2 As ADODB.Recordset
Dim cnFavorites As ADODB.Connection
Dim rsFavs3 As ADODB.Recordset
Dim rsCats As ADODB.Recordset
Dim RH As Integer
Dim UnderL As Boolean
Private Sub Command1_Click()
    If MSFlexGrid1.CellFontBold = False Then
        MSFlexGrid1.Col = 1
        Form1.Combo1.Text = MSFlexGrid1.Text
        Call Form1.cmdGo_Click
        Unload Me
    Else
        MsgBox ("'" & MSFlexGrid1.Text & "' is one of your categories, not a valid URL.")
    End If
End Sub
Private Sub Command2_Click()
    Dim answer As Integer
    If MSFlexGrid1.CellFontBold = True Then
        answer = MsgBox("Are you sure you wish to delete this entire category?", vbYesNo)
    Else
        answer = MsgBox("Are you sure you wish to delete this favorite?", vbYesNo)
    End If
    Set rsFavs2 = New ADODB.Recordset
    Dim strSQLRetrieve As String
    If answer = vbYes Then
        If MSFlexGrid1.CellFontBold = True Then
            strSQLRetrieve = "DELETE FROM Categories Where Category Like '" & MSFlexGrid1.Text & "' "
            rsFavs2.Open strSQLRetrieve, _
                cnFavorites, adOpenStatic, _
                adLockOptimistic
            Set rsFavs3 = New ADODB.Recordset
            strSQLRetrieve = "DELETE FROM Favorites Where Category Like '" & MSFlexGrid1.Text & "' "
            rsFavs3.Open strSQLRetrieve, _
                cnFavorites, adOpenStatic, _
                adLockOptimistic
        Else
            MSFlexGrid1.Col = 1
            strSQLRetrieve = "DELETE FROM Favorites Where URLAddress Like '" & MSFlexGrid1.Text & "' "
            rsFavs2.Open strSQLRetrieve, _
                cnFavorites, adOpenStatic, _
                adLockOptimistic
            
            MSFlexGrid1.Col = 0
        End If
        Call Populate_Grid
    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set cnFavorites = New ADODB.Connection
    cnFavorites.Provider = "Microsoft.Jet.OLEDB.4.0;"
    cnFavorites.ConnectionString = "Persist Security Info = False;" _
       & "Data Source = Components/db1.mdb"
    cnFavorites.Open
    If Expanded = False Then
        UnderL = True
    Else
        UnderL = False
    End If
    Call Populate_Grid
End Sub

Private Sub Populate_Grid()
    With MSFlexGrid1
        Dim first As Boolean
        first = True
        .Redraw = False
        .Cols = 2
        .ColWidth(0) = MSFlexGrid1.Width
        .ColWidth(1) = 0
        .Row = 0
        .Col = 0
        RH = .RowHeight(0)
        .Rows = 1
        Set rsCats = New ADODB.Recordset
        Dim strSQLRetrieve As String
        strSQLRetrieve = "SELECT * from Categories Where UserName Like '" & UserNames & "' ORDER BY Category"
        rsCats.Open strSQLRetrieve, _
            cnFavorites, adOpenStatic, _
            adLockOptimistic
        Do Until rsCats.EOF
            Set rsFavs = New ADODB.Recordset
            strSQLRetrieve = "SELECT * from Favorites WHERE Category Like '" & rsCats!Category & "' AND UserName Like '" & UserNames & "' "
            rsFavs.Open strSQLRetrieve, _
                cnFavorites, adOpenStatic, _
                adLockOptimistic
                If first = False Then
                    .Rows = .Rows + 1
                    .Row = .Row + 1
                Else
                    first = False
                End If
                .Col = 0
                .CellFontBold = True
                .Text = rsCats!Category
                Do Until rsFavs.EOF
                    .Rows = .Rows + 1
                    .Row = .Row + 1
                    .Col = 0
                    .Text = "          " & rsFavs!Title
                    .Col = 1
                    .Text = rsFavs!URLAddress
                    rsFavs.MoveNext
                Loop
            rsCats.MoveNext
        Loop
        .Redraw = True
    .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 0
    .CellFontBold = True
    .RowHeight(.Row) = 0
    .Row = 0
    If UnderL = True Then
        UnderL = False
        Do Until .Row = .Rows - 2
            If .CellFontBold = True Then
                Call MSFlexGrid1_DblClick
            End If
            .Row = .Row + 1
        Loop
    End If
    End With
End Sub

Private Sub MSFlexGrid1_DblClick()
    On Error Resume Next
    Dim tempIndex As Integer
    With MSFlexGrid1
    tempIndex = .Row
        If .CellFontBold = False Then
            Call Command1_Click
        Else
            If .CellFontUnderline = True Then
                .CellFontUnderline = False
                .Row = .Row + 1
                Do Until .CellFontBold = True Or (.Row >= .Rows)
                    .RowHeight(.Row) = RH
                    If .Row <> .Rows - 1 Then
                        .Row = .Row + 1
                    Else
                        .CellFontBold = True
                    End If
                Loop
            Else
                .CellFontUnderline = True
                .Row = .Row + 1
                Do Until .CellFontBold = True Or (.Row >= .Rows)
                    .RowHeight(.Row) = 0
                    If .Row <> .Rows - 1 Then
                        .Row = .Row + 1
                    Else
                        .CellFontBold = True
                    End If
                Loop
            End If
        End If
    .Row = tempIndex
    End With
End Sub
