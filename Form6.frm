VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10515
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   7605
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   6240
      Top             =   1560
   End
   Begin VB.OLE OLE1 
      Class           =   "SoundRec"
      Height          =   735
      Left            =   8400
      OleObjectBlob   =   "Form6.frx":30186
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    OLE1.DoVerb
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Form4.Show (vbModal)
    Unload Me
End Sub
