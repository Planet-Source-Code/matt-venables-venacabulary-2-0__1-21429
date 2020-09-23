VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form results 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   7260
   ClientTop       =   6135
   ClientWidth     =   4680
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Vocabulary Results"
      Height          =   3135
      Left            =   0
      MousePointer    =   15  'Size All
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4335
         ExtentX         =   7646
         ExtentY         =   4683
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
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
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   255
         Left            =   4440
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   120
         Width           =   135
      End
   End
End
Attribute VB_Name = "results"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Left = dict.Left + 50
Me.Top = (dict.Top + dict.Height) - 250


End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Label1_Click()
Unload Me

End Sub
