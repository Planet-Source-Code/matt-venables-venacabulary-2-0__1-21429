VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form dict 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   ClientHeight    =   2190
   ClientLeft      =   7095
   ClientTop       =   4050
   ClientWidth     =   4725
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4680
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "vena · cabulary 2.0"
      Height          =   2175
      Left            =   0
      MousePointer    =   15  'Size All
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CheckBox Check2 
         BackColor       =   &H80000009&
         Caption         =   "Message Unfound Words"
         Height          =   165
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H8000000E&
         Caption         =   "Short Definitions"
         Height          =   165
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   1095
         Left            =   240
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   4215
      End
      Begin VB.Line Line1 
         X1              =   2160
         X2              =   2160
         Y1              =   1680
         Y2              =   2040
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         Height          =   255
         Left            =   1320
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Edit"
         Height          =   255
         Left            =   840
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "File"
         Height          =   255
         Left            =   360
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   255
         Left            =   4440
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Labelsearch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "search!"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2160
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   1680
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      Caption         =   "about: <font size = 2>"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   4335
   End
   Begin VB.Menu mnufile 
      Caption         =   "file"
      Visible         =   0   'False
      Begin VB.Menu mnuwordoftheday 
         Caption         =   "word of the day"
      End
      Begin VB.Menu mnuclear 
         Caption         =   "clear"
      End
      Begin VB.Menu break1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "exit"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "edit"
      Visible         =   0   'False
      Begin VB.Menu mnucopy 
         Caption         =   "copy definitions"
      End
      Begin VB.Menu mnupaste 
         Caption         =   "paste words"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "help"
      Visible         =   0   'False
      Begin VB.Menu mnuabout 
         Caption         =   "about"
      End
      Begin VB.Menu break2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuhelpbutton 
         Caption         =   "help"
      End
   End
End
Attribute VB_Name = "dict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Left = (Screen.Width / 2) - (Me.Width / 2)
Me.Top = (Screen.Height / 2) - (results.Height / 2) - (Me.Height / 2)
Me.Width = 4680
Me.Height = 2220

End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me

End Sub

Private Sub Label2_Click()
End

End Sub

Private Sub Label3_Click()
PopupMenu mnufile

End Sub

Private Sub Label4_Click()
PopupMenu mnuedit

End Sub

Private Sub Label5_Click()
PopupMenu mnuhelp

End Sub

Private Sub Labelsearch_Click()
Dim splitreturn, words As String
Dim short As Boolean
Dim message As Boolean


Label1.Caption = "about: <font size = 2>"
Labelsearch.Caption = "searching..."
If Check2.Value = 1 Then
    message = True
Else
    message = False
End If
If Check1.Value = 1 Then
    short = True
Else
    short = False
End If
splitreturn = Splitter(Text1.Text, ",")

a = 1
On Error GoTo a:
b:
    words = splitreturn(a)
    DefineWord words, Inet1, short, message
    a = a + 1
GoTo b:
a:

Labelsearch.Caption = "search!"
results.WebBrowser1.Navigate (Label1.Caption)
results.Show

End Sub

Private Sub mnuabout_Click()
MsgBox "vena·cabulary 2.0 was created by matt venables.", vbInformation, "vena·cabulary"
End Sub

Private Sub mnuclear_Click()
Text1.Text = ""
Label1.Caption = "about: <font size = 2>"
End Sub

Private Sub mnucopy_Click()
Clipboard.SetText (Label1.Caption)
End Sub

Private Sub mnuexit_Click()
Label2_Click

End Sub

Private Sub mnuhelpbutton_Click()
MsgBox "type all the vocabulary words you need to enter seperated by commas.  vena·cabulary 2.0 will search for the words and print the definitions for you to have.  if there is no definition for the word, it means it is not in the dictionary, cencored or not entered correctly.", vbInformation, "vena·cabulary"


End Sub

Private Sub mnupaste_Click()
Text1.Text = Clipboard.GetText
End Sub

Private Sub mnuwordoftheday_Click()
Text1.Text = WordoftheDay(Inet1)
Call Labelsearch_Click

End Sub

