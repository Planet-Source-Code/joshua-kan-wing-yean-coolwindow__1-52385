VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Main Form"
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.CoolWindow CoolWindow 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6376
      Caption         =   "Main Form"
      AutoRedraw      =   -1  'True
      ShowInTaskbar   =   0   'False
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483640
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit Example"
         Height          =   495
         Left            =   5280
         TabIndex        =   3
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmdSub 
         Caption         =   "Call Sub Form"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtWelcome 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1815
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   720
         Width           =   6495
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    MsgBox "Please Vote Me, Thanks", vbOKOnly, "Welcome"
    Unload Me
    Unload Form2
End Sub

Private Sub cmdSub_Click()
    Form2.Show
End Sub

Private Sub Form_Load()
    txtWelcome.Text = "Welcome and thanks your guys/gals are going to download my code. If yours have any problems with my control, please write down yours comments as well. Actually this is my first control, hope to get yours vote."
End Sub

Private Sub Form_Resize()
On Error GoTo Here:
If Form1.Height >= 500 Then
    txtWelcome.Height = Form1.Height - 2500
End If
If Form1.Width >= 400 Then
    txtWelcome.Width = Form1.Width - 480
End If
Here:
End Sub
