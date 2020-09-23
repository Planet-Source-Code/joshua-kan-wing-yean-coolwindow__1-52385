VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Sub Form"
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   LinkTopic       =   "Form2"
   ScaleHeight     =   3735
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   Begin Project1.CoolWindow CoolWindow1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   6588
      Caption         =   "Sub Form"
      AutoRedraw      =   -1  'True
      ShowMaximizeButton=   0   'False
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
      Begin VB.CommandButton cmdClose 
         Caption         =   "Unload Me"
         Height          =   495
         Left            =   7800
         TabIndex        =   1
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000007&
         Caption         =   $"Form2.frx":0000
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
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2760
         Width           =   9255
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000007&
         Caption         =   "                 ENTER"
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
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   9255
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000007&
         Caption         =   "                 C:\Windows\system32>regsvr32 CoolWindow.OCX"
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
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   9255
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
         Caption         =   "2. And then fire up your console/command prompt, please type in the following:"
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
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   9255
      End
      Begin VB.Label lblSetup2 
         BackColor       =   &H80000007&
         Caption         =   "1. First you need to put the CoolWindow.OCX to the C:\Windows\system32\ (Recommented)"
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
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   9255
      End
      Begin VB.Label lblSetup1 
         BackColor       =   &H80000007&
         Caption         =   "How to use it ?"
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
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   5295
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

