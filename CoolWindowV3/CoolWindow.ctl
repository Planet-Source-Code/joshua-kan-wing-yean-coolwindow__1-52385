VERSION 5.00
Begin VB.UserControl CoolWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7125
   ControlContainer=   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   7125
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2160
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1320
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   1080
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "CoolWindow"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   50
      Width           =   2415
   End
   Begin VB.Image ImageResize 
      Height          =   75
      Left            =   7040
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   75
   End
   Begin VB.Image CloseOver 
      Height          =   270
      Left            =   6720
      Picture         =   "CoolWindow.ctx":0000
      ToolTipText     =   "Close This Window"
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image ButtonMaxOver 
      Height          =   270
      Left            =   6420
      Picture         =   "CoolWindow.ctx":0361
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image ButtonMinOver 
      Height          =   270
      Left            =   6120
      Picture         =   "CoolWindow.ctx":06BA
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image CloseForm 
      Height          =   270
      Left            =   6690
      Picture         =   "CoolWindow.ctx":0A0E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   300
   End
   Begin VB.Image ButtonMax 
      Height          =   270
      Left            =   6420
      Picture         =   "CoolWindow.ctx":0D58
      Top             =   120
      Width           =   300
   End
   Begin VB.Image ButtonMin 
      Height          =   270
      Left            =   6120
      Picture         =   "CoolWindow.ctx":108F
      Top             =   120
      Width           =   300
   End
   Begin VB.Image ButtonTray 
      Height          =   380
      Left            =   5760
      Picture         =   "CoolWindow.ctx":13BC
      Stretch         =   -1  'True
      Top             =   80
      Width           =   1200
   End
   Begin VB.Image titleMiddle 
      Height          =   555
      Left            =   75
      Picture         =   "CoolWindow.ctx":1750
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6950
   End
   Begin VB.Image titleLeft 
      Height          =   4770
      Left            =   0
      Picture         =   "CoolWindow.ctx":305B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   75
   End
   Begin VB.Image titleRight 
      Height          =   4770
      Left            =   7040
      Picture         =   "CoolWindow.ctx":359B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   75
   End
   Begin VB.Image titleBottom 
      Height          =   75
      Left            =   75
      Picture         =   "CoolWindow.ctx":3ADB
      Stretch         =   -1  'True
      Top             =   4695
      Width           =   6960
   End
End
Attribute VB_Name = "CoolWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const WM_NCLBUTTONDOWN = &HA1 'Send A Message to tell computer left mouse button is down now
Const HTCAPTION = 2
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private DBClick As Boolean
Private MaxCheck As Boolean

Private Type PointAPI
        X As Long
        Y As Long
End Type

Dim Checker As Integer
Dim MaxCheckInt As Integer
Dim MoveMe As String
Dim aX As Integer, aY As Integer
Dim WinWnd As Object
Dim MinShow As Boolean
Dim MaxShow As Boolean
Dim CloseShow As Boolean
'Dim UnInFirst As Integer
Private WithEvents MYFORM As Form
Attribute MYFORM.VB_VarHelpID = -1
Dim WMYFORM As Integer
Dim HMYFORM As Integer
Dim Check_Resizing As Boolean



'Private WithEvents MYFORM As Form
Private Sub userForm_Resize()
Dim MYFORM As Object
Set MYFORM = UserControl.Parent
UserControl.Width = MYFORM.Width
UserControl.Height = MYFORM.Height
End Sub
Private Sub ButtonMax_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonMax.Visible = False
    ButtonMaxOver.Visible = True
    If CloseShow = True Then
        CloseForm.Visible = True
    End If
    CloseOver.Visible = False
    If MinShow = True Then
        ButtonMin.Visible = True
    End If
    ButtonMinOver.Visible = False
    MousePointer = vbDefault
    MaxShow = True
End Sub

Private Sub ButtonMaxOver_Click()
Set MYFORM = UserControl.Parent
    MaxCheckInt = 1
    MaxCheckInt = MaxCheckInt + 1
    If MaxCheck = False And MaxCheckInt = 2 Then
      MYFORM.WindowState = 2
      Call userForm_Resize
      ImageResize.Visible = False
      MaxCheckInt = 1
      MaxCheck = True
      DBClick = True
    End If
    If MaxCheck = True And MaxCheckInt = 2 Then
      MYFORM.WindowState = 0
      Call userForm_Resize
      ImageResize.Visible = True
      MaxCheckInt = 1
      MaxCheck = False
      DBClick = False
    End If
End Sub

Private Sub ButtonMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonMin.Visible = False
    ButtonMinOver.Visible = True
    If MaxShow = True Then
        ButtonMax.Visible = True
    End If
    ButtonMaxOver.Visible = False
    If CloseShow = True Then
        CloseForm.Visible = True
    End If
    CloseOver.Visible = False
    MousePointer = vbDefault
    MinShow = True
End Sub

Private Sub ButtonMinOver_Click()
Set MYFORM = UserControl.Parent
    MYFORM.WindowState = 1
    Check_Resizing = False
    Timer3.Enabled = True
End Sub
Private Sub ButtonTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If CloseShow = True Then
        CloseForm.Visible = True
    End If
    If MaxShow = True Then
        ButtonMax.Visible = True
    End If
    CloseOver.Visible = False
    ButtonMaxOver.Visible = False
    If MinShow = True Then
        ButtonMin.Visible = True
    End If
    ButtonMinOver.Visible = False
    MousePointer = vbDefault
End Sub

Private Sub CloseForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CloseForm.Visible = False
    CloseOver.Visible = True
    If MaxShow = True Then
        ButtonMax.Visible = True
    End If
    ButtonMaxOver.Visible = False
    If MinShow = True Then
        ButtonMin.Visible = True
    End If
    ButtonMinOver.Visible = False
    MousePointer = vbDefault
    CloseShow = True
End Sub
Private Sub CloseOver_Click()
Set MYFORM = UserControl.Parent
    Unload MYFORM
End Sub

Private Sub lbltitle_DblClick()
    Call titleMiddle_DblClick
End Sub

Private Sub lbltitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
  Set MYFORM = UserControl.Parent
  If CloseShow = True Then
    CloseForm.Visible = True
  End If
  If MaxShow = True Then
    ButtonMax.Visible = True
  End If
  CloseOver.Visible = False
  ButtonMaxOver.Visible = False
  If MinShow = True Then
    ButtonMin.Visible = True
  End If
  ButtonMinOver.Visible = False
  MousePointer = vbDefault
  
  If Button = 1 Then
     Call ReleaseCapture
     lngReturnValue = SendMessage(MYFORM.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
  
End Sub



Private Sub Timer2_Timer()
Dim MYFORM As Object
Set MYFORM = UserControl.Parent

If MYFORM.Width <> WMYFORM Or MYFORM.Height <> HMYFORM Then
   Call userForm_Resize
End If
End Sub

Private Sub Timer3_Timer()
Dim MYFORM As Object
Set MYFORM = UserControl.Parent
If MYFORM.WindowState <> 1 Then
    Check_Resizing = True
    Call userForm_Resize
End If
End Sub

'Private Sub UserControl_Click()
'Set MYFORM = UserControl.Parent
'    MYFORM.Width = WMYFORM + 200
'    MYFORM.Height = HMYFORM + 200
'End Sub

Private Sub UserControl_Initialize()
    DBClick = False
    MaxCheck = False
    Check_Resizing = True
    'MaxShow = True
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If CloseShow = True Then
        CloseForm.Visible = True
    End If
    If MaxShow = True Then
        ButtonMax.Visible = True
    End If
    CloseOver.Visible = False
    ButtonMaxOver.Visible = False
    If MinShow = True Then
        ButtonMin.Visible = True
    End If
    ButtonMinOver.Visible = False
    MousePointer = vbDefault
    'Call titleMiddle_MouseMove
    
    'Rpos.Y = UserControl.Height
End Sub

Private Sub UserControl_Resize()
  '   If Me.Width <= 3749 Then Me.Width = 3750
'    If Me.Height <= 1349 Then Me.Height = 1350
'    On Error Resume Next
    Dim MYFORM As Object
    Set MYFORM = UserControl.Parent
    lbltitle.Width = UserControl.Width / 3
    lbltitle.Left = UserControl.Width / 2.9
    titleMiddle.Width = UserControl.Width - 150
    titleRight.Left = UserControl.Width - 80
    titleRight.Height = UserControl.Height
    titleLeft.Height = UserControl.Height
    titleBottom.Width = UserControl.Width - 50
    titleBottom.Top = UserControl.Height - 80
    ButtonTray.Left = UserControl.Width - 1370
    CloseOver.Left = UserControl.Width - 500
    CloseForm.Left = UserControl.Width - 500
    ButtonMax.Left = UserControl.Width - 780
    ButtonMaxOver.Left = UserControl.Width - 780
    ButtonMin.Left = UserControl.Width - 1090
    ButtonMinOver.Left = UserControl.Width - 1090
    ImageResize.Top = UserControl.Height - 80
    ImageResize.Left = UserControl.Width - 80
    If MaxCheck <> True Then
        Call UserControl_Show
    End If
    'If UnInFirst <> 0 Then

End Sub
Private Sub ImageResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set MYFORM = UserControl.Parent
If MYFORM.WindowState <> 2 Then
    MousePointer = 99
    MouseIcon = LoadPicture(App.Path & "\size2_rm.cur")
End If
End Sub

Private Sub titleBottom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result As Long
'Dim Existingheight As Integer
Set MYFORM = UserControl.Parent
    Dim Pos As PointAPI
    If MYFORM.WindowState <> 2 Then
        MoveMe = "Yes"
        Do
'        Existingheight = UserControl.Height
        Result = GetCursorPos(Pos)
        aY% = Pos.Y
            DoEvents
        Result = GetCursorPos(Pos)
        
        UserControl.Height = UserControl.Height + (Pos.Y - aY%) * 32
        MYFORM.Height = UserControl.Height
        If UserControl.Height <= 1349 Then UserControl.Height = 1350
        On Error Resume Next
        Loop Until MoveMe = "No"
    Exit Sub
    End If
End Sub

Private Sub titleBottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set MYFORM = UserControl.Parent
If MYFORM.WindowState <> 2 Then
    MousePointer = 99
    MouseIcon = LoadPicture(App.Path & "\size4_rm.cur")
End If
End Sub

Private Sub titleBottom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveMe = "No"
End Sub

Private Sub titleMiddle_DblClick()
If MaxShow = True Then
Set MYFORM = UserControl.Parent

    Checker = 1
    Checker = Checker + 1
    If DBClick = False And Checker = 2 Then
       MYFORM.WindowState = 2
       Call userForm_Resize
       ImageResize.Visible = False
       DBClick = True
       MaxCheck = True
       Checker = 1
    End If
    
    If DBClick = True And Checker = 2 Then
       MYFORM.WindowState = 0
       Call userForm_Resize
       ImageResize.Visible = True
       DBClick = False
       MaxCheck = False
       Checker = 1
    End If
End If
End Sub

Private Sub titleMiddle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lngReturnValue As Long
  Set MYFORM = UserControl.Parent
  If CloseShow = True Then
    CloseForm.Visible = True
  End If
  If MaxShow = True Then
    ButtonMax.Visible = True
  End If
  CloseOver.Visible = False
  ButtonMaxOver.Visible = False
  If MinShow = True Then
    ButtonMin.Visible = True
  End If
  ButtonMinOver.Visible = False
  MousePointer = vbDefault
  
  If Button = 1 Then
     Call ReleaseCapture
     lngReturnValue = SendMessage(MYFORM.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub
Private Sub ImageResize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Result As Long
    Dim Pos As PointAPI
        MoveMe = "Yes"
        Do
        Result = GetCursorPos(Pos)
        aX% = Pos.X
        aY% = Pos.Y
            DoEvents
        Result = GetCursorPos(Pos)
        
        UserControl.Width = UserControl.Width + (Pos.X - aX%) * 35
        UserControl.Height = UserControl.Height + (Pos.Y - aY%) * 35
        MYFORM.Width = UserControl.Width
        MYFORM.Height = UserControl.Height
        If UserControl.Width <= 4200 Then UserControl.Width = 4250
        If UserControl.Height <= 1349 Then UserControl.Height = 1350
        On Error Resume Next
        Loop Until MoveMe = "No"
    Exit Sub
End Sub
Private Sub ImageResize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveMe = "No"
End Sub

Private Sub titleRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Result As Long
    Set MYFORM = UserControl.Parent
    Dim Pos As PointAPI
    If MYFORM.WindowState <> 2 Then
        MoveMe = "Yes"
        Do
        Result = GetCursorPos(Pos)
        aX% = Pos.X
            DoEvents
        Result = GetCursorPos(Pos)
        
        UserControl.Width = UserControl.Width + (Pos.X - aX%) * 32
        MYFORM.Width = UserControl.Width
        If UserControl.Width <= 4200 Then UserControl.Width = 4250
        On Error Resume Next
        Loop Until MoveMe = "No"
    Exit Sub
    End If
End Sub

Private Sub titleRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set MYFORM = UserControl.Parent
If MYFORM.WindowState <> 2 Then
    MousePointer = 99
    MouseIcon = LoadPicture(App.Path & "\size3_m.cur")
End If
End Sub

Private Sub titleRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveMe = "No"
End Sub

Private Sub UserControl_Show()

''If UnInFirst <> 0 Then
    Dim ocx As Object
'    Dim HTOP As Integer
'    Dim HLEFT As Integer
    On Error GoTo Here:
        For Each ocx In UserControl.Parent
            If TypeOf ocx Is CoolWindow Then
                ocx.Left = 0
                ocx.Top = 0
                If UserControl.Width <= 4200 Then UserControl.Width = 4250
                If UserControl.Height < 1349 Then UserControl.Height = 1350
                Exit For
            End If
        Next
Here:
Timer1.Enabled = True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim MYFORM As Object
Set MYFORM = UserControl.Parent
    Call PropBag.WriteProperty("Caption", lbltitle.Caption, "CoolWindow")
    Call PropBag.WriteProperty("TitleFontColor", lbltitle.ForeColor, &H80000012)
'    Call PropBag.WriteProperty("TitleFontName", lbltitle.FontName, "Times New Roman")
'    Call PropBag.WriteProperty("TitleFontSize", lbltitle.FontSize, 12)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("ShowMinimizeButton", ButtonMin.Visible, True)
    Call PropBag.WriteProperty("ShowMaximizeButton", ButtonMax.Visible, True)
    Call PropBag.WriteProperty("ShowCloseButton", CloseForm.Visible, True)
    Call PropBag.WriteProperty("ShowButtonTray", ButtonTray.Visible, True)
    Call PropBag.WriteProperty("ShowInTaskbar", MYFORM.ShowInTaskbar, True)
   
    Call PropBag.WriteProperty("TitleFont", lbltitle.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000007)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim MYFORM As Object
Set MYFORM = UserControl.Parent
    lbltitle.Caption = PropBag.ReadProperty("Caption", "CoolWindow")
    lbltitle.ForeColor = PropBag.ReadProperty("TitleFontColor", &H80000012)
'    lbltitle.FontName = PropBag.ReadProperty("TitleFontName", "Times New Roman")
'    lbltitle.FontSize = PropBag.ReadProperty("TitleFontSize", 12)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    ButtonMin.Visible = PropBag.ReadProperty("ShowMinimizeButton", True)
    ButtonMax.Visible = PropBag.ReadProperty("ShowMaximizeButton", True)
    CloseForm.Visible = PropBag.ReadProperty("ShowCloseButton", True)
    ButtonTray.Visible = PropBag.ReadProperty("ShowButtonTray", True)
    MYFORM.ShowInTaskbar = PropBag.ReadProperty("ShowInTaskbar", True)
   
    Set lbltitle.Font = PropBag.ReadProperty("TitleFont", Ambient.Font)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000007)
End Sub
Public Property Get Caption() As String
    Caption = lbltitle.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lbltitle.Caption() = New_Caption
    UserControl.Parent.Caption = New_Caption
    PropertyChanged "Caption"
End Property

Public Property Get TitleFontColor() As OLE_COLOR
    TitleFontColor = lbltitle.ForeColor
End Property
Public Property Let TitleFontColor(ByVal New_TitleFontColor As OLE_COLOR)
    lbltitle.ForeColor() = New_TitleFontColor
    PropertyChanged "TitleFontColor"
End Property

'Public Property Get TitleFontName() As String
'    TitleFontName = lbltitle.FontName
'End Property
'
'Public Property Let TitleFontName(ByVal New_TitleFontName As String)
'    lbltitle.FontName() = New_TitleFontName
'    PropertyChanged "TitleFontName"
'End Property
'
'Public Property Get TitleFontSize() As Single
'    TitleFontSize = lbltitle.FontSize
'End Property
'
'Public Property Let TitleFontSize(ByVal New_TitleFontSize As Single)
'    lbltitle.FontSize() = New_TitleFontSize
'    PropertyChanged "TitleFontSize"
'End Property

Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

Public Property Get ShowMinimizeButton() As Boolean
    ShowMinimizeButton = ButtonMin.Visible
End Property

Public Property Let ShowMinimizeButton(ByVal New_ShowMinimizeButton As Boolean)
    ButtonMin.Visible() = New_ShowMinimizeButton
    If New_ShowMinimizeButton = False Then
        ButtonMin.Visible = False
        MinShow = False
    Else
        ButtonMin.Visible = True
        MinShow = True
    End If
    PropertyChanged "ShowMinimizeButton"
End Property

Public Property Get ShowMaximizeButton() As Boolean
    ShowMaximizeButton = ButtonMax.Visible
End Property

Public Property Let ShowMaximizeButton(ByVal New_ShowMaximizeButton As Boolean)
    ButtonMax.Visible() = New_ShowMaximizeButton
    If New_ShowMaximizeButton = False Then
        ButtonMax.Visible = False
        MaxShow = False
    Else
        ButtonMax.Visible = True
        MaxShow = True
    End If
    PropertyChanged "ShowMaximizeButton"
End Property

Public Property Get ShowCloseButton() As Boolean
    ShowCloseButton = CloseForm.Visible
End Property

Public Property Let ShowCloseButton(ByVal New_ShowCloseButton As Boolean)
    CloseForm.Visible() = New_ShowCloseButton
    If New_ShowCloseButton = False Then
        CloseForm.Visible = False
        CloseShow = False
    Else
        CloseForm.Visible = True
        CloseShow = True
    End If
    PropertyChanged "ShowCloseButton"
End Property

Public Property Get ShowButtonTray() As Boolean
    ShowButtonTray = ButtonTray.Visible
End Property

Public Property Let ShowButtonTray(ByVal New_ShowButtonTray As Boolean)
    ButtonTray.Visible() = New_ShowButtonTray
    If New_ShowButtonTray = False Then
        ButtonTray.Visible = False
    Else
        ButtonTray.Visible = True
    End If
    PropertyChanged "ShowButtonTray"
End Property

Public Property Get ShowInTaskbar() As Boolean
Dim MYFORM As Object
Set MYFORM = UserControl.Parent
    ShowInTaskbar = MYFORM.ShowInTaskbar
End Property

Public Property Let ShowInTaskbar(ByVal New_ShowInTaskbar As Boolean)
Dim MYFORM As Object
Set MYFORM = UserControl.Parent
    MYFORM.ShowInTaskbar() = New_ShowInTaskbar
    PropertyChanged "ShowInTaskbar"
End Property

Private Sub Timer1_Timer()
Dim MYFORM As Object
Set MYFORM = UserControl.Parent
    'MYFORM.Caption = ""
    MYFORM.BorderStyle = 0
    MYFORM.Appearance = 0
'    WinWnd.Top = HTOP
'    WinWnd.Left = HLEFT
If MYFORM.WindowState <> 2 And Check_Resizing <> False Then
    MYFORM.Width = UserControl.Width
    MYFORM.Height = UserControl.Height
    WMYFORM = MYFORM.Width
    HMYFORM = MYFORM.Height
End If
If ButtonMax.Visible = True Then
    MaxShow = True
End If
Timer1.Enabled = False
Timer2.Enabled = True
End Sub

Public Property Get TitleFont() As Font
Attribute TitleFont.VB_Description = "Returns a Font object."
    Set TitleFont = lbltitle.Font
End Property

Public Property Set TitleFont(ByVal New_TitleFont As Font)
    Set lbltitle.Font = New_TitleFont
    PropertyChanged "TitleFont"
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

