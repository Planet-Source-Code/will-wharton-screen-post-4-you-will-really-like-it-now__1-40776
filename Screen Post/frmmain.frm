VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Screen Post"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2610
   ControlBox      =   0   'False
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   2610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmds 
      Caption         =   "Add to System Tray"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   2415
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmmain.frx":0442
   End
   Begin VB.CommandButton cmdopen 
      Caption         =   "&Open Post"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmdabout 
      Caption         =   "About"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "&New Post"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1440
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "RTF"
      Filter          =   "Screen Post (*.spt)|*.spt|All Files (*.*)|*.*"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdabout_Click()
    frmAbout.Show
End Sub

Private Sub cmdnew_Click()
  Dim newpost As frmpost
  Set newpost = New frmpost
  newpost.Visible = True
End Sub

Private Sub cmdopen_Click()
    rtb.Text = ""
    cd.CancelError = True
    On Error GoTo Errhandler:
    cd.Flags = cdlOFNFileMustExist
    cd.ShowOpen
    rtb.LoadFile cd.FileName, rtfRTF
    
    Dim newpost As frmpost
    Set newpost = New frmpost
    newpost.Visible = True
    newpost.rtb = frmmain.rtb
Errhandler:
End Sub

Private Sub cmds_Click()
    Me.WindowState = 1
End Sub

Private Sub Form_Load()
    'Dim Reg As Object
    'Set Reg = CreateObject("wscript.shell")
    'Reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
    cd.InitDir = App.Path
    Me.Show
    Me.Refresh
    With nid
    .cbSize = Len(nid)
    .hwnd = Me.hwnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon
    .szTip = " Screen Post " & vbNullChar
    End With
    cmdabout_Click
    End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Result As Long
    Dim msg As Long
    If Me.ScaleMode = vbPixels Then
    msg = x
    Else
    msg = x / Screen.TwipsPerPixelX
    End If
    Select Case msg
    Case WM_LBUTTONUP
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
    Shell_NotifyIcon NIM_DELETE, nid
    Case WM_LBUTTONDBLCLK
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
    Shell_NotifyIcon NIM_DELETE, nid
    Case WM_RBUTTONUP
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
    Shell_NotifyIcon NIM_DELETE, nid
    End Select
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
    Me.Hide
    Shell_NotifyIcon NIM_ADD, nid
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, nid
    End
End Sub
