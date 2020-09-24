VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmpost 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3345
   Icon            =   "frmpost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmpost.frx":0442
   ScaleHeight     =   3810
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   3240
      Top             =   2040
   End
   Begin ScreenPost.TransMe TransMe1 
      Left            =   3120
      Top             =   1440
      _extentx        =   847
      _extenty        =   847
   End
   Begin VB.TextBox txttrans 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Text            =   "255"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CheckBox chktop 
      BackColor       =   &H0080FFFF&
      Caption         =   "Always On Top"
      Height          =   255
      Left            =   120
      MaskColor       =   &H0080FFFF&
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtchanged 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "1"
      Top             =   3960
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4683
      _Version        =   393217
      BackColor       =   8454143
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmpost.frx":29B4C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2040
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "RTF"
      Filter          =   "Screen Post (*.spt)|*.spt|All Files (*.*)|*.*"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transparentacy"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   3480
      Width           =   855
   End
   Begin VB.Line spl2 
      BorderWidth     =   3
      X1              =   1680
      X2              =   1680
      Y1              =   3120
      Y2              =   3360
   End
   Begin VB.Label lbltitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Post"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   -50
      Width           =   2655
   End
   Begin VB.Line Spl3 
      BorderWidth     =   3
      X1              =   2400
      X2              =   2400
      Y1              =   3120
      Y2              =   3360
   End
   Begin VB.Image imgprint 
      Height          =   240
      Left            =   1320
      Picture         =   "frmpost.frx":29BCE
      ToolTipText     =   "Print"
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image imgi 
      Height          =   240
      Left            =   2040
      Picture         =   "frmpost.frx":29F10
      ToolTipText     =   "Italic"
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image imgb 
      Height          =   240
      Left            =   1800
      Picture         =   "frmpost.frx":2A252
      ToolTipText     =   "Bold"
      Top             =   3120
      Width           =   240
   End
   Begin VB.Line spl1 
      BorderWidth     =   3
      X1              =   480
      X2              =   480
      Y1              =   3120
      Y2              =   3360
   End
   Begin VB.Image imgclose 
      Height          =   240
      Left            =   2880
      Picture         =   "frmpost.frx":2A594
      ToolTipText     =   "Close"
      Top             =   50
      Width           =   240
   End
   Begin VB.Image imghelp 
      Height          =   240
      Left            =   2520
      Picture         =   "frmpost.frx":2A8D6
      Stretch         =   -1  'True
      ToolTipText     =   "Help"
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image imgnew 
      Height          =   240
      Left            =   120
      Picture         =   "frmpost.frx":2AD18
      Stretch         =   -1  'True
      ToolTipText     =   "New"
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image imginfo 
      Height          =   240
      Left            =   2880
      Picture         =   "frmpost.frx":2B15A
      Stretch         =   -1  'True
      ToolTipText     =   "Info"
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image imgopen 
      Height          =   240
      Left            =   960
      Picture         =   "frmpost.frx":2B89C
      Stretch         =   -1  'True
      ToolTipText     =   "Open"
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image imgsave 
      Height          =   240
      Left            =   600
      Picture         =   "frmpost.frx":2BCDE
      Stretch         =   -1  'True
      ToolTipText     =   "Save"
      Top             =   3120
      Width           =   240
   End
End
Attribute VB_Name = "frmpost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Private Const WM_SYSCOMMAND = &H112




Private Sub chktop_Click()
    If chktop.Value = 1 Then
        AlwaysOnTop.AlwaysOnTop Me, True
    Else
    
    AlwaysOnTop.AlwaysOnTop Me, False
    End If
    
End Sub

Private Sub Form_Load()
    rtb.Text = "<" & Now & ">" & vbCrLf
    cd.InitDir = App.Path
    
    If Me.Picture <> 0 Then
        Call SetAutoRgn(Me)
    End If
End Sub

Private Sub imgb_Click()
    If rtb.SelBold = False Then
        rtb.SelBold = True
    Else
        rtb.SelBold = False
    End If
End Sub

Private Sub imgclose_Click()
    If txtchanged.Text = "1" Then ' 1 = not changed
       Unload Me                 ' 2 = changed
    Else
        If MsgBox("Would you like to save your current file?", vbYesNo, "Save") = vbYes Then
        save
        Unload Me
        Else
        Unload Me
        End If
    End If
End Sub

Private Sub imghelp_Click()
    frmAbout.Show
End Sub

Private Sub imgi_Click()
    If rtb.SelItalic = False Then
        rtb.SelItalic = True
    Else
        rtb.SelItalic = False
    End If
End Sub
Private Sub lbltitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, _
        HTCAPTION, 0&
End Sub

Private Sub imginfo_Click()
    MsgBox ("This was made by DogBone. -= dogbonevb@hotmail.com =- " & vbNewLine & "Please vote")
End Sub

Private Sub imgnew_Click()
    If txtchanged.Text = "1" Then ' 1 = not changed
        rtb.Text = ""             ' 2 = changed
        rtb.Text = "<" & Now & ">" & vbCrLf
        txtchanged.Text = "1"
    Else
       If MsgBox("Would you like to save your current file?", vbYesNo, "Save") = vbYes Then
       save
       rtb.Text = ""
       rtb.Text = "<" & Now & ">" & vbCrLf
        txtchanged.Text = "1"
        End If
    End If
End Sub

Private Sub imgopen_Click()

    If txtchanged.Text = "1" Then ' 1 = not changed
        rtb.Text = ""             ' 2 = changed
        rtb.Text = ""
        txtchanged.Text = "1"
        cd.CancelError = True
        On Error GoTo Errhandler:
        cd.Flags = cdlOFNFileMustExist
        cd.ShowOpen
        rtb.LoadFile cd.FileName, rtfRTF
        txtchanged.Text = "1"
    Else
        If MsgBox("Would you like to save your current file?", vbYesNo, "Save") = vbYes Then
        save
    
        rtb.Text = ""
        txtchanged.Text = "1"
        cd.CancelError = True
        On Error GoTo Errhandler:
        cd.Flags = cdlOFNFileMustExist
        cd.ShowOpen
        rtb.LoadFile cd.FileName, rtfRTF
        'display filename (without path) on status bar
        Else
        rtb.Text = ""
        txtchanged.Text = "1"
        cd.CancelError = True
        On Error GoTo Errhandler:
        cd.Flags = cdlOFNFileMustExist
        cd.ShowOpen
        rtb.LoadFile cd.FileName, rtfRTF
        
        End If
    End If
Errhandler:
    'if Cancel clicked, then exit procedure

End Sub

Private Sub imgprint_Click()
    rtb.SelPrint (Printer.hdc)
End Sub

Private Sub imgsave_Click()
    save
    txtchanged.Text = "1"
End Sub

Private Sub rtb_Change()
    If txtchanged.Text = "1" Then ' 1 = not changed
       txtchanged.Text = "2"     ' 2 = changed
    Else
       txtchanged.Text = "1"
    End If
End Sub

Private Sub save()

    cd.CancelError = True
    On Error GoTo Errhandler:
    cd.ShowSave
    'save specified file in RTF format
    rtb.SaveFile cd.FileName, rtfRTF
    'display filename (without path) on status bar
Errhandler:
    'Cancel button clicked

End Sub
Private Sub form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, _
        HTCAPTION, 0&
End Sub

Private Sub Timer1_Timer()
If txttrans.Text > 255 Then
ElseIf txttrans.Text < 1 Then
Else

TransMe1.MakeTransparent Me.hWnd, txttrans.Text
Timer1.Enabled = False
End If
End Sub

Private Sub txttrans_Change()
Timer1.Enabled = True
End Sub
