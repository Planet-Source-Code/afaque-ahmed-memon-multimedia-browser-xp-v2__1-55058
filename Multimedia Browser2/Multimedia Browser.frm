VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form1 
   Caption         =   "MultiMedia Browser XP V2|eXPerience Browsing!"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11865
   FillColor       =   &H80000001&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000001&
   Icon            =   "Multimedia Browser.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   7845
   ScaleWidth      =   11865
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   500
      Left            =   2040
      Max             =   6000
      SmallChange     =   500
      TabIndex        =   13
      Top             =   7560
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7575
      LargeChange     =   500
      Left            =   11600
      Max             =   6000
      SmallChange     =   500
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture22 
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7620
      Left            =   2040
      ScaleHeight     =   7620
      ScaleWidth      =   10125
      TabIndex        =   15
      Top             =   0
      Width           =   10125
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         DrawMode        =   1  'Blackness
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   20000
         Left            =   0
         ScaleHeight     =   19995
         ScaleWidth      =   19995
         TabIndex        =   16
         Top             =   0
         Width           =   20000
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3000
      Top             =   360
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      CausesValidation=   0   'False
      Height          =   7500
      Left            =   2070
      TabIndex        =   8
      Top             =   10
      Width           =   9735
      ExtentX         =   17171
      ExtentY         =   13229
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
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
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7500
      Left            =   2070
      TabIndex        =   1
      Top             =   10
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   13229
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      MousePointer    =   4
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Multimedia Browser.frx":6852
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   0
      ScaleHeight     =   8175
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   235
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   0
         Picture         =   "Multimedia Browser.frx":68C9
         ScaleHeight     =   330
         ScaleWidth      =   255
         TabIndex        =   10
         ToolTipText     =   "Click Here to Expand"
         Top             =   3360
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7845
      Left            =   0
      ScaleHeight     =   7845
      ScaleWidth      =   2055
      TabIndex        =   2
      Top             =   0
      Width           =   2055
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Height          =   1440
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Height          =   1005
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   1575
      End
      Begin VB.FileListBox File2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Height          =   1005
         Left            =   120
         Pattern         =   "*.Bmp;*.Gif;*.Jpg;*.Ico;*.Wmf"
         TabIndex        =   19
         Top             =   3840
         Width           =   1575
      End
      Begin VB.FileListBox File3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Height          =   1005
         Left            =   120
         Pattern         =   "*.mpg;*.mp3;*avi;*.m3u;*.asf;*.wav;*.mid;*.rmi;*.dat;*.wmv"
         TabIndex        =   18
         Top             =   5160
         Width           =   1575
      End
      Begin VB.FileListBox File4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Height          =   1005
         Left            =   120
         Pattern         =   "*.html;*.htm"
         TabIndex        =   17
         Top             =   6480
         Width           =   1575
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   320
         Width           =   1575
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1800
         Top             =   1680
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   8175
         Left            =   1800
         ScaleHeight     =   8175
         ScaleWidth      =   240
         TabIndex        =   11
         Top             =   0
         Width           =   240
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   0
            Picture         =   "Multimedia Browser.frx":6D83
            ScaleHeight     =   330
            ScaleWidth      =   255
            TabIndex        =   12
            ToolTipText     =   "Click Here to Contract"
            Top             =   3360
            Width           =   255
         End
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   0
         ScaleHeight     =   300
         ScaleWidth      =   2055
         TabIndex        =   22
         Top             =   0
         Width           =   2055
         Begin VB.Image Image4 
            Height          =   270
            Left            =   960
            Picture         =   "Multimedia Browser.frx":723D
            ToolTipText     =   "Show Media Player"
            Top             =   0
            Width           =   270
         End
         Begin VB.Image Image3 
            Height          =   270
            Left            =   600
            Picture         =   "Multimedia Browser.frx":766F
            ToolTipText     =   "Show Image Viewer"
            Top             =   0
            Width           =   255
         End
         Begin VB.Image Image2 
            Height          =   285
            Left            =   240
            Picture         =   "Multimedia Browser.frx":7A59
            ToolTipText     =   "Show Text Viewer"
            Top             =   0
            Width           =   285
         End
         Begin VB.Image Image1 
            Height          =   270
            Left            =   1320
            Picture         =   "Multimedia Browser.frx":7F0F
            ToolTipText     =   "Show HTML Viewer"
            Top             =   0
            Width           =   270
         End
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Html Files"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   6240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Multimedia Files"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Graphic Files"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "All Files (Text) "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   1695
      End
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   7500
      Left            =   2070
      TabIndex        =   0
      Top             =   10
      Width           =   9735
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   3
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   -1  'True
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   -1  'True
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   -1  'True
      SendMouseClickEvents=   -1  'True
      SendMouseMoveEvents=   -1  'True
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   -1  'True
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   -1  'True
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   -1  'True
      Volume          =   -200
      WindowlessVideo =   -1  'True
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnutxt 
         Caption         =   "Tex&t Viewer"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuimg 
         Caption         =   "Ima&ge Viewer"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnummp 
         Caption         =   "Mult&imedia Player"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuhtm 
         Caption         =   "Html Viewer"
         Shortcut        =   ^H
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuhlp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabt 
         Caption         =   "A&bout"
         Shortcut        =   ^B
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1.Path = Dir1.Path
File2.Path = Dir1.Path
File3.Path = Dir1.Path
File4.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo a:
Dir1.Path = Drive1
On Error GoTo a:
a:
If Err.Number = 68 Then
MsgBox "Device is not Ready", vbInformation, "Multimedia Browser XP"
End If
End Sub


Private Sub File1_Click()

RichTextBox1.Visible = True
Picture22.Visible = False
MediaPlayer1.Visible = False
WebBrowser1.Visible = False
HScroll1.Visible = False
VScroll1.Visible = False

If Len(Dir1) = 3 Then
fil = Dir1 & File1
RichTextBox1.FileName = fil
On Error GoTo w:
Else
file = Dir1 & "\" & File1
RichTextBox1.FileName = file
On Error GoTo w:
End If
w:
End Sub

Private Sub File2_Click()
On Error Resume Next
HScroll1.Visible = True
VScroll1.Visible = True
RichTextBox1.Visible = False
Picture22.Visible = True
Picture1.Visible = True

MediaPlayer1.Visible = False
WebBrowser1.Visible = False
 
If Len(Dir1) = 3 Then
fil = Dir1 & File2
Picture1.Picture = LoadPicture(fil)
On Error GoTo w:
Else
file = Dir1 & "\" & File2

file = File2.Path & "\" & File2.FileName
If Err.Number = 50003 Then
MsgBox "Device is not Ready", vbInformation, "MultiMedia Browser XP"
End If
Picture1.Picture = LoadPicture(file)
If Err.Number = 50003 Then
MsgBox "Device is not Ready", vbInformation, "MultiMedia Browser XP"
End If

On Error GoTo w:
End If
w:
End Sub
Private Sub File3_Click()
MediaPlayer1.Visible = True
HScroll1.Visible = False
VScroll1.Visible = False
RichTextBox1.Visible = False
Picture1.Visible = False
WebBrowser1.Visible = False
 
If Len(Dir1) = 3 Then
fil = Dir1 & File3
MediaPlayer1.FileName = (fil)
On Error GoTo w:
Else
file = File3.Path & "\" & File3.FileName
MediaPlayer1.FileName = (file)
On Error GoTo w:
End If
w:

End Sub

Private Sub File4_Click()
HScroll1.Visible = False
VScroll1.Visible = False
RichTextBox1.Visible = False
Picture22.Visible = False
MediaPlayer1.Visible = False
WebBrowser1.Visible = True
 
If Len(Dir1) = 3 Then
fil = Dir1 & File4
WebBrowser1.Navigate fil
On Error GoTo w:
Else
file = Dir1 & "\" & File4
WebBrowser1.Navigate file
On Error GoTo w:
End If
w:
End Sub


Private Sub Form_Load()
WebBrowser1.Offline = True
WebBrowser1.Silent = True
Form1.Hide
frmSplash.Show
Timer1.Enabled = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload frmAbout
Unload frmSplash
End Sub

Private Sub mnuabt_Click()
Load frmSplash
frmSplash.Show
End Sub

Private Sub mnuexit_Click()
End
End Sub


Private Sub Image1_Click()
RichTextBox1.Visible = False
Picture22.Visible = False
MediaPlayer1.Visible = False
WebBrowser1.Visible = True
End Sub

Private Sub Image2_Click()
RichTextBox1.Visible = True
Picture22.Visible = False
MediaPlayer1.Visible = False
WebBrowser1.Visible = False
End Sub

Private Sub Image3_Click()
RichTextBox1.Visible = False
Picture22.Visible = True
MediaPlayer1.Visible = False
WebBrowser1.Visible = False
End Sub

Private Sub Image4_Click()
RichTextBox1.Visible = False
Picture22.Visible = False
MediaPlayer1.Visible = True
WebBrowser1.Visible = False
End Sub

Private Sub mnuhtm_Click()
Image1_Click
End Sub

Private Sub mnuimg_Click()
Image3_Click
End Sub

Private Sub mnummp_Click()
Image4_Click
End Sub

Private Sub mnutxt_Click()
Image2_Click
End Sub

Private Sub Picture3_Click()
Picture4_Click
End Sub

Private Sub Picture4_Click()
HScroll1.Left = 2040
HScroll1.Width = 9495

Picture2.Visible = True
Picture3.Visible = False

WebBrowser1.Left = 2070
WebBrowser1.Width = 9735

Picture22.Left = 2070
Picture22.Width = 9735

MediaPlayer1.Left = 2070
MediaPlayer1.Width = 9735

RichTextBox1.Left = 2070
RichTextBox1.Width = 9735
End Sub

Private Sub Picture5_Click()
Picture6_click
End Sub

Private Sub Picture6_click()
HScroll1.Width = 11300
HScroll1.Left = 235

Picture2.Visible = False
Picture3.Visible = True

WebBrowser1.Left = 270
WebBrowser1.Width = 11570

Picture22.Left = 270
Picture22.Width = 11570

MediaPlayer1.Left = 270
MediaPlayer1.Width = 11570

RichTextBox1.Left = 270
RichTextBox1.Width = 11570
End Sub

Private Sub Timer1_Timer()
frmSplash.Hide
Form1.Show
Timer1.Enabled = False
End Sub
Private Sub HScroll1_Change()
Picture1.Left = -HScroll1.Value
End Sub

Private Sub Timer3_Timer()

End Sub

Private Sub VScroll1_Change()
Picture1.Top = -VScroll1.Value
End Sub
Private Sub HScroll1_scroll()
Picture1.Left = -HScroll1.Value
End Sub

Private Sub VScroll1_scroll()
Picture1.Top = -VScroll1.Value
End Sub
