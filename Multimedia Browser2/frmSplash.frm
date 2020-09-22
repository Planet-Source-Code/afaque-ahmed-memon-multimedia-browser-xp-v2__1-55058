VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   ClientHeight    =   2685
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   3465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   3615
      TabIndex        =   4
      Top             =   1560
      Width           =   3615
      Begin VB.Image Image5 
         Height          =   270
         Left            =   2280
         Picture         =   "frmSplash.frx":0000
         ToolTipText     =   "Show HTML Viewer"
         Top             =   0
         Width           =   270
      End
      Begin VB.Image Image2 
         Height          =   285
         Left            =   840
         Picture         =   "frmSplash.frx":0432
         ToolTipText     =   "Show Text Viewer"
         Top             =   0
         Width           =   285
      End
      Begin VB.Image Image3 
         Height          =   270
         Left            =   1320
         Picture         =   "frmSplash.frx":08E8
         ToolTipText     =   "Show Image Viewer"
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   270
         Left            =   1800
         Picture         =   "frmSplash.frx":0CD2
         ToolTipText     =   "Show Media Player"
         Top             =   0
         Width           =   270
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "V2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   2640
      Picture         =   "frmSplash.frx":1104
      Stretch         =   -1  'True
      Top             =   480
      Width           =   465
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.afaque.tk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "eXPerience Browsing!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1020
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MultiMedia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   2430
   End
   Begin VB.Label lblCopyright 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2004 Afaque Ahmed Memon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Browser"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1830
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "XP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   2520
      TabIndex        =   0
      Top             =   840
      Width           =   645
   End
   Begin VB.Image Image1 
      Height          =   2715
      Left            =   0
      Picture         =   "frmSplash.frx":7956
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3540
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub
Private Sub Image1_Click()
Unload Me
End Sub
