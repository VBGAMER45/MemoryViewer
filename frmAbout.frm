VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Memory Viewer & Dumper"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "About vbgamer45"
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   840
      Width           =   1467
   End
   Begin VB.Label lblWebSite 
      Caption         =   "http://www.visualbasiczone.com/products/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label lblTitle 
      Caption         =   "Memory Viewer And Dumper"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Memory Viewer & Dumper Version: " & Version
End Sub

Private Sub lblWebSite_Click()
 ShellExecute Me.hwnd, vbNullString, "http://www.visualbasiczone.com/products/", vbNullString, "C:\", SW_SHOWNORMAL

End Sub
