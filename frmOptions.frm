VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChangeGrid 
      Caption         =   "Change"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdChangeBackround 
      Caption         =   "Change"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdChangeHexColor 
      Caption         =   "Change"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdChangeOffset 
      Caption         =   "Change"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Grid Line Color"
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
      Left            =   1080
      TabIndex        =   13
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblColor4 
      Caption         =   "ABCDEFGH"
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblColor3 
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Backround Color"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblColor2 
      Caption         =   "ABCDEFGH"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Hex && Ascii Color"
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
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblColor1 
      Caption         =   "ABCDEFGH"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Offset Color"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChangeBackround_Click()
    CommonDialog1.ShowColor
    lblColor3.BackColor = CommonDialog1.Color
End Sub

Private Sub cmdChangeGrid_Click()
    CommonDialog1.ShowColor
    lblColor4.ForeColor = CommonDialog1.Color
End Sub

Private Sub cmdChangeHexColor_Click()
    CommonDialog1.ShowColor
    lblColor2.ForeColor = CommonDialog1.Color
End Sub

Private Sub cmdChangeOffset_Click()
    CommonDialog1.ShowColor
    lblColor1.ForeColor = CommonDialog1.Color
    
End Sub

Private Sub cmdDone_Click()
    cOffSetColor = lblColor1.ForeColor
    cHexColor = lblColor2.ForeColor
    cGridColor = lblColor4.ForeColor
    cBackroundColor = lblColor3.BackColor
    frmMain.BackColor = cBackroundColor
    'Save Settings For Furture use
    Call SaveSetting("Memory Editor Viewer", "Options", "OffSetColor", cOffSetColor)
    Call SaveSetting("Memory Editor Viewer", "Options", "HexColor", cHexColor)
    Call SaveSetting("Memory Editor Viewer", "Options", "BackroundColor", cBackroundColor)
    Call SaveSetting("Memory Editor Viewer", "Options", "GridColor", cGridColor)
    
    'Call frmMain.PrintOffsetList(mBaseAddress, mCurrentAddress, mMaxAddress)
    Call frmMain.DrawGrid
    Call frmMain.PrintOffsetList(mBaseAddress, mCurrentAddress, mMaxAddress)
    Call frmMain.PrintMemoryHex
    Unload Me
End Sub



Private Sub Form_Load()
'Load Color values
    lblColor1.ForeColor = cOffSetColor
    lblColor2.ForeColor = cHexColor
    lblColor3.BackColor = cBackroundColor
    lblColor4.ForeColor = cGridColor
End Sub
