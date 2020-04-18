VERSION 5.00
Begin VB.Form frmJump 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jump To Address"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdJumpHex 
      Caption         =   "&Jump"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdJumpDecimal 
      Caption         =   "&Jump"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtJumpHex 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Text            =   "0"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtJumpDecimal 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "0"
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Jump to Hex Address"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Jump to Decmial Address"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmJump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdJumpDecimal_Click()
If ProgramLoaded = False Then Exit Sub
 
 If txtJumpDecimal.Text < mBaseAddress Then
    MsgBox "Your Decimal number is less than the base address!", vbExclamation
    Exit Sub
 End If
  If txtJumpDecimal.Text > mMaxAddress Then
    MsgBox "Your Decimal number is greater than the max address!", vbExclamation
    Exit Sub
 End If
    mCurrentAddress = txtJumpDecimal.Text
    
    frmMain.PrintOffsetList mBaseAddress, mCurrentAddress, mMaxAddress
    
    Call frmMain.PrintMemoryHex
     frmMain.vScroll1.Value = (mCurrentAddress - mBaseAddress) / 16
End Sub

Private Sub cmdJumpHex_Click()
    If ProgramLoaded = False Then Exit Sub
    Dim Number As Long
    Number = modGlobals.Hex2Dec(txtJumpHex.Text)
 If Number < mBaseAddress Then
    MsgBox "Your Hexadecimal number is less than the base address!", vbExclamation
    Exit Sub
 End If
  If Number > mMaxAddress Then
    MsgBox "Your Hexadecimal number is greater than the max address!", vbExclamation
    Exit Sub
 End If
    mCurrentAddress = Number
    frmMain.PrintOffsetList mBaseAddress, mCurrentAddress, mMaxAddress
    Call frmMain.PrintMemoryHex
    frmMain.vScroll1.Value = (mCurrentAddress - mBaseAddress) / 16
End Sub

Private Sub Form_Load()
    txtJumpDecimal.Text = mBaseAddress
    txtJumpHex.Text = BigDecToHex(mBaseAddress)
    
End Sub

Private Sub txtJumpDecimal_Change()
    If IsNumeric(txtJumpDecimal.Text) = False Then txtJumpDecimal.Text = mBaseAddress
End Sub

Private Sub txtJumpHex_Change()
    'Vaild Hex?
    
End Sub
