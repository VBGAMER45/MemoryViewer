VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDumpMem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save Memory Dump To File"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4515
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3960
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton mnuCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdDump 
      Caption         =   "&Dump"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dump Address"
      Height          =   1095
      Left            =   2760
      TabIndex        =   3
      Top             =   0
      Width           =   1695
      Begin VB.OptionButton optHex 
         Caption         =   "Hex"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optDec 
         Caption         =   "Decimal"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox txtEnd 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtBegin 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "End Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Begin Address: "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmDumpMem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDump_Click()
On Error GoTo errHandle
    Dim LowerLimit As Double
    Dim HightLimit As Double
    Dim Buffer As String * 1000
    Dim readlen As Long
    Dim addr As Long
    If txtEnd.Text = "" Then Exit Sub
    If txtBegin.Text = "" Then Exit Sub
    If optHex.Value = True Then
        LowerLimit = modGlobals.Hex2Dec(txtBegin.Text)
        HighLimit = modGlobals.Hex2Dec(txtEnd.Text)
    Else
        LowerLimit = txtBegin.Text
        HighLimit = txtEnd.Text
    
    End If

    CD1.DefaultExt = ".txt"
    CD1.DialogTitle = "Save Memory Title"
    CD1.FileName = ""
    CD1.ShowSave
    If CD1.FileName = "" Then Exit Sub
    
    Dim f As Long
    f = FreeFile
    Open CD1.FileName For Binary Access Write Lock Write As #f
        For addr = LowerLimit To HighLimit Step 1000 '11536388
            Call ReadProcessMemory(myHandle, addr, Buffer, 1000, readlen)
            Put #f, , Buffer
         Next
     Close #f

    MsgBox "Memory Dumped to " & App.Path & "\dump.txt", vbInformation
Exit Sub
errHandle:
    MsgBox "Error_frmDumpMem_cmdDump_Click(): " & err.Number & " " & err.Description
End Sub

Private Sub Form_Load()

    txtBegin.Text = mBaseAddress
    txtEnd.Text = mBaseAddress + 1000 ' mMaxAddress
End Sub

Private Sub mnuCancel_Click()
    Unload Me
End Sub

Private Sub txtBegin_Change()
    If optDec.Value = True Then
        If IsNumeric(txtBegin.Text) = False Then txtBegin.Text = mBaseAddress
    End If
End Sub

Private Sub txtEnd_Change()
    If optDec.Value = True Then
        If IsNumeric(txtEnd.Text) = False Then txtEnd.Text = mMaxAddress
    End If
End Sub
