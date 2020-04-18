VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSelectProcess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Process"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5745
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.OptionButton OptChoice3 
      Height          =   375
      Left            =   5040
      Picture         =   "frmSelectProcess.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton OptChoice2 
      Height          =   375
      Left            =   4560
      Picture         =   "frmSelectProcess.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton OptChoice1 
      Height          =   375
      Left            =   4080
      Picture         =   "frmSelectProcess.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picSmall 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   960
      Picture         =   "frmSelectProcess.frx":03DE
      ScaleHeight     =   510
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   840
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectProcess.frx":06E8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLarge 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   5880
      ScaleHeight     =   510
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox lstProcess 
      Height          =   2010
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectProcess.frx":0C82
            Key             =   "Test"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView AppRun 
      Height          =   3375
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5953
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblInfo 
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   5535
   End
End
Attribute VB_Name = "frmSelectProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim LowerLimit As Long
Dim HighLimit As Long




Private Sub AppRun_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lblInfo.Caption = Item.ToolTipText
    lblInfo.ToolTipText = "Click this label to copy the text."

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Dim myProcess As PROCESSENTRY32
    Dim mySnapshot As Long

    'first clear our listbox
    lstProcess.Clear

    myProcess.dwSize = Len(myProcess)

    'create snapshot
    mySnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)

    'get first process
    ProcessFirst mySnapshot, myProcess
    lstProcess.AddItem myProcess.szExeFile ' set exe name
    PIDs(lstProcess.ListCount - 1) = myProcess.th32ProcessID ' set PID

    'while there are more processes
    While ProcessNext(mySnapshot, myProcess)
        lstProcess.AddItem myProcess.szExeFile ' set exe name
        Call Load_Filename(modGlobals.ExePathFromProcessId(myProcess.th32ProcessID), myProcess.szExeFile, myProcess.th32ProcessID)
        PIDs(lstProcess.ListCount - 1) = myProcess.th32ProcessID ' ' store PID
    Wend

End Sub
Public Function InitProcessOpener(pid As Long)
    Dim pHandle As Long
    Dim lpMem As Long, ret As Long, lLenMBI As Long
    Dim mbi As MEMORY_BASIC_INFORMATION
  
    
    MyPid = pid 'Save Pid for later use
    pHandle = OpenProcess(16, False, pid)
   
    If (pHandle = 0) Then
        InitProcessOpener = False
        myHandle = 0
    Else
        InitProcessOpener = True
        myHandle = pHandle

        lLenMBI = Len(mbi)
        'Determine applications memory addresses range
        Call GetSystemInfo(si)
        lpMem = si.lpMinimumApplicationAddress
        mbi.RegionSize = 0
        ret = VirtualQueryEx(pHandle, ByVal lpMem, mbi, lLenMBI)
        'Max Address
        mMaxAddress = si.lpMaximumApplicationAddress
        
    'MsgBox BaseModuleHandleFromProcessId(pid)
    'Ms'gBox ExePathFromProcessId(pid)
        mBaseAddress = BaseModuleHandleFromProcessId(pid)
        mCurrentAddress = BaseModuleHandleFromProcessId(pid)
    End If

End Function

Private Sub cmdSelect_Click()
    Dim Buffer As String * 1000
    Dim readlen As Long
    Dim addr As Long
    Dim itemx As ListItem
    Dim ItemName As String
    Dim Number As Long
' `Get the Selected the Item
    For Each itemx In AppRun.ListItems
        If itemx.Selected = True Then
                ItemName = itemx.Text
                'MsgBox ItemName
                'MsgBox itemx.Tag
                Number = itemx.Tag
                Exit For
        End If
        
    Next itemx
    
'    Lowerlimit = txtLow.Text
   ' HighLimit = txtHigh.Text
    
    'If lstProcess.ListIndex = -1 Then MsgBox "Please select a process to open.", vbCritical, "Select Process": Exit Sub
    If ItemName = "System" Then MsgBox "Please select a process to open.", vbCritical, "Select Process": Exit Sub
    
    'Kill Old Dump file
    Call KillOldDump
    'Save ProcessNumber for later use
    'ProcessNumber = lstProcess.ListIndex
    ProcessNumber = 0
    'Get the base Address
   ' mBaseAddress = GetModuleHandleA(lstProcess.List(lstProcess.ListIndex))
    'Get the current Address
   ' mCurrentAddress = GetModuleHandleA(lstProcess.List(lstProcess.ListIndex))

   ' ProccessName = lstProcess.List(lstProcess.ListIndex)
    ProccessName = ItemName
    'If Not InitProcessOpener(PIDs(lstProcess.ListIndex)) Then MsgBox "Could not open process. sorry :(", vbCritical, "Memory Editor Pro": Exit Sub
    If Not InitProcessOpener(Number) Then MsgBox "Could not open process. sorry :(", vbCritical, "Memory Editor Pro": Exit Sub
        'MsgBox Number
 'Open App.Path & "\dump.txt" For Binary Access Write Lock Write As #9
 'For addr = Lowerlimit To HighLimit Step 1000 '11536388
   ' Call ReadProcessMemory(myHandle, addr, Buffer, 1000, readlen)
   ' Put #9, , Buffer
 'Next
 'Close #9
 
    'ScrollBars setup
    Call frmMain.SetupScrollBars
    
    ProgramLoaded = True
    frmMain.PrintOffsetList mBaseAddress, mCurrentAddress, mMaxAddress
    Call frmMain.PrintMemoryHex

  
    If ProccessName = "" Then
        frmMain.Caption = "Memory Viewer & Dumper Version: " & Version & "  by VisualBasicZone.com"
    Else
        frmMain.Caption = "Memory Viewer & Dumper Version: " & Version & "  by VisualBasicZone.com - " & ProccessName
    End If
    frmMain.mnuFileDumpMemory.Enabled = True
    



Unload Me

End Sub

Sub KillOldDump()
On Error Resume Next
   ' Kill (App.Path & "\dump.txt")
End Sub

Private Sub Form_Load()
    cmdRefresh_Click
End Sub

Private Sub txtHigh_Change()
    If IsNumeric(txtHigh.Text) = False Then txtHigh.Text = 0
    
 
End Sub

Private Sub txtLow_Change()
    If IsNumeric(txtLow.Text) = False Then txtLow.Text = 0
    
End Sub
Private Sub Load_Filename(sExeName As String, ExeTitle As String, pid As Long, Optional KeyNumber As String)
'Load the Icon into the List View

ReDim glLargeIcons(lIcons)
ReDim glSmallIcons(lIcons)

On Error GoTo ErrFound


Dim lIndex

lIndex = "0"

'Get Icon from the File
Call ExtractIconEx(sExeName, lIndex, glLargeIcons(lIndex), glSmallIcons(lIndex), 1)

With picLarge
    Set .Picture = LoadPicture("")
     .AutoRedraw = True
    Call DrawIconEx(.hdc, 0, 0, glLargeIcons(lIndex), LARGE_ICON, LARGE_ICON, 0, 0, DI_NORMAL)
    'Debug.Print "Icon:" & glLargeIcons(lIndex)
     .Refresh
End With

Mykey = sExeName & "(" & "-" & KeyNumber & ")"
    If glLargeIcons(lIndex) <> 0 Then
        ImageList1.ListImages.Add , Mykey, picLarge.image
    Else
        ImageList1.ListImages.Add , Mykey, picSmall.image
    End If
    
txtMax = sExeName
Dim t As ListItem
' Add Icon to Listview

Set t = AppRun.ListItems.Add(, txtMax, ExeTitle, Mykey)
t.Tag = pid
t.ToolTipText = "PID: " & pid & " Filename: " & sExeName
'Debug.Print "pid-" & pid
ErrFound:

End Sub

Private Sub lblInfo_Click()
    Clipboard.SetText lblInfo.Caption
End Sub

Private Sub lstProcess_Click()
    MsgBox PIDs(lstProcess.ListIndex)
End Sub


Private Sub OptChoice1_Click()
     If OptChoice1.Value = True Then AppRun.View = lvwSmallIcon
End Sub

Private Sub OptChoice2_Click()
     If OptChoice2.Value = True Then AppRun.View = lvwReport
End Sub

Private Sub OptChoice3_Click()
    If OptChoice3.Value = True Then AppRun.View = lvwIcon
End Sub
