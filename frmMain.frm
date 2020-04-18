VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Memory Viewer & Dumper"
   ClientHeight    =   5385
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11025
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   11025
   StartUpPosition =   2  'CenterScreen
   Begin MemViewerDumper.vScrollXL vScroll1 
      Height          =   4935
      Left            =   10680
      TabIndex        =   3
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   8705
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":27A2
   End
   Begin VB.Timer tmrKeyDown 
      Interval        =   100
      Left            =   360
      Top             =   1080
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5130
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   7200
      Top             =   360
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Line lineAsciiH 
      BorderColor     =   &H00000000&
      Index           =   0
      Visible         =   0   'False
      X1              =   7080
      X2              =   7080
      Y1              =   300
      Y2              =   5220
   End
   Begin VB.Line lineHeight 
      Index           =   0
      X1              =   1520
      X2              =   1520
      Y1              =   300
      Y2              =   5400
   End
   Begin VB.Line lineWide 
      Index           =   0
      X1              =   1520
      X2              =   7005
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   175
      Left            =   1800
      Top             =   445
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   0
      X2              =   6960
      Y1              =   260
      Y2              =   260
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   1560
      X2              =   1560
      Y1              =   5280
      Y2              =   0
   End
   Begin VB.Label lblOffset 
      BackStyle       =   0  'Transparent
      Caption         =   "OffSet"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAttach 
         Caption         =   "&Attach To Process"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileDumpMemory 
         Caption         =   "Dump Memory"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsJump 
         Caption         =   "&Jump to Offset"
         Shortcut        =   ^J
      End
   End
   Begin VB.Menu mnuOptionS 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionGrid 
         Caption         =   "Grid Lines"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptOptions 
         Caption         =   "Options"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'Memory Editor Pro by vbgamer45
'##############################################

Dim CurrentOffset As Double
Dim ClickX As Long
Dim ClickY As Long
Dim Clicked As Boolean


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   'Scroll with keyboard

    If KeyCode = vbKeyDown Then
        If vScroll1.Value < vScroll1.Max Then
            vScroll1.Value = vScroll1.Value + 1
        End If
    End If
    If KeyCode = vbKeyUp Then
        If vScroll1.Value > 0 Then
            vScroll1.Value = vScroll1.Value - 1
        End If
    End If
End Sub


Private Sub Form_Load()
    Version = "0.02B"
    ProcessNumber = -1
    'Set color varibles
    cBackroundColor = vbWhite
    cOffSetColor = vbRed
    cHexColor = vbBlue
    ProgramLoaded = False

 
    DrawGridLines = True
    Me.Caption = "Memory Viewer & Dumper: " & Version & "  VisualBasicZone.com"
'
    'Get Settings For use later on
    cOffSetColor = GetSetting("Memory Editor Viewer", "Options", "OffSetColor", vbRed)
    cHexColor = GetSetting("Memory Editor Viewer", "Options", "HexColor", vbBlue)
    cBackroundColor = GetSetting("Memory Editor Viewer", "Options", "BackroundColor", vbWhite)
    cGridColor = GetSetting("Memory Editor Viewer", "Options", "GridColor", vbBlack)
   
    'Set The cursor box
    Shape1.Left = lineHeight(0).X1
    Shape1.Top = lineWide(0).Y1
    Shape1.Width = 350
    Shape2.Left = lineAsciiH(0).X1
    Shape2.Top = lineWide(0).Y1
    'Load edit boxes
    Dim AsciiB As Integer
    Dim CurX As Integer
    Dim CurY As Integer
 

    CurY = 300
    CurX = 7100
    
   ' For i = 1 To 50 '101 'Step 16
       ' For g = 1 To 16
       ' AsciiB = txtAscii.UBound + 1
       ' Load txtAscii(AsciiB)
       ' With txtAscii(AsciiB)
       '     .Text = ""
       '     .Visible = True
       '     .Left = CurX
       '     .Top = CurY
       '     .MaxLength = 1
       '     .Enabled = False
       '     .FontSize = 8
       '     .ForeColor = cHexColor
      '  End With
      '  CurX = CurX + 210
      '  Next
      '  CurY = CurY + 140
      '  CurX = 7100
    'Next
    
    For i = 1 To 50
        AsciiB = lineWide.UBound + 1
        Load lineWide(AsciiB)
        With lineWide(AsciiB)
            .X1 = lineWide(0).X1
            .X2 = lineWide(0).X2
            .Y1 = lineWide(AsciiB - 1).Y1 + 165
            .Y2 = lineWide(AsciiB - 1).Y2 + 165
            .Visible = False
        End With
    Next
    For i = 1 To 16
        AsciiB = lineHeight.UBound + 1
        Load lineHeight(AsciiB)
        With lineHeight(AsciiB)
            .X1 = lineHeight(AsciiB - 1).X1 + 320
            .X2 = lineHeight(AsciiB - 1).X2 + 320
            .Y1 = lineHeight(0).Y1
            .Y2 = lineHeight(0).Y2
            .Visible = False
        End With
    Next i
   
       For i = 1 To 16
        AsciiB = lineAsciiH.UBound + 1
        Load lineAsciiH(AsciiB)
        With lineAsciiH(AsciiB)
            .X1 = lineAsciiH(AsciiB - 1).X1 + 210
            .X2 = lineAsciiH(AsciiB - 1).X2 + 210
            .Y1 = lineAsciiH(0).Y1
            .Y2 = lineAsciiH(0).Y2
            .Visible = False
        End With
    Next i
    
  'Set up Form

   Me.BackColor = cBackroundColor 'vbWhite
   Me.FontName = "MS Sans Serif"
   Me.CurrentX = 0
   Me.CurrentY = 0
   Me.FontBold = True
   Me.ForeColor = cGridColor
   Me.FontSize = 10
   'Set Up Lines
   Me.Print "OffSet"
   Me.Line (0, 300)-(Me.Width, 300)
   Me.Line (1520, Me.Height)-(1520, 0)
   vScroll1.Height = Me.Height - vScroll1.Top - 950
   'Print Red Letters
   Me.CurrentX = 1600
   Me.CurrentY = 0
   Me.ForeColor = cOffSetColor 'vbRed
   Me.Print "0    1    2   3   4   5   6   7   8   9  A   B   C   D   E   F"
   'Hide Menu options not using
  
   'Status Text
   StatusBar1.SimpleText = "Please use Attach Process to begin."
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If ProgramLoaded = False Then Exit Sub
Clicked = True
ClickX = x
ClickY = y
        If Button = 2 Then
      
           
        
         
        Else
            On Error Resume Next
            For i = 0 To 55
                For g = 0 To 15
                If y >= lineWide(i).Y1 And y <= lineWide(i + 1).Y1 Then
                    If x >= lineHeight(g).X1 And x <= lineHeight(g + 1).X1 Then
                        Shape1.Top = lineWide(i).Y1
                        Shape1.Left = lineHeight(g).X1
                        Shape1.Width = lineHeight(g + 1).X1 - lineHeight(g).X1 + 20
                        Shape1.Visible = True
                        BoxXpos = Shape1.Left
                        BoxYpos = Shape1.Top
                        
                       ' MsgBox "Number: " & i * 16 + g
                        Exit Sub
                    End If
                End If
                Next g
            Next i
        
       
        'Box 2
            For i = 0 To 55
                For g = 0 To 15
                If y >= lineWide(i).Y1 And y <= lineWide(i + 1).Y1 Then
                    If x >= lineAsciiH(g).X1 And x <= lineAsciiH(g + 1).X1 Then
                        Shape2.Top = lineWide(i).Y1
                        Shape2.Left = lineAsciiH(g).X1
                        Shape2.Width = lineAsciiH(g + 1).X1 - lineAsciiH(g).X1 + 20
                        Shape2.Visible = True
                        BoxAsciiXpos = Shape2.Left
                        BoxAsciiYpos = Shape2.Top
                        Exit Sub
                    End If
                End If
                Next g
            Next i
        
        End If
        
        
    'End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Show Offste and information


If ProgramLoaded = False Then Exit Sub
If x < lineHeight(0).X1 Then Exit Sub
If y < lineWide(0).Y1 Then Exit Sub
'If X > lineHeight(lineHeight.UBound).X1 Then Exit Sub
If y > lineWide(lineWide.UBound).Y1 Then Exit Sub

Dim Number As Long


            On Error Resume Next
            For i = 0 To 55
                For g = 0 To 15
                If y >= lineWide(i).Y1 And y <= lineWide(i + 1).Y1 Then
                    If x >= lineHeight(g).X1 And x <= lineHeight(g + 1).X1 Then
                        Number = i * 16 + g + 1
                        CurrentOffset = BigDecToHex(OffsetArray(Number))
                       'MsgBox CurrentOffset
                       ' StatusBar1.SimpleText = "OffsetHex: " & BigDecToHex(txtAscii(Number).Tag) & " OffsetDec:" & txtAscii(Number).Tag & " Ascii:" & HexArray(Number) & " Hex:" & BigDecToHex(HexArray(Number)) & " BaseAddress:" & BigDecToHex(mBaseAddress)
                       StatusBar1.SimpleText = "OffsetHex: " & BigDecToHex(OffsetArray(Number)) & " OffsetDec:" & OffsetArray(Number) & " Ascii:" & HexArray(Number) & " Hex:" & BigDecToHex(HexArray(Number)) & " BaseAddress:" & BigDecToHex(mBaseAddress)
                        
                        Exit Sub
                    End If
                End If
                Next g
            Next i
            
If x >= lineAsciiH(0).X1 Then
          On Error Resume Next
            For i = 0 To 55
                For g = 0 To 15
                If y >= lineWide(i).Y1 And y <= lineWide(i + 1).Y1 Then
                    If x >= lineAsciiH(g).X1 And x <= lineAsciiH(g + 1).X1 Then
                        Number = i * 16 + g + 1
                         CurrentOffset = BigDecToHex(OffsetArray(Number))
                        StatusBar1.SimpleText = "OffsetHex: " & BigDecToHex(OffsetArray(Number)) & " OffsetDec:" & OffsetArray(Number) & " Ascii:" & HexArray(Number) & " Hex:" & BigDecToHex(HexArray(Number)) & " BaseAddress:" & BigDecToHex(mBaseAddress)
                        Exit Sub
                    End If
                End If
                Next g
            Next i
End If
            
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Shape1.Width = 350
    Shape1.Height = 175
    Clicked = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    End
End Sub

Private Sub Form_Resize()
On Error Resume Next
   Me.FontName = "MS Sans Serif"
   Me.ForeColor = cGridColor
   Me.Line (0, 300)-(Me.Width, 300)
   Me.Line (1520, Me.Height)-(1520, 0)
   vScroll1.Height = Me.Height - vScroll1.Top - 950
   Call DrawGrid
   Call PrintMemoryHex
End Sub



Private Sub mnuFileAttach_Click()
    frmSelectProcess.Show vbModal, Me
End Sub

Private Sub mnuFileDumpMemory_Click()
    frmDumpMem.Show vbModal, Me
End Sub

Private Sub mnuFileExit_Click()


    End
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub
Sub PrintOffsetList(BaseAddress As Long, CurrentAddr As Long, MaxAddr As Long)
On Error Resume Next
If ProgramLoaded = False Then Exit Sub
   Dim newAddr As Long, i As Long
   Me.Cls
   'Draw the grid
   Call DrawGrid
   Me.FontName = "MS Sans Serif"
   Me.CurrentX = 0
   Me.CurrentY = 0
   Me.FontBold = True
   Me.ForeColor = cGridColor
   Me.FontSize = 10
   Me.Print "OffSet"
   Me.Line (0, 300)-(Me.Width, 300)
   Me.Line (1520, Me.Height)-(1520, 0)
   'Me.Line (7000, Me.Height)-(7000, 0)
   vScroll1.Height = Me.Height - vScroll1.Top - 950
   'Print top header numbers
   Me.CurrentX = 1600
   Me.CurrentY = 0
   Me.ForeColor = cOffSetColor 'vbRed
   Me.Print "0    1    2   3   4   5   6   7   8   9   A   B   C   D   E   F"
    Me.FontName = "Lucida Console"
    Me.FontBold = True
    Me.FontSize = 8
    Me.CurrentX = 0
    Me.CurrentY = 300
    Me.ForeColor = cOffSetColor 'vbRed
    newAddr = CurrentAddr
  
    For i = 0 To Me.Height \ 120
        Me.Print " " + BigDecToHex(newAddr)
        newAddr = newAddr + 16
     
    Next


End Sub
Sub SetupScrollBars()
    vScroll1.Min = 0
    vScroll1.Max = (mMaxAddress - mBaseAddress) / 16 'mMaxAddress / 16
End Sub
Public Sub PrintMemoryHex()
    If ProgramLoaded = False Then Exit Sub
    If ProcessNumber = -1 Then Exit Sub
Dim Buffer As String * 1000
Dim readlen As Long


Call ReadProcessMemory(myHandle, mCurrentAddress, Buffer, 1000, readlen)
   ' HexBuffer = Buffer
    Me.CurrentY = 300
    Me.FontBold = False
    Me.FontSize = 8
    Me.ForeColor = cHexColor 'vbBlue
    Me.FontName = "Lucida Console"
    
Dim DATA As String
    On Error Resume Next
   
For i = 1 To 1001 Step 16

Me.CurrentX = 1600
DATA = ""
    DATA = DATA & BigDecToHex(Asc(Mid(Buffer, i, 1))) & " "
    HexArray(i) = Asc(Mid(Buffer, i, 1))
    For g = 1 To 15
        DATA = DATA & BigDecToHex(Asc(Mid(Buffer, i + g, 1))) & " "
        HexArray(i + g) = Asc(Mid(Buffer, i + g, 1))
    Next
    Print DATA
    
Next


DATA = ""
Me.CurrentY = 300
For i = 1 To 1001 Step 16

   DATA = ""
    If Asc(Mid(Buffer, i, 1)) < 32 Then
      DATA = DATA & ". "

      OffsetArray(i) = mCurrentAddress + i - 1
    Else
        DATA = DATA & Mid(Buffer, i, 1) & " "

        OffsetArray(i) = mCurrentAddress + i - 1
    End If

    For g = 1 To 15
        If Asc(Mid(Buffer, i + g, 1)) < 32 Then
            DATA = DATA & ". "
            If EditMode = True Then
            'txtAscii(i + g).Text = "."
            'txtAscii(i + g).Tag = mCurrentAddress + i - 1 + g
            'txtAscii(i + g).Refresh
            End If
            OffsetArray(i + g) = mCurrentAddress + i - 1 + g
        Else
            
            DATA = DATA & Mid(Buffer, i + g, 1) & " "
            If EditMode = True Then
            'txtAscii(i + g).Text = Mid(Buffer, i + g, 1)
            'txtAscii(i + g).Tag = mCurrentAddress + i - 1 + g '+ g + (i - 1) * 16
            'txtAscii(i + g).Refresh
            End If
            OffsetArray(i + g) = mCurrentAddress + i - 1 + g
        End If
    Next
'Me.CurrentY = 300
Me.CurrentX = 7100
If EditMode = False Then
    Me.Print DATA
End If

Next
'Draw the grid
'Call DrawGrid
End Sub

Private Sub mnuOptionGrid_Click()
    If mnuOptionGrid.Checked = True Then

        DrawGridLines = False
        Call DrawGrid
        Call PrintOffsetList(mBaseAddress, mCurrentAddress, mMaxAddress)
        Call PrintMemoryHex
        
        mnuOptionGrid.Checked = False
    Else

        DrawGridLines = True
        Call DrawGrid
        Call PrintOffsetList(mBaseAddress, mCurrentAddress, mMaxAddress)
        Call PrintMemoryHex
        mnuOptionGrid.Checked = True
    End If
End Sub

Private Sub mnuOptOptions_Click()
    frmOptions.Show vbModal, Me
End Sub



Private Sub mnuToolsJump_Click()
    If ProgramLoaded = False Then
        MsgBox "You need to attach to a process first.", vbInformation
    
    Else
        frmJump.Show
    End If
End Sub



Private Sub VScroll1_Change()
    If ProgramLoaded = False Then Exit Sub
    mCurrentAddress = mBaseAddress + (vScroll1.Value * 16#)
    PrintOffsetList mBaseAddress, mCurrentAddress, mMaxAddress
    Call PrintMemoryHex
End Sub

Private Sub VScroll1_Scroll()
    If ProgramLoaded = False Then Exit Sub
    mCurrentAddress = mBaseAddress + (vScroll1.Value * (16#))
    PrintOffsetList mBaseAddress, mCurrentAddress, mMaxAddress
    Call PrintMemoryHex
    
End Sub

Public Function IsWindowsNT() As Boolean
   Dim verinfo As OSVERSIONINFO
   verinfo.dwOSVersionInfoSize = Len(verinfo)
   If (GetVersionEx(verinfo)) = 0 Then Exit Function
   If verinfo.dwPlatformId = 2 Then IsWindowsNT = True
End Function

Sub DrawGrid()
    Dim i As Long
    If DrawGridLines = True Then
    'Draw Hex Grid
    For i = 0 To 81
        y = lineHeight(0).Y1 + (i * 165) - 165
        If i Mod 2 <> 0 Then
            Me.Line (lineWide(0).X1, y)-Step(5110, 154), &HC4B6B5, BF
        End If
    Next
     For i = 0 To 81
        y = lineAsciiH(0).Y1 + (i * 165) - 165
        If i Mod 2 <> 0 Then
            Me.Line (lineAsciiH(0).X1, y)-Step(3350, 154), &HC4B6B5, BF
        End If
    Next
    
        Me.ForeColor = cGridColor
        For i = 0 To lineWide.UBound
            'Me.Line (lineWide(i).X1, lineWide(i).Y1)-(lineWide(i).X2, lineWide(i).Y2)
            Me.Line (lineWide(i).X1, lineWide(i).Y1)-(vScroll1.Left, lineWide(i).Y2)
        Next
        For i = 0 To lineHeight.UBound
            
            'Me.Line (lineHeight(i).X1, lineHeight(i).Y1)-(lineHeight(i).X2, lineHeight(i).Y2)
            Me.Line (lineHeight(i).X1, lineHeight(i).Y1)-(lineHeight(i).X2, Me.Height)
        Next
   'Draw Ascii Grid
        For i = 0 To lineAsciiH.UBound
            Me.Line (lineAsciiH(i).X1, lineAsciiH(i).Y1)-(lineAsciiH(i).X2, Me.Height)
        Next

        
        
    Else
       ' Call Me.PrintOffsetList(mBaseAddress, mCurrentAddress, mMaxAddress)
        'Call Me.PrintMemoryHex
        
    End If
End Sub
