VERSION 5.00
Begin VB.Form frmMemoryProject 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Memory Project"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   Icon            =   "frmMemoryProject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter a title for this Memory Project"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   600
      Picture         =   "frmMemoryProject.frx":27A2
      Top             =   240
      Width           =   240
   End
End
Attribute VB_Name = "frmMemoryProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCreate_Click()
    If txtTitle.Text = "" Then
        MsgBox "You need to enter some kind of title for your project", vbInformation
        Exit Sub
    End If

    memProject.ProjectTitle = txtTitle.Text
    ProjectLoaded = True
    ProjectSaved = False
    frmMain.StatusBar1.SimpleText = "Use Attach Process to begin your memory project"
    Unload Me
    
End Sub
