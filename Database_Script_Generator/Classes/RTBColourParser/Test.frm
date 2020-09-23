VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test RichTextBox Colour Parser"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   4200
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Query Analyser"
      Height          =   405
      Index           =   2
      Left            =   2820
      TabIndex        =   3
      Top             =   60
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Visual C"
      Height          =   405
      Index           =   1
      Left            =   1470
      TabIndex        =   2
      Top             =   60
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Visual Basic"
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1305
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5595
      Left            =   90
      TabIndex        =   0
      Top             =   540
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9869
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Test.frx":0000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mRTB As RTBColourParser

Private Sub Command1_Click(Index As Integer)
    With mRTB
        Select Case Index
        Case 0
            .EditorMode = emVB
            .LoadKeyWords App.Path & "\VBKeyWords.txt"
            
            '.LoadFile App.Path & "\RTBColourParser.cls", RichTextBox1
            .LoadFile "k:\dv8dev\lib\dvformat.bas", RichTextBox1
        
        Case 1
            .EditorMode = emVC
            .LoadKeyWords App.Path & "\VCKeyWords.txt"
            
            .LoadFile "C:\Program Files\Microsoft Visual Studio\MyProjects\TestConsole\TestConsole.cpp", RichTextBox1
            
        Case 2
            .EditorMode = emDefault
            .LoadKeyWords App.Path & "\QAKeyWords.txt"
            
            .LoadFile "\\dv3\data\dv8.sql", RichTextBox1
        End Select
    End With
End Sub

Private Sub Form_Load()
    Set mRTB = New RTBColourParser
End Sub


Private Sub mRTB_BeginLoading(Max As Long)
    ProgressBar1.Visible = True
    ProgressBar1.Max = Max
    ProgressBar1.Value = ProgressBar1.Min
End Sub


Private Sub mRTB_EndLoading()
    ProgressBar1.Visible = False
End Sub


Private Sub mRTB_Loading(Value As Long)
    ProgressBar1.Value = Value
End Sub


