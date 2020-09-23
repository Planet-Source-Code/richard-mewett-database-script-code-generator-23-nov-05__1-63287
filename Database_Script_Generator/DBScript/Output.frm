VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmOutput 
   Caption         =   "Script Output"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkHighlight 
      Caption         =   "Highlight Keywords"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   6510
      Value           =   1  'Checked
      Width           =   1995
   End
   Begin RichTextLib.RichTextBox rtfSource 
      Height          =   6315
      Left            =   2130
      TabIndex        =   1
      Top             =   75
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   11139
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Output.frx":0000
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   6345
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   11192
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblFile 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2130
      TabIndex        =   3
      Top             =   6450
      Width           =   7335
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_USER As Long = &H400
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72

Private mRTB As RTBColourParser
Public Sub ShowOutput(nMode As RTBEditorMode)
    Set mRTB = New RTBColourParser
             
    mRTB.EditorMode = nMode
    
    If nMode = emVB Then
        mRTB.LoadKeyWords App.Path & "\VBKeyWords.txt"
    Else
        mRTB.LoadKeyWords App.Path & "\QAKeyWords.txt"
    End If
    
    lvfiles_ItemClick lvFiles.SelectedItem
    
    Me.Show vbModal
End Sub


Private Sub chkHighlight_Click()
    If Not lvFiles.SelectedItem Is Nothing Then
        lvfiles_ItemClick lvFiles.SelectedItem
    End If
End Sub

Private Sub Form_Load()
    With lvFiles
        .ColumnHeaders.Add , , "File", 2000
        .View = lvwReport
    End With
    
    'Prevent WordWrap
    SendMessage rtfSource.hwnd, EM_SETTARGETDEVICE, 0, 0
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        lvFiles.Height = Me.Height - 960
        
        rtfSource.Height = Me.Height - 990
        rtfSource.Width = Me.Width - 2265
        
        chkHighlight.Top = Me.Height - 795
        lblFile.Top = Me.Height - 855
        lblFile.Width = Me.Width - 2265
    End If
End Sub


Private Sub lvfiles_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lblFile.Caption = "Loading file " & Item.Text & "..."
    lblFile.Refresh
    Screen.MousePointer = vbHourglass
    
    If chkHighlight.Value Then
        mRTB.LoadFile frmMain.txtFolder.Text & "\" & Item.Text, rtfSource
    Else
        With rtfSource
            .LoadFile frmMain.txtFolder.Text & "\" & Item.Text
            
            .SelStart = 1
            .SelLength = Len(.Text)
            .SelColor = vbBlack
        End With
    End If
    
    lblFile.Caption = "Viewing file " & Item.Text
    lblFile.Refresh
    Screen.MousePointer = vbDefault
End Sub


