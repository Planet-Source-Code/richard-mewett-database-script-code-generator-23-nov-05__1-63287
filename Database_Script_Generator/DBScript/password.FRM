VERSION 5.00
Begin VB.Form frmPW 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Database"
   ClientHeight    =   1605
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   1050
      TabIndex        =   1
      Top             =   330
      Width           =   3255
      Begin VB.TextBox txtPassword 
         Height          =   295
         IMEMode         =   3  'DISABLE
         Left            =   990
         MaxLength       =   14
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   270
         Width           =   2115
      End
      Begin VB.Label labPasword 
         Caption         =   "Password"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   315
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1740
      TabIndex        =   4
      Top             =   1155
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3045
      TabIndex        =   5
      Top             =   1155
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "password.frx":0000
      Top             =   300
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Database password required:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1050
      TabIndex        =   0
      Top             =   90
      Width           =   2490
   End
End
Attribute VB_Name = "frmPW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Me.Tag = 0
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Me.Tag = 1
    Me.Hide
End Sub

Private Sub Form_Activate()
    txtPassword.SetFocus
End Sub

