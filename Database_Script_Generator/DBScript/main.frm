VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Database Script Generator"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6315
      Left            =   90
      TabIndex        =   8
      Top             =   1290
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11139
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Script"
      TabPicture(0)   =   "main.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cdlScript"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdInvert"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdTagAll"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdGenerate"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdClose"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lstTables"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "&Options"
      TabPicture(1)   =   "main.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(2)=   "Label7"
      Tab(1).Control(3)=   "Label8"
      Tab(1).Control(4)=   "Label9"
      Tab(1).Control(5)=   "Label10"
      Tab(1).Control(6)=   "txtADOReference"
      Tab(1).Control(7)=   "txtADOXReference"
      Tab(1).Control(8)=   "txtDAOReference"
      Tab(1).Control(9)=   "txtSQLCollate"
      Tab(1).Control(10)=   "cmdSave"
      Tab(1).Control(11)=   "cmdDefault"
      Tab(1).ControlCount=   12
      Begin VB.CommandButton cmdDefault 
         Caption         =   "Default"
         Height          =   345
         Left            =   -68160
         TabIndex        =   38
         Top             =   5880
         Width           =   1245
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   345
         Left            =   -66870
         TabIndex        =   39
         Top             =   5880
         Width           =   1245
      End
      Begin VB.TextBox txtSQLCollate 
         Height          =   315
         Left            =   -73470
         TabIndex        =   18
         Top             =   2340
         Width           =   7845
      End
      Begin VB.TextBox txtDAOReference 
         Height          =   315
         Left            =   -73470
         TabIndex        =   15
         Top             =   1530
         Width           =   7845
      End
      Begin VB.TextBox txtADOXReference 
         Height          =   315
         Left            =   -73470
         TabIndex        =   13
         Top             =   1140
         Width           =   7845
      End
      Begin VB.TextBox txtADOReference 
         Height          =   315
         Left            =   -73470
         TabIndex        =   11
         Top             =   750
         Width           =   7845
      End
      Begin VB.Frame Frame1 
         Height          =   5445
         Left            =   4950
         TabIndex        =   20
         Top             =   360
         Width           =   4425
         Begin VB.ListBox lstFormat 
            Height          =   1620
            Left            =   120
            TabIndex        =   28
            Top             =   2070
            Width           =   4185
         End
         Begin VB.CheckBox chkFieldAttributes 
            Caption         =   "Field Attributes"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   1260
            Width           =   1545
         End
         Begin VB.CheckBox chkTables 
            Caption         =   "Tables"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   780
            Width           =   1545
         End
         Begin VB.CheckBox chkIndexes 
            Caption         =   "Indexes"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   1020
            Width           =   1545
         End
         Begin VB.CheckBox chkFieldProperties 
            Caption         =   "Field Properties"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   1500
            Width           =   1545
         End
         Begin VB.TextBox txtFolder 
            Height          =   285
            Left            =   780
            TabIndex        =   34
            Top             =   5040
            Width           =   3165
         End
         Begin VB.CommandButton cmdFolder 
            Caption         =   "..."
            Height          =   285
            Left            =   3960
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   5040
            Width           =   375
         End
         Begin VB.CheckBox chkCreateScriptPerTable 
            Caption         =   "Create Script per Table"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   3870
            Width           =   2865
         End
         Begin VB.CheckBox chkDropTable 
            Caption         =   "DROP TABLE Statement"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   4590
            Width           =   2385
         End
         Begin VB.CheckBox chkCreateDatabaseModule 
            Caption         =   "Create Database Module"
            CausesValidation=   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   4110
            Width           =   2385
         End
         Begin VB.CheckBox chkViewSystemTables 
            Caption         =   "View System Tables"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   2865
         End
         Begin VB.CheckBox chkCreateProjectFile 
            Caption         =   "Create Project File"
            CausesValidation=   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   4350
            Width           =   2385
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Script Objects:"
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
            Left            =   120
            TabIndex        =   21
            Top             =   210
            Width           =   1275
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Folder"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   5070
            Width           =   435
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Script Properties:"
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
            Left            =   120
            TabIndex        =   27
            Top             =   1830
            Width           =   1485
         End
      End
      Begin VB.ListBox lstTables 
         Height          =   5325
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   19
         Top             =   450
         Width           =   4725
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   345
         Left            =   8130
         TabIndex        =   41
         Top             =   5880
         Width           =   1245
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
         Height          =   345
         Left            =   6840
         TabIndex        =   40
         Top             =   5880
         Width           =   1245
      End
      Begin VB.CommandButton cmdTagAll 
         Caption         =   "Tag All"
         Height          =   345
         Left            =   120
         TabIndex        =   36
         Top             =   5850
         Width           =   1245
      End
      Begin VB.CommandButton cmdInvert 
         Caption         =   "Invert"
         Height          =   345
         Left            =   1410
         TabIndex        =   37
         Top             =   5850
         Width           =   1245
      End
      Begin MSComDlg.CommonDialog cdlScript 
         Left            =   4020
         Top             =   5820
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "VB6 Project:"
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
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "SQL Server:"
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
         Left            =   -74880
         TabIndex        =   16
         Top             =   2100
         Width           =   1050
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Collate"
         Height          =   195
         Left            =   -74880
         TabIndex        =   17
         Top             =   2400
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "DAO Reference"
         Height          =   195
         Left            =   -74880
         TabIndex        =   14
         Top             =   1590
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ADOX Reference"
         Height          =   195
         Left            =   -74880
         TabIndex        =   12
         Top             =   1200
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ADO Reference"
         Height          =   195
         Left            =   -74880
         TabIndex        =   10
         Top             =   810
         Width           =   1140
      End
   End
   Begin VB.Frame fraDatabase 
      Caption         =   "Database"
      Height          =   1185
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   9495
      Begin MSComDlg.CommonDialog cdlDatabase 
         Left            =   3990
         Top             =   660
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.OptionButton optADO 
         Caption         =   "ADO"
         Height          =   195
         Left            =   2670
         TabIndex        =   4
         Top             =   810
         Width           =   795
      End
      Begin VB.OptionButton optDAO 
         Caption         =   "DAO"
         Height          =   195
         Left            =   1740
         TabIndex        =   3
         Top             =   810
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.TextBox txtDatabase 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   9225
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   345
         Left            =   5550
         TabIndex        =   5
         Top             =   690
         Width           =   1245
      End
      Begin VB.CommandButton cmdBuild 
         Caption         =   "Build..."
         Height          =   345
         Left            =   6840
         TabIndex        =   6
         Top             =   690
         Width           =   1245
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   345
         Left            =   8130
         TabIndex        =   7
         Top             =   690
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Connection Method:"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   810
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'#############################################################################################################################
'Title:     DBScript (An Tool for generating Source Code from a database)
'Author:    Richard Mewett
'Created:   01/06/05
'Version:   1.0.2 (23rd November 2005)

'Copyright © 2005 Richard Mewett. All rights reserved.

'## Am planning to add more languages (C++, VB.NET etc) and other DBMS such as MySQL ##

'Updates (dd/mm/yy):
'23 Nov 05  - Fixed syntax bug with ADO code generation (unwanted parenthesis on Open Recordset) - Thanks Bhupendra Aole!

'This is based on one of my old Submissions (DBCode). Major difference is more flexibility
'and ADO. The primary purpose of the original code was to assist me in upsizing a Jet based
'system to SQL Server.

'NOTE:
'This software is provided "as-is," without any express or implied warranty.
'In no event shall the author be held liable for any damages arising from the
'use of this software.
'If you do not agree with these terms, do not install "DBScript". Use of
'the program implicitly means you have agreed to these terms.
'
'Permission is granted to anyone to use this software for any purpose,
'including commercial use, and to alter and redistribute it, provided that
'the following conditions are met:
'
'1. All redistributions of source code files must retain all copyright
'   notices that are currently in place, and this list of conditions without
'   any modification.
'
'2. All redistributions in binary form must retain all occurrences of the
'   above copyright notice and web site addresses that are currently in
'   place (for example, in the About boxes).
'
'3. Modified versions in source or binary form must be plainly marked as
'   such, and must not be misrepresented as being the original software.

'########################################################################################
'Code Declarations
Private Enum ConvertTypeEnum
    ctSQLServer = 0
    ctVB = 1
End Enum

Private Const SCRIPT_QUERY_ANALYSER = 1
Private Const SCRIPT_VBCODE = 2
Private Const SCRIPT_VBCLASS_MODULE = 3

Private Const PREFIX_CLASS = "c"
Private Const PREFIX_MODULE = "m"

Private Const DEF_ADOREF = "Reference=*\G{00000205-0000-0010-8000-00AA006D2EA4}#2.5#0#..\..\..\..\Program Files\Common Files\system\ado\msado25.tlb#Microsoft ActiveX Data Objects 2.5 Library"
Private Const DEF_ADOXREF = "Reference=*\G{00000600-0000-0010-8000-00AA006D2EA4}#2.5#0#..\..\..\..\Program Files\Common Files\System\ado\msADOX.dll#Microsoft ADO Ext. 2.5 for DDL and Security"
Private Const DEF_DAOREF = "Reference=*\G{00025E01-0000-0000-C000-000000000046}#5.0#0#..\..\..\..\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll#Microsoft DAO 3.6 Object Library"

Private Const DEF_COLLATE = "Latin1_General_CI_AS"

Private mDB As DAO.Database
Private mCN As ADODB.Connection

Private Function ADO_GetConnection(sConnection As String, Optional sPassword As String, Optional bExclusive As Boolean) As ADODB.Connection
    '////////////////////////////////////////////////////////////////////////////////////
    'Return an ADO connection to a database. Prompts for password if necessary
    '////////////////////////////////////////////////////////////////////////////////////
    
    Dim cn As ADODB.Connection
    Dim lErr As Long
    Dim sPW As String
    
    On Local Error GoTo ADO_GetConnectionError
   
    Set cn = New ADODB.Connection
   
    If InStr(sConnection, ";") > 0 Then
        'SQL Server ODBC String
        If Left$(UCase$(sConnection), 5) = "ODBC;" Then
            cn.Open Mid$(sConnection, 6)
        Else
            cn.Open sConnection
        End If
    Else
        sPW = sPassword
        Do
            lErr = 0
            
            With cn
                .Provider = "Microsoft.Jet.OLEDB.4.0; Data Source=" & sConnection
                
                If Len(sPW) > 0 Then
                    .Properties("Jet OLEDB:Database Password") = sPW
                End If
                .Open
            End With
            
            If lErr = 0 Then
                Exit Do
            End If
        Loop
    End If
    
    Set ADO_GetConnection = cn
    Exit Function
    
ADO_GetConnectionError:
    lErr = Err.Number
    
    Select Case lErr
    Case -2147467259
        MsgBox "The database does not exist. It may have been moved or renamed.", vbCritical
    
    Case -2147217843
        With frmPW
            .txtPassword.Text = ""
            .Show vbModal
            
            If Val(.Tag) = 1 Then
                sPW = .txtPassword.Text
                Resume Next
            Else
                Exit Function
            End If
        End With
    
    Case Else
        MsgBox Err.Description, vbCritical
        
    End Select
End Function

Private Sub ADO_LoadObjects()
    Dim cat As ADOX.Catalog
    Dim tbl As ADOX.Table
    
    On Local Error GoTo ADO_LoadObjectsError
    
    Screen.MousePointer = vbHourglass
    
    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = mCN
      
    With lstTables
        .Clear
      
        For Each tbl In cat.Tables
            If (tbl.Type = "TABLE") Or chkViewSystemTables.Value Then
                .AddItem tbl.Name
            End If
        Next tbl
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ADO_LoadObjectsError:
    MsgBox Err.Description, vbCritical
End Sub

Private Function ADO_MapDataType(DataType As Long, ConvertType As ConvertTypeEnum) As String
    '////////////////////////////////////////////////////////////////////////////////////
    'Maps a ADO DataType to an SQL Server Query Analyser or VB one
    '////////////////////////////////////////////////////////////////////////////////////
 
    If ConvertType = ctSQLServer Then
        Select Case DataType
        Case adTinyInt
            ADO_MapDataType = "tinyint"
        Case adSmallInt
            ADO_MapDataType = "smallint"
        Case adInteger
            ADO_MapDataType = "int"
        Case adBigInt
            ADO_MapDataType = "bigint"
        Case adSingle
            ADO_MapDataType = "real"
        Case adDouble
            ADO_MapDataType = "float"
        Case adCurrency
            ADO_MapDataType = "money"
        Case adDecimal
            ADO_MapDataType = "decimal"
        Case adNumeric
            ADO_MapDataType = "numeric"
        Case adBoolean
            ADO_MapDataType = "bit"
        Case adVariant
            ADO_MapDataType = "variant"
        Case adDate
            ADO_MapDataType = "smalldatetime"
        Case adDBTime
            ADO_MapDataType = "datetime"
        Case adDBTimeStamp
            ADO_MapDataType = "dbtimestamp"
        Case adChar
            ADO_MapDataType = "char"
        Case adVarChar
            ADO_MapDataType = "varchar"
        Case adBinary
            ADO_MapDataType = "binary"
        Case adVarBinary
            ADO_MapDataType = "varbinary"
        Case adLongVarBinary
            ADO_MapDataType = "varbinary"
        End Select
    Else
        Select Case DataType
        Case adTinyInt, adUnsignedTinyInt
            ADO_MapDataType = "Byte"
        Case adSmallInt, adUnsignedSmallInt
            ADO_MapDataType = "Integer"
        Case adInteger, adUnsignedInt, adBigInt, adUnsignedBigInt
            ADO_MapDataType = "Long"
        Case adSingle
            ADO_MapDataType = "Single"
        Case adDouble
            ADO_MapDataType = "Double"
        Case adCurrency
            ADO_MapDataType = "Currency"
        Case adBoolean
            ADO_MapDataType = "Boolean"
        Case adDate, adDBDate, adDBTime
            ADO_MapDataType = "Date"
        Case adBSTR, adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar
            ADO_MapDataType = "String"
        Case Else
            ADO_MapDataType = "Variant"
        End Select
    End If
End Function

Private Sub ADO_ScriptQA(dbTbl As ADOX.Table, nHandle As Integer, bScriptPerTable As Boolean)
    '////////////////////////////////////////////////////////////////////////////////////
    'Create SQL script for Query Analyser to generate the source DAO Table
    '////////////////////////////////////////////////////////////////////////////////////
    
    Dim fld As ADOX.Column
    Dim idx As ADOX.Index
    Dim bMemo As Boolean
    Dim bPrimary As Boolean
    Dim sBuffer As String
    Dim sText As String
    Dim sTab As String

    On Local Error GoTo ScriptQAError
    sTab = Space$(4)

    If bScriptPerTable Then
        nHandle = FreeFile
        Open txtFolder.Text & "\" & TblName(dbTbl.Name) & ".sql" For Output As #nHandle
        
        frmOutput.lvFiles.ListItems.Add , , TblName(dbTbl.Name) & ".sql"
    End If

    Print #nHandle, "-- " & String(80, "#")
                
    If chkDropTable.Value Then
        Print #nHandle, "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[" & TblName(dbTbl.Name) & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
        Print #nHandle, "drop table [dbo].[" & TblName(dbTbl.Name) & "]"
        Print #nHandle, "GO"
        Print #nHandle, ""
    End If
    
    If chkTables.Value Then
        Print #nHandle, "CREATE TABLE [dbo].[" & TblName(dbTbl.Name) & "] ("
        
        bMemo = False
        sBuffer = ""
        
        For Each fld In dbTbl.Columns
            If fld.Properties("Autoincrement").Value Then
                sText = "[" & fld.Name & "] [int] IDENTITY (1,1) NOT NULL"
            Else
                Select Case fld.Type
                Case dbText
                    sText = "[" & fld.Name & "] [nvarchar] (" & fld.DefinedSize & ") COLLATE " & txtSQLCollate.Text
                Case dbMemo
                    bMemo = True
                    sText = "[" & fld.Name & "] [ntext] COLLATE " & txtSQLCollate.Text
                Case Else
                    sText = "[" & fld.Name & "] [" & DAO_MapDataType(fld.Type, ctSQLServer) & "]"
                End Select
                
                If fld.Properties("Nullable").Value Then
                    sText = sText & " NULL"
                Else
                    sText = sText & " NOT NULL"
                End If
            End If
            
            If Len(sBuffer) = 0 Then
                sBuffer = sTab & sText
            Else
                sBuffer = sBuffer & " ," & vbCrLf & sTab & sText
            End If
        Next fld
        
        Print #nHandle, sBuffer
        
        If bMemo Then
            Print #nHandle, ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
        Else
            Print #nHandle, ") ON [PRIMARY]"
        End If
        Print #nHandle, "GO"
        Print #nHandle, ""
    End If
    
    If chkIndexes.Value Then
        For Each idx In dbTbl.Indexes
            If IsJetDB() Then
                bPrimary = idx.PrimaryKey
            End If
            
            If bPrimary Then
                Print #nHandle, "ALTER TABLE [dbo].[" & TblName(dbTbl.Name) & "] WITH NOCHECK ADD"
                Print #nHandle, "CONSTRAINT [PK_" & TblName(dbTbl.Name) & "] PRIMARY KEY  CLUSTERED"
                Print #nHandle, "("
                
                sBuffer = ""
                For Each fld In idx.Columns
                    If Len(sBuffer) = 0 Then
                        sBuffer = "[" & fld.Name & "]"
                    Else
                        sBuffer = sBuffer & " ," & vbCrLf & "[" & fld.Name & "]"
                    End If
                Next fld
                
                Print #nHandle, sBuffer
                Print #nHandle, ") ON [PRIMARY]"
                Print #nHandle, "GO"
                Print #nHandle, ""
            Else
                Print #nHandle, "CREATE NONCLUSTERED INDEX [IK_" & idx.Name & "] ON [dbo].[" & TblName(dbTbl.Name) & "]"
                Print #nHandle, "("
                
                sBuffer = ""
                For Each fld In idx.Columns
                    If Len(sBuffer) = 0 Then
                        sBuffer = "[" & fld.Name & "]"
                    Else
                        sBuffer = sBuffer & " ," & vbCrLf & "[" & fld.Name & "]"
                    End If
                Next fld
                
                Print #nHandle, sBuffer
                Print #nHandle, ") ON [PRIMARY]"
                Print #nHandle, "GO"
                Print #nHandle, ""
            End If
        Next idx
    End If
    
    If chkFieldProperties.Value Then
        sBuffer = ""
    
        For Each fld In dbTbl.Columns
            If fld.Properties("Autoincrement").Value Then
                'No default allowed
            Else
                Select Case fld.Type
                Case dbText
                    sText = "CONSTRAINT [DF_" & TblName(dbTbl.Name) & "_" & fld.Name & "] DEFAULT '' FOR [" & fld.Name & "]"
                Case dbInteger, dbLong, dbCurrency, dbSingle, dbDouble
                    If fld.Properties("Default").Value <> 0 Then
                        sText = "CONSTRAINT [DF_" & TblName(dbTbl.Name) & "_" & fld.Name & "] DEFAULT (" & fld.Properties("Default").Value & ") FOR [" & fld.Name & "]"
                    Else
                        sText = "CONSTRAINT [DF_" & TblName(dbTbl.Name) & "_" & fld.Name & "] DEFAULT (0) FOR [" & fld.Name & "]"
                    End If
                Case dbBoolean
                    If Val(fld.Properties("Default").Value) <> 0 Then
                        sText = "CONSTRAINT [DF_" & TblName(dbTbl.Name) & "_" & fld.Name & "] DEFAULT (1) FOR [" & fld.Name & "]"
                    Else
                        sText = "CONSTRAINT [DF_" & TblName(dbTbl.Name) & "_" & fld.Name & "] DEFAULT (0) FOR [" & fld.Name & "]"
                    End If
                Case Else
                    sText = ""
                End Select
            
                If Len(sText) > 0 Then
                    If Len(sBuffer) = 0 Then
                        sBuffer = sTab & sText
                    Else
                        sBuffer = sBuffer & " ," & vbCrLf & sTab & sText
                    End If
                End If
            End If
        Next fld
    
        If Len(sBuffer) > 0 Then
            Print #nHandle, "ALTER TABLE [dbo].[" & TblName(dbTbl.Name) & "] WITH NOCHECK ADD"
            Print #nHandle, sBuffer
            Print #nHandle, "GO"
            Print #nHandle, ""
        End If
    End If
    
    If bScriptPerTable Then
        Close #nHandle
        nHandle = 0
    End If

    Exit Sub

ScriptQAError:
    MsgBox Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub ADO_ScriptVBClass(dbTbl As ADOX.Table)
    '////////////////////////////////////////////////////////////////////////////////////
    'Create a VB6 Class Module from an ADO Table
    '////////////////////////////////////////////////////////////////////////////////////
    
    Dim dbFld As ADOX.Column
    Dim rs As ADODB.Recordset
    
    Dim nHandle As Integer
    Dim sComment As String
    Dim sField As String
    Dim sTab As String
    
    On Local Error GoTo ScriptVBClassError
    
    sTab = Space$(4)

    nHandle = FreeFile
    Open txtFolder.Text & "\" & dbTbl.Name & ".cls" For Output As #nHandle
    
    frmOutput.lvFiles.ListItems.Add , , dbTbl.Name & ".cls"
    
    WriteClassHeader nHandle, dbTbl.Name
    Print #nHandle, "'Source Database: " & txtDatabase.Text
    Print #nHandle, "'Source Table: " & dbTbl.Name
    Print #nHandle, ""
    
    For Each dbFld In dbTbl.Columns
        sField = FormatName(dbFld.Name)
        
        Select Case dbFld.Type
        Case dbBinary
            sComment = " '(Binary Field)"
        Case dbMemo
            sComment = " '(Memo Field)"
        Case Else
            sComment = ""
        End Select
        
        Print #nHandle, "Private m" & sField & " As " & ADO_MapDataType(dbFld.Type, ctVB) & sComment
    Next dbFld
    
    Print #nHandle, ""
    Print #nHandle, "Public Sub LoadData (cn as ADODB.Connection)"
    Print #nHandle, Space$(4) & "Dim rs As New ADODB.RecordSet"
    Print #nHandle, Space$(4) & "Dim SQL As String"
    Print #nHandle, ""
    Print #nHandle, Space$(4) & "'\\ Edit SQL WHERE as required \\"
    Print #nHandle, Space$(4) & "SQL = " & Chr$(34) & "SELECT * FROM " & dbTbl.Name & Chr$(34)
    Print #nHandle, Space$(4) & "rs.Open SQL, cn, adOpenKeyset"
    Print #nHandle, Space$(4) & "With rs"
    Print #nHandle, Space$(8) & "If .RecordCount > 0 Then"
    
    For Each dbFld In dbTbl.Columns
        sField = FormatName(dbFld.Name)
        Print #nHandle, Space$(12) & "m" & sField & " = !" & dbFld.Name
    Next dbFld
    
    Print #nHandle, Space$(8) & "End If"
    Print #nHandle, Space$(8) & ".Close"
    Print #nHandle, Space$(4) & "End With"
    Print #nHandle, Space$(4) & "Set rs = Nothing"
    Print #nHandle, "End Sub"
    Print #nHandle, ""
    
    Print #nHandle, "Public Sub SaveData (cn as ADODB.Connection)"
    Print #nHandle, Space$(4) & "Dim rs As New ADODB.RecordSet"
    Print #nHandle, Space$(4) & "Dim SQL As String"
    Print #nHandle, ""
    Print #nHandle, Space$(4) & "'\\ Edit SQL WHERE as required \\"
    Print #nHandle, Space$(4) & "SQL = " & Chr$(34) & "SELECT * FROM " & dbTbl.Name & Chr$(34)
    Print #nHandle, Space$(4) & "rs.Open SQL, cn, adOpenKeyset"
    Print #nHandle, Space$(4) & "With rs"
    Print #nHandle, Space$(8) & "If .RecordCount = 0 Then"
    Print #nHandle, Space$(12) & ".AddNew"
    Print #nHandle, Space$(8) & "End If"
    
    For Each dbFld In dbTbl.Columns
        If Not dbFld.Properties("Autoincrement").Value Then
            sField = FormatName(dbFld.Name)
            If dbFld.Type = dbDate Then
                Print #nHandle, Space$(8) & "If IsDate(m" & sField & ") Then"
                Print #nHandle, Space$(12) & "!" & dbFld.Name & " = m" & sField
                Print #nHandle, Space$(8) & "Else"
                Print #nHandle, Space$(12) & "!" & dbFld.Name & " = Empty"
                Print #nHandle, Space$(8) & "End If"
            Else
                Print #nHandle, Space$(8) & "!" & dbFld.Name & " = m" & sField
            End If
        End If
    Next dbFld
    
    Print #nHandle, Space$(8) & ".Update"
    Print #nHandle, Space$(8) & ".Close"
    Print #nHandle, Space$(4) & "End With"
    Print #nHandle, Space$(4) & "Set rs = Nothing"
    Print #nHandle, "End Sub"
    Print #nHandle, ""
    
    For Each dbFld In dbTbl.Columns
        sField = FormatName(dbFld.Name)
        Print #nHandle, "Public Property Get " & sField & " () As " & ADO_MapDataType(dbFld.Type, ctVB)
        Print #nHandle, sTab & sField & " = m" & sField
        Print #nHandle, "End Property"
        Print #nHandle, ""
        Print #nHandle, "Public Property Let " & sField & " (NewValue As " & ADO_MapDataType(dbFld.Type, ctVB) & ")"
        Print #nHandle, sTab & "m" & sField & " = NewValue"
        Print #nHandle, "End Property"
        Print #nHandle, ""
    Next dbFld
    
    Close #nHandle
    Exit Sub

ScriptVBClassError:
    MsgBox Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub ADO_ScriptVBCode(dbTbl As ADOX.Table, nHandle As Integer, bScriptPerTable As Boolean)
    '////////////////////////////////////////////////////////////////////////////////////
    'Create VB6 code to generate the source  ADO Table
    '////////////////////////////////////////////////////////////////////////////////////
    
    Dim dbFld As ADOX.Column
    Dim fldLU As ADOX.Column
    Dim dbIdx As ADOX.Index
    Dim sBuffer As String
    Dim sText As String
    Dim sTab As String
    
    On Local Error GoTo ADO_ScriptVBCodeError
    
    sTab = Space$(4)

    If bScriptPerTable Then
        nHandle = FreeFile
        Open txtFolder.Text & "\" & dbTbl.Name & ".bas" For Output As #nHandle
        WriteModuleHeader nHandle, dbTbl.Name
        
        frmOutput.lvFiles.ListItems.Add , , dbTbl.Name & ".bas"
    End If
    
    Print #nHandle, "Public Sub Create_" & Replace(dbTbl.Name, " ", "_") & "(dbCat As ADOX.Catalog)"
    
    Print #nHandle, sTab & "Dim dbCol As ADOX.Column"
    Print #nHandle, sTab & "Dim dbIdx As ADOX.Index"
    Print #nHandle, sTab & "Dim dbTbl As ADOX.Table"
    
    Print #nHandle, ""
    Print #nHandle, sTab & "'Source Database: " & txtDatabase.Text
    Print #nHandle, ""

    Print #nHandle, sTab & "'" & String(120, "#")
    Print #nHandle, sTab & "'Code to generate Objects for Table: " & dbTbl.Name

    If chkTables.Value Then
        Print #nHandle, sTab & "Set dbTbl = New ADOX.Table"
        Print #nHandle, sTab & "With dbTbl"
        Print #nHandle, sTab & sTab & "Set .ParentCatalog = dbCat"
        Print #nHandle, sTab & sTab & ".Name = " & Chr$(34) & dbTbl.Name & Chr$(34) & ""
        Print #nHandle, ""
        
        For Each dbFld In dbTbl.Columns
            If (dbFld.Type = dbText) Then
                Print #nHandle, sTab & sTab & ".Columns.Append " & Chr$(34) & dbFld.Name & Chr$(34) & "," & GetADOConstant(dbFld.Type) & "," & dbFld.DefinedSize
            Else
                Print #nHandle, sTab & sTab & ".Columns.Append " & Chr$(34) & dbFld.Name & Chr$(34) & "," & GetADOConstant(dbFld.Type)
            End If
            
            If chkFieldAttributes.Value And dbFld.Attributes <> 0 Then
                Print #nHandle, sTab & sTab & ".Columns(" & Chr$(34) & dbFld.Name & Chr$(34) & ").Attributes = " & dbFld.Attributes
            End If
            
            'If chkFieldProperties.Value Then
            '    If dbFld.AllowZeroLength Then
            '        Print #nHandle, sTab & sTab & ".Columns(" & Chr$(34) & dbFld.Name & Chr$(34) & ").Properties(" & Chr$(34) & "Nullable" & Chr$(34) & ") = " & True
            '    End If
            '    If dbFld.Required Then
            '        Print #nHandle, sTab & sTab & ".Columns(" & Chr$(34) & dbFld.Name & Chr$(34) & ").Properties(" & Chr$(34) & "Required" & Chr$(34) & ") = " & True
            '    End If
            'End If
        Next dbFld
        
        Print #nHandle, sTab & "End With"
        Print #nHandle, ""
        Print #nHandle, sTab & "dbCat.Tables.Append dbTbl"
    Else
        Print #nHandle, sTab & "Set dbTbl = dbCat.Tables(" & Chr$(34) & dbTbl.Name & Chr$(34) & ")"
    End If

    If chkIndexes.Value Then
        For Each dbIdx In dbTbl.Indexes
            Print #nHandle, ""
            Print #nHandle, sTab & "Set dbIdx = New ADOX.Index"
            Print #nHandle, sTab & "With dbIdx"
            Print #nHandle, sTab & sTab & ".Name = " & Chr$(34) & dbIdx.Name & Chr$(34) & ""
            If dbIdx.PrimaryKey Then
                Print #nHandle, sTab & sTab & ".PrimaryKey = True"
            End If
            If dbIdx.Unique Then
                Print #nHandle, sTab & sTab & ".Unique = True"
            End If

            For Each dbFld In dbIdx.Columns
                Set fldLU = dbTbl.Columns(dbFld.Name)

                If (fldLU.Type = dbText) Then
                    Print #nHandle, sTab & sTab & ".Columns.Append " & Chr$(34) & dbFld.Name & Chr$(34) & "," & GetADOConstant(fldLU.Type) & "," & fldLU.DefinedSize
                Else
                    Print #nHandle, sTab & sTab & ".Columns.Append " & Chr$(34) & dbFld.Name & Chr$(34) & "," & GetADOConstant(fldLU.Type)
                End If
            Next dbFld

            Print #nHandle, sTab & "End With"
            Print #nHandle, sTab & "dbTbl.Indexes.Append dbIdx"
        Next dbIdx
    End If
    
    Print #nHandle, "End Sub"

    If bScriptPerTable Then
        Close #nHandle
        nHandle = 0
    End If
    Exit Sub

ADO_ScriptVBCodeError:
    MsgBox Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub chkViewSystemTables_Click()
    If optDAO.Value Then
        If Not mDB Is Nothing Then
            DAO_LoadObjects
        End If
    Else
        If Not mCN Is Nothing Then
            ADO_LoadObjects
        End If
    End If
End Sub

Private Sub cmdBrowse_Click()
    On Local Error GoTo BrowseError

    With cdlDatabase
        .Filter = "Jet Database (*.mdb)|*.mdb|All Files (*.*)|*.*"
        .Flags = cdlOFNHideReadOnly
        .ShowOpen
        
        txtDatabase.Text = .Filename
    End With
    
    cmdConnect.Value = True
    Exit Sub
    
BrowseError:
    Exit Sub
End Sub

Private Sub cmdBuild_Click()
    With frmDBConnection
        .ConnectionString = txtDatabase.Text
        .Show vbModal
        
        If Len(.ConnectionString) <> 0 Then
            txtDatabase.Text = .ConnectionString
        End If
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdConnect_Click()
    If cmdConnect.Caption = "Connect" Then
        If optDAO.Value Then
            Set mDB = DAO_GetConnection(txtDatabase.Text)
            If Not mDB Is Nothing Then
                DAO_LoadObjects
            End If
        Else
            Set mCN = ADO_GetConnection(txtDatabase.Text)
            If Not mCN Is Nothing Then
                ADO_LoadObjects
            End If
        End If
    Else
        If optDAO.Value Then
            mDB.Close
            Set mDB = Nothing
        Else
            mCN.Close
            Set mCN = Nothing
        End If
    End If
    
    SetState
End Sub

Private Sub cmdDefault_Click()
    txtADOReference.Text = DEF_ADOREF
    txtADOXReference.Text = DEF_ADOXREF
    txtDAOReference.Text = DEF_DAOREF
    
    txtSQLCollate.Text = DEF_COLLATE
End Sub

Private Sub cmdFolder_Click()
    Dim sFolder As String
    
    sFolder = GetFolder(Me.hwnd, "Scripts Folder:", txtFolder.Text)
    If Len(sFolder) > 0 Then
        txtFolder.Text = sFolder
    End If
End Sub

Private Sub cmdGenerate_Click()
    '////////////////////////////////////////////////////////////////////////////////////
    'Create Scripts for selected Tables
    '////////////////////////////////////////////////////////////////////////////////////
    
    Dim cat As ADOX.Catalog
    Dim nHandle As Integer
    Dim nH As Integer
    Dim nIndex As Integer
    Dim nPos As Integer
    Dim sTemp As String
    Dim nScriptType As Integer
    
    txtFolder.Text = Trim$(txtFolder.Text)
    If Len(txtFolder.Text) = 0 Then
        MsgBox "You must specify an output folder.", vbExclamation
        txtFolder.SetFocus
        Exit Sub
    End If
    
    Unload frmOutput
    
    nScriptType = lstFormat.ItemData(lstFormat.ListIndex)
    
    If (chkCreateScriptPerTable.Value = vbUnchecked) And chkCreateScriptPerTable.Enabled Then
        With cdlScript
            .CancelError = True
            .Flags = cdlOFNHideReadOnly
            
            If lstFormat.ListIndex = 0 Then
                .Filter = "SQL Script (.sql)|*.sql|All files (*.*)|*.*"
            Else
                .Filter = "VB Module (.bas)|*.sql|All files (*.*)|*.*"
            End If
            
            If Len(.InitDir) = 0 Then
                .InitDir = txtFolder.Text
            End If
            
            .ShowSave
                    
            nPos = InStrRev(.Filename, "\")
            If nPos > 0 Then
                txtFolder.Text = Left$(.Filename, nPos - 1)
                frmOutput.lvFiles.ListItems.Add , , Mid$(.Filename, nPos + 1)
            Else
                frmOutput.lvFiles.ListItems.Add , , .Filename
            End If
            
            nHandle = FreeFile
            Open .Filename For Output As #nHandle
            
            If (nScriptType = SCRIPT_VBCODE) Then
                WriteModuleHeader nHandle, GetFileFromPath(.Filename)
            End If
        End With
    End If
    
    If (chkCreateDatabaseModule.Value = vbChecked) And (nScriptType = SCRIPT_VBCODE) Then
        frmOutput.lvFiles.ListItems.Add 1, , "CreateDB.bas"
         
        nH = FreeFile
        Open txtFolder.Text & "\CreateDB.bas" For Output As #nH
        WriteModuleHeader nH, "CreateDB"
        Print #nH, "Public Sub CreateDatabase(sDatabase As String)"
        
        If optADO.Value Then
             Print #nH, Space$(4) & "Dim dbCat As ADOX.Catalog"
             Print #nH, ""
             Print #nH, Space$(4) & "Set dbCat = New ADOX.Catalog"
             Print #nH, Space$(4) & "dbCat.Create " & Chr$(34) & "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr$(34) & " & sDatabase"
             Print #nH, ""
         Else
             Print #nH, Space$(4) & "Dim db As DAO.Database"
             Print #nH, ""
             Print #nH, Space$(4) & "Set db = DBEngine.CreateDatabase(sDatabase, dbLangGeneral, dbVersion40)"
             Print #nH, ""
        End If
         
        Print #nH, Space$(4) & "'Code to Generate Table Objects"
        With lstTables
            For nIndex = 0 To .ListCount - 1
                If .Selected(nIndex) Then
                    If optADO.Value Then
                        Print #nH, Space$(4) & "Create_" & Replace(.List(nIndex), " ", "_") & " dbCat"
                    Else
                        Print #nH, Space$(4) & "Create_" & Replace(.List(nIndex), " ", "_") & " db"
                    End If
                End If
            Next nIndex
        End With
        
        Print #nH, ""
        If optADO.Value Then
            Print #nH, Space$(4) & "Set dbCat = Nothing"
        Else
            Print #nH, Space$(4) & "Set db = Nothing"
        End If
         
        Print #nH, "End Sub"
        Close #nH
    End If
    
    If (chkCreateProjectFile.Value = vbChecked) And (nScriptType = SCRIPT_VBCODE Or nScriptType = SCRIPT_VBCLASS_MODULE) Then
        If nScriptType = SCRIPT_VBCODE Then
            sTemp = "CreateDB"
        Else
            sTemp = "DBClasses"
        End If
        
        frmOutput.lvFiles.ListItems.Add 1, , sTemp & ".vbp"
        
        nH = FreeFile
        Open txtFolder.Text & "\" & sTemp & ".vbp" For Output As #nH
        Print #nH, "Type=Exe"
        Print #nH, "Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\..\..\..\WINDOWS\system32\stdole2.tlb#OLE Automation"
        If optADO.Value Then
            Print #nH, txtADOReference.Text
            Print #nH, txtADOXReference
        Else
            Print #nH, txtDAOReference
        End If
        
        If (chkCreateDatabaseModule.Value = vbChecked) And (nScriptType = SCRIPT_VBCODE) Then
            Print #nH, "Module=" & PREFIX_MODULE & "CreateDB;CreateDB.bas"
        End If
        
        With lstTables
            For nIndex = 0 To .ListCount - 1
                If .Selected(nIndex) Then
                    sTemp = Replace(.List(nIndex), " ", "_")
                    If nScriptType = SCRIPT_VBCODE Then
                        Print #nH, "Module=" & PREFIX_MODULE & sTemp & ";" & sTemp & ".bas"
                    Else
                        Print #nH, "Class=" & PREFIX_CLASS & sTemp & ";" & sTemp & ".cls"
                    End If
                End If
            Next nIndex
        End With
        
        Close #nH
    End If
    
    With lstTables
        For nIndex = 0 To .ListCount - 1
            If .Selected(nIndex) Then
                If optADO.Value Then
                    If cat Is Nothing Then
                        Set cat = New ADOX.Catalog
                        cat.ActiveConnection = mCN
                    End If
                    
                    Select Case lstFormat.ListIndex
                    Case 0
                        ADO_ScriptQA cat.Tables(.List(nIndex)), nHandle, chkCreateScriptPerTable.Value
                    Case 1
                        ADO_ScriptVBCode cat.Tables(.List(nIndex)), nHandle, chkCreateScriptPerTable.Value
                    Case 2
                        ADO_ScriptVBClass cat.Tables(.List(nIndex))
                    Case 3
                    End Select
                Else
                    Select Case lstFormat.ListIndex
                    Case 0
                        DAO_ScriptQA mDB.TableDefs(.List(nIndex)), nHandle, chkCreateScriptPerTable.Value
                    Case 1
                        DAO_ScriptVBCode mDB.TableDefs(.List(nIndex)), nHandle, chkCreateScriptPerTable.Value, False
                    Case 2
                        DAO_ScriptVBCode mDB.TableDefs(.List(nIndex)), nHandle, chkCreateScriptPerTable.Value, True
                    Case 3
                        DAO_ScriptVBClass mDB.TableDefs(.List(nIndex))
                    End Select
                End If
            End If
        Next nIndex
    End With
    
    If nHandle > 0 Then
        Close #nHandle
    End If
    
    If lstFormat.ListIndex = 0 Then
        frmOutput.ShowOutput emDefault
    Else
        frmOutput.ShowOutput emVB
    End If
End Sub

Private Sub cmdInvert_Click()
    Dim nIndex As Integer
    
    With lstTables
        For nIndex = 0 To .ListCount - 1
            .Selected(nIndex) = Not .Selected(nIndex)
        Next nIndex
    End With
End Sub

Private Sub cmdSave_Click()
    SaveSetting App.EXEName, "Options", "ADOReference", txtADOReference.Text
    SaveSetting App.EXEName, "Options", "ADOXReference", txtADOXReference.Text
    SaveSetting App.EXEName, "Options", "DAOReference", txtDAOReference.Text
    
    SaveSetting App.EXEName, "Options", "SQLCollate", txtSQLCollate.Text
End Sub

Private Sub cmdTagAll_Click()
    Dim nIndex As Integer
    
    With lstTables
        For nIndex = 0 To .ListCount - 1
            .Selected(nIndex) = True
        Next nIndex
    End With
End Sub

Private Function DAO_GetConnection(sConnection As String, Optional sPassword As String, Optional bExclusive As Boolean) As DAO.Database
    '////////////////////////////////////////////////////////////////////////////////////
    'Return an DAO connection to a database. Prompts for password if necessary
    '////////////////////////////////////////////////////////////////////////////////////
    
    Dim db As DAO.Database
    Dim lErr As Long
    Dim sPW As String
    
    On Local Error GoTo DAO_GetConnection
    
    If Len(sConnection) = 0 Then
        Set db = DBEngine.OpenDatabase(sConnection, False, False)
    Else
        If InStr(sConnection, ";") > 0 Then
            Set db = DBEngine.OpenDatabase("", dbDriverNoPrompt, False, sConnection)
        Else
            On Local Error GoTo DAO_GetConnectionError
            
            sPW = sPassword
            Do
                lErr = 0
                If Len(sPW) > 0 Then
                    Set db = DBEngine.OpenDatabase(sConnection, bExclusive, False, "MS Access;PWD=" & sPW)
                Else
                    Set db = DBEngine.OpenDatabase(sConnection, bExclusive, False)
                End If
                
                If lErr = 0 Then
                    Exit Do
                End If
            Loop
        End If
    End If
    
    Set DAO_GetConnection = db
    Exit Function

DAO_GetConnection:
    Resume Next

DAO_GetConnectionError:
    lErr = Err.Number
    
    Select Case lErr
    Case 3024
        MsgBox "The database does not exist. It may have been moved or renamed.", vbCritical
    
    Case 3031
        With frmPW
            .txtPassword.Text = ""
            .Show vbModal
            
            If Val(.Tag) = 1 Then
                sPW = .txtPassword.Text
                Resume Next
            Else
                Exit Function
            End If
        End With
    
    Case Else
        MsgBox Err.Description, vbCritical
        
    End Select
End Function

Private Sub DAO_LoadObjects()
    Dim tbl As DAO.TableDef
    
    On Local Error GoTo DAO_LoadObjectsError
    
    Screen.MousePointer = vbHourglass
    
    With lstTables
        .Clear
      
        For Each tbl In mDB.TableDefs
            If (tbl.Attributes And dbSystemObject) = 0 Or chkViewSystemTables.Value Then
                .AddItem TblName(tbl.Name)
            End If
        Next tbl
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
DAO_LoadObjectsError:
    MsgBox Err.Description, vbCritical
End Sub

Private Function DAO_MapDataType(DataType As Long, ConvertType As ConvertTypeEnum) As String
    '////////////////////////////////////////////////////////////////////////////////////
    'Maps a DAO DataType to an SQL Server Query Analyser or VB one
    '////////////////////////////////////////////////////////////////////////////////////
    
    If ConvertType = ctSQLServer Then
        Select Case DataType
        Case dbBigInt
            DAO_MapDataType = "bigint"
        Case dbBinary
            DAO_MapDataType = "binary"
        Case dbBoolean
            DAO_MapDataType = "bit"
        Case dbByte
            DAO_MapDataType = "byte"
        Case dbChar
            DAO_MapDataType = "char"
        Case dbCurrency
            DAO_MapDataType = "money"
        Case dbDate
            DAO_MapDataType = "smalldatetime"
        Case dbDecimal
            DAO_MapDataType = "decimal"
        Case dbDouble
            DAO_MapDataType = "float"
        Case dbFloat
            DAO_MapDataType = "float"
        Case dbInteger
            DAO_MapDataType = "smallint"
        Case dbLong
            DAO_MapDataType = "int"
        Case dbLongBinary
            DAO_MapDataType = "binary"
        Case dbMemo
            DAO_MapDataType = "ntext"
        Case dbNumeric
            DAO_MapDataType = "numeric"
        Case dbSingle
            DAO_MapDataType = "real"
        Case dbText
            DAO_MapDataType = "nvarchar"
        Case dbTime
            DAO_MapDataType = "datetime"
        Case dbTimeStamp
            DAO_MapDataType = "timestamp"
        Case dbVarBinary
            DAO_MapDataType = "varbinary"
        End Select
    Else
        Select Case DataType
        Case dbBoolean
            DAO_MapDataType = "Boolean"
        Case dbByte
            DAO_MapDataType = "Byte"
        Case dbCurrency
            DAO_MapDataType = "Currency"
        Case dbDate, dbTime
            DAO_MapDataType = "Date"
        Case dbDouble
            DAO_MapDataType = "Double"
        Case dbInteger
            DAO_MapDataType = "Integer"
        Case dbLong
            DAO_MapDataType = "Long"
        Case dbMemo, dbText
            DAO_MapDataType = "String"
        Case dbSingle
            DAO_MapDataType = "Single"
        Case Else
            DAO_MapDataType = "Variant"
        End Select
    End If
End Function

Private Sub DAO_ScriptQA(tbl As DAO.TableDef, nHandle As Integer, bScriptPerTable As Boolean)
    '////////////////////////////////////////////////////////////////////////////////////
    'Create SQL script for Query Analyser to generate the source DAO Table
    '////////////////////////////////////////////////////////////////////////////////////
    
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim bMemo As Boolean
    Dim bPrimary As Boolean
    Dim sBuffer As String
    Dim sText As String
    Dim sTab As String

    On Local Error GoTo ScriptQAError
    sTab = Space$(4)

    If bScriptPerTable Then
        nHandle = FreeFile
        Open txtFolder.Text & "\" & TblName(tbl.Name) & ".sql" For Output As #nHandle
        
        frmOutput.lvFiles.ListItems.Add , , TblName(tbl.Name) & ".sql"
    End If

    Print #nHandle, "-- " & String(80, "#")
                
    If chkDropTable.Value Then
        Print #nHandle, "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[" & TblName(tbl.Name) & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
        Print #nHandle, "drop table [dbo].[" & TblName(tbl.Name) & "]"
        Print #nHandle, "GO"
        Print #nHandle, ""
    End If
    
    If chkTables.Value Then
        Print #nHandle, "CREATE TABLE [dbo].[" & TblName(tbl.Name) & "] ("
        
        bMemo = False
        sBuffer = ""
        
        For Each fld In tbl.Fields
            If (fld.Attributes And dbAutoIncrField) = dbAutoIncrField Then
                sText = "[" & fld.Name & "] [int] IDENTITY (1,1) NOT NULL"
            Else
                Select Case fld.Type
                Case dbText
                    sText = "[" & fld.Name & "] [nvarchar] (" & fld.Size & ") COLLATE " & txtSQLCollate.Text
                Case dbMemo
                    bMemo = True
                    sText = "[" & fld.Name & "] [ntext] COLLATE " & txtSQLCollate.Text
                Case Else
                    sText = "[" & fld.Name & "] [" & DAO_MapDataType(fld.Type, ctSQLServer) & "]"
                End Select
                
                If fld.Required Then
                    sText = sText & " NOT NULL"
                Else
                    sText = sText & " NULL"
                End If
            End If
            
            If Len(sBuffer) = 0 Then
                sBuffer = sTab & sText
            Else
                sBuffer = sBuffer & " ," & vbCrLf & sTab & sText
            End If
        Next fld
        
        Print #nHandle, sBuffer
        
        If bMemo Then
            Print #nHandle, ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
        Else
            Print #nHandle, ") ON [PRIMARY]"
        End If
        Print #nHandle, "GO"
        Print #nHandle, ""
    End If
    
    If chkIndexes.Value Then
        For Each idx In tbl.Indexes
            If IsJetDB() Then
                bPrimary = idx.Primary
            End If
            
            If bPrimary Then
                Print #nHandle, "ALTER TABLE [dbo].[" & TblName(tbl.Name) & "] WITH NOCHECK ADD"
                Print #nHandle, "CONSTRAINT [PK_" & TblName(tbl.Name) & "] PRIMARY KEY  CLUSTERED"
                Print #nHandle, "("
                
                sBuffer = ""
                For Each fld In idx.Fields
                    If Len(sBuffer) = 0 Then
                        sBuffer = "[" & fld.Name & "]"
                    Else
                        sBuffer = sBuffer & " ," & vbCrLf & "[" & fld.Name & "]"
                    End If
                Next fld
                
                Print #nHandle, sBuffer
                Print #nHandle, ") ON [PRIMARY]"
                Print #nHandle, "GO"
                Print #nHandle, ""
            Else
                Print #nHandle, "CREATE NONCLUSTERED INDEX [IK_" & idx.Name & "] ON [dbo].[" & TblName(tbl.Name) & "]"
                Print #nHandle, "("
                
                sBuffer = ""
                For Each fld In idx.Fields
                    If Len(sBuffer) = 0 Then
                        sBuffer = "[" & fld.Name & "]"
                    Else
                        sBuffer = sBuffer & " ," & vbCrLf & "[" & fld.Name & "]"
                    End If
                Next fld
                
                Print #nHandle, sBuffer
                Print #nHandle, ") ON [PRIMARY]"
                Print #nHandle, "GO"
                Print #nHandle, ""
            End If
        Next idx
    End If
    
    If chkFieldProperties.Value Then
        sBuffer = ""
    
        For Each fld In tbl.Fields
            If (fld.Attributes And dbAutoIncrField) = dbAutoIncrField Then
                'No default allowed
            Else
                Select Case fld.Type
                Case dbText
                    sText = "CONSTRAINT [DF_" & TblName(tbl.Name) & "_" & fld.Name & "] DEFAULT '' FOR [" & fld.Name & "]"
                Case dbInteger, dbLong, dbCurrency, dbSingle, dbDouble
                    If Val(fld.DefaultValue) <> 0 Then
                        sText = "CONSTRAINT [DF_" & TblName(tbl.Name) & "_" & fld.Name & "] DEFAULT (" & fld.DefaultValue & ") FOR [" & fld.Name & "]"
                    Else
                        sText = "CONSTRAINT [DF_" & TblName(tbl.Name) & "_" & fld.Name & "] DEFAULT (0) FOR [" & fld.Name & "]"
                    End If
                Case dbBoolean
                    If Val(fld.DefaultValue) <> 0 Then
                        sText = "CONSTRAINT [DF_" & TblName(tbl.Name) & "_" & fld.Name & "] DEFAULT (1) FOR [" & fld.Name & "]"
                    Else
                        sText = "CONSTRAINT [DF_" & TblName(tbl.Name) & "_" & fld.Name & "] DEFAULT (0) FOR [" & fld.Name & "]"
                    End If
                Case Else
                    sText = ""
                End Select
            
                If Len(sText) > 0 Then
                    If Len(sBuffer) = 0 Then
                        sBuffer = sTab & sText
                    Else
                        sBuffer = sBuffer & " ," & vbCrLf & sTab & sText
                    End If
                End If
            End If
        Next fld
    
        If Len(sBuffer) > 0 Then
            Print #nHandle, "ALTER TABLE [dbo].[" & TblName(tbl.Name) & "] WITH NOCHECK ADD"
            Print #nHandle, sBuffer
            Print #nHandle, "GO"
            Print #nHandle, ""
        End If
    End If
    
    If bScriptPerTable Then
        Close #nHandle
        nHandle = 0
    End If

    Exit Sub

ScriptQAError:
    MsgBox Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub DAO_ScriptVBClass(dbTbl As DAO.TableDef)
    '////////////////////////////////////////////////////////////////////////////////////
    'Create a VB6 Class Module from an DAO Table
    '////////////////////////////////////////////////////////////////////////////////////
    
    Dim dbFld As DAO.Field
    Dim nHandle As Integer
    Dim sComment As String
    Dim sField As String
    Dim sTab As String
    
    On Local Error GoTo ScriptVBClassError
    
    sTab = Space$(4)

    nHandle = FreeFile
    Open txtFolder.Text & "\" & TblName(dbTbl.Name) & ".cls" For Output As #nHandle
    
    frmOutput.lvFiles.ListItems.Add , , TblName(dbTbl.Name) & ".cls"
    
    WriteClassHeader nHandle, TblName(dbTbl.Name)
    Print #nHandle, "'Source Database: " & txtDatabase.Text
    Print #nHandle, "'Source Table: " & TblName(dbTbl.Name)
    Print #nHandle, ""
    
    For Each dbFld In dbTbl.Fields
        sField = FormatName(dbFld.Name)
        
        Select Case dbFld.Type
        Case dbBinary
            sComment = " '(Binary Field)"
        Case dbMemo
            sComment = " '(Memo Field)"
        Case Else
            sComment = ""
        End Select
        
        Print #nHandle, "Private m" & sField & " As " & DAO_MapDataType(dbFld.Type, ctVB) & sComment
    Next dbFld
    
    Print #nHandle, ""
    Print #nHandle, "Public Sub LoadData (db as DAO.Database)"
    Print #nHandle, Space$(4) & "Dim rs As DAO.RecordSet"
    Print #nHandle, Space$(4) & "Dim SQL As String"
    Print #nHandle, ""
    Print #nHandle, Space$(4) & "'\\ Edit SQL WHERE as required \\"
    Print #nHandle, Space$(4) & "SQL = " & Chr$(34) & "SELECT * FROM " & TblName(dbTbl.Name) & Chr$(34)
    Print #nHandle, Space$(4) & "Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot)"
    Print #nHandle, Space$(4) & "With rs"
    Print #nHandle, Space$(8) & "If .RecordCount > 0 Then"
    For Each dbFld In dbTbl.Fields
        sField = FormatName(dbFld.Name)
        Print #nHandle, Space$(12) & "m" & sField & " = !" & dbFld.Name
    Next dbFld
    
    Print #nHandle, Space$(8) & "End If"
    Print #nHandle, Space$(8) & ".Close"
    Print #nHandle, Space$(4) & "End With"
    Print #nHandle, Space$(4) & "Set rs = Nothing"
    Print #nHandle, "End Sub"
    Print #nHandle, ""
    
    Print #nHandle, "Public Sub SaveData (db as DAO.Database)"
    Print #nHandle, Space$(4) & "Dim rs As DAO.RecordSet"
    Print #nHandle, Space$(4) & "Dim SQL As String"
    Print #nHandle, ""
    Print #nHandle, Space$(4) & "'\\ Edit SQL WHERE as required \\"
    Print #nHandle, Space$(4) & "SQL = " & Chr$(34) & "SELECT * FROM " & TblName(dbTbl.Name) & Chr$(34)
    Print #nHandle, Space$(4) & "Set rs = db.OpenRecordSet(SQL, dbOpenDynaset, dbSeeChanges)"
    Print #nHandle, Space$(4) & "With rs"
    Print #nHandle, Space$(8) & "If .RecordCount = 0 Then"
    Print #nHandle, Space$(12) & ".AddNew"
    Print #nHandle, Space$(8) & "Else"
    Print #nHandle, Space$(12) & ".Edit"
    Print #nHandle, Space$(8) & "End If"
    
    For Each dbFld In dbTbl.Fields
        If (dbFld.Attributes And dbAutoIncrField) <> dbAutoIncrField Then
            sField = FormatName(dbFld.Name)
            If dbFld.Type = dbDate Then
                Print #nHandle, Space$(8) & "If IsDate(m" & sField & ") Then"
                Print #nHandle, Space$(12) & "!" & dbFld.Name & " = m" & sField
                Print #nHandle, Space$(8) & "Else"
                Print #nHandle, Space$(12) & "!" & dbFld.Name & " = Empty"
                Print #nHandle, Space$(8) & "End If"
            Else
                Print #nHandle, Space$(8) & "!" & dbFld.Name & " = m" & sField
            End If
        End If
    Next dbFld
    
    Print #nHandle, Space$(8) & ".Update"
    Print #nHandle, Space$(8) & ".Close"
    Print #nHandle, Space$(4) & "End With"
    Print #nHandle, Space$(4) & "Set rs = Nothing"
    Print #nHandle, "End Sub"
    Print #nHandle, ""
    
    For Each dbFld In dbTbl.Fields
        sField = FormatName(dbFld.Name)
        Print #nHandle, "Public Property Get " & sField & " () As " & DAO_MapDataType(dbFld.Type, ctVB)
        Print #nHandle, sTab & sField & " = m" & sField
        Print #nHandle, "End Property"
        Print #nHandle, ""
        Print #nHandle, "Public Property Let " & sField & " (NewValue As " & DAO_MapDataType(dbFld.Type, ctVB) & ")"
        Print #nHandle, sTab & "m" & sField & " = NewValue"
        Print #nHandle, "End Property"
        Print #nHandle, ""
    Next dbFld
    
    Close #nHandle
    Exit Sub

ScriptVBClassError:
    MsgBox Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub DAO_ScriptVBCode(dbTbl As DAO.TableDef, nHandle As Integer, bScriptPerTable As Boolean, bSQL As Boolean)
    '////////////////////////////////////////////////////////////////////////////////////
    'Create VB6 code to generate the source  DAO Table
    '////////////////////////////////////////////////////////////////////////////////////
    
    Dim dbFld As DAO.Field
    Dim fldLU As DAO.Field
    Dim dbIdx As DAO.Index
    Dim sBuffer As String
    Dim sText As String
    Dim sTab As String
    
    On Local Error GoTo DAO_ScriptVBCodeError
    
    sTab = Space$(4)
    
    If bScriptPerTable Then
        nHandle = FreeFile
        Open txtFolder.Text & "\" & dbTbl.Name & ".bas" For Output As #nHandle
        WriteModuleHeader nHandle, dbTbl.Name
        
        frmOutput.lvFiles.ListItems.Add , , dbTbl.Name & ".bas"
    End If
        
    Print #nHandle, "Public Sub Create_" & Replace(dbTbl.Name, " ", "_") & "(DB As DAO.Database)"
    
    If bSQL Then
        Print #nHandle, sTab & "Dim SQL As String"
    Else
        Print #nHandle, sTab & "Dim dbFld As DAO.Field"
        Print #nHandle, sTab & "Dim dbIdx As DAO.Index"
        Print #nHandle, sTab & "Dim dbTbl As DAO.TableDef"
    End If
    
    Print #nHandle, ""
    Print #nHandle, sTab & "'Source Database: " & txtDatabase.Text
    Print #nHandle, ""

    Print #nHandle, sTab & "'" & String(120, "#")
    Print #nHandle, sTab & "'Code to generate Objects for Table: " & dbTbl.Name

    If bSQL Then
        If chkTables.Value Then
            Print #nHandle, sTab & "SQL = " & Chr$(34) & "CREATE TABLE " & dbTbl.Name & " (" & Chr$(34)
    
            sBuffer = ""
            For Each dbFld In dbTbl.Fields
                If (dbFld.Type = dbText) Then
                    sText = sTab & "SQL = SQL & " & Chr$(34) & dbFld.Name & " TEXT (" & dbFld.Size & ")"
                Else
                    sText = sTab & "SQL = SQL & " & Chr$(34) & dbFld.Name & " " & GetJetDDLType(dbFld.Type)
                End If
    
                If Len(sBuffer) > 0 Then
                    sBuffer = sBuffer & "," & Chr$(34) & vbCrLf & sText
                Else
                    sBuffer = sText
                End If
            Next dbFld
    
            Print #nHandle, sBuffer & ")"
            Print #nHandle, sTab & "DB.Execute SQL"
            Print #nHandle, ""
        End If

        If chkIndexes.Value Then
            For Each dbIdx In dbTbl.Indexes
                sBuffer = ""
                For Each dbFld In dbIdx.Fields
                    If Len(sBuffer) > 0 Then
                        sBuffer = sBuffer & "," & dbFld.Name
                    Else
                        sBuffer = dbFld.Name
                    End If
                Next dbFld

                If dbIdx.Primary Then
                    Print #nHandle, sTab & "DB.Execute " & Chr$(34) & "CREATE UNIQUE INDEX " & dbIdx.Name & " ON " & dbTbl.Name & "(" & sBuffer & ") WITH PRIMARY"
                ElseIf dbIdx.Unique Then
                    Print #nHandle, sTab & "DB.Execute " & Chr$(34) & "CREATE UNIQUE INDEX " & dbIdx.Name & " ON " & dbTbl.Name & "(" & sBuffer & ")"
                Else
                    Print #nHandle, sTab & "DB.Execute " & Chr$(34) & "CREATE INDEX " & dbIdx.Name & " ON " & dbTbl.Name & "(" & sBuffer & ")"
                End If
            Next dbIdx
        End If
    Else
        If chkTables.Value Then
            Print #nHandle, sTab & "Set dbTbl = DB.CreateTableDef(" & Chr$(34) & dbTbl.Name & Chr$(34) & ")"
            Print #nHandle, sTab & "With dbTbl"
            
            For Each dbFld In dbTbl.Fields
                If (dbFld.Type = dbText) Or (dbFld.Type = dbMemo) Then
                    Print #nHandle, sTab & sTab & ".Fields.Append .CreateField(" & Chr$(34) & dbFld.Name & Chr$(34) & "," & GetDAOConstant(dbFld.Type) & "," & dbFld.Size & ")"
                Else
                    Print #nHandle, sTab & sTab & ".Fields.Append .CreateField(" & Chr$(34) & dbFld.Name & Chr$(34) & "," & GetDAOConstant(dbFld.Type) & ")"
                End If
    
                If chkFieldAttributes.Value And dbFld.Attributes <> 0 Then
                    Print #nHandle, sTab & sTab & ".Fields(" & Chr$(34) & dbFld.Name & Chr$(34) & ").Attributes = " & dbFld.Attributes
                End If
                
                If chkFieldProperties.Value Then
                    If dbFld.AllowZeroLength Then
                        Print #nHandle, sTab & sTab & ".Fields(" & Chr$(34) & dbFld.Name & Chr$(34) & ").AllowZeroLength = True"
                    End If
                    If Len(dbFld.DefaultValue) > 0 Then
                        If dbFld.DefaultValue = Chr$(34) & Chr$(34) Then
                            Print #nHandle, sTab & sTab & ".Fields(" & Chr$(34) & dbFld.Name & Chr$(34) & ").DefaultValue = Chr$(34) & Chr$(34)"
                        Else
                            Print #nHandle, sTab & sTab & ".Fields(" & Chr$(34) & dbFld.Name & Chr$(34) & ").DefaultValue = " & dbFld.DefaultValue
                        End If
                    End If
                    If dbFld.Required Then
                        Print #nHandle, sTab & sTab & ".Fields(" & Chr$(34) & dbFld.Name & Chr$(34) & ").Required = True"
                    End If
                End If
            Next dbFld
            
            Print #nHandle, sTab & "End With"
            Print #nHandle, ""
            Print #nHandle, sTab & "DB.TableDefs.Append dbTbl"
        Else
            Print #nHandle, sTab & "Set dbTbl = DB.TableDefs(" & Chr$(34) & dbTbl.Name & Chr$(34) & ")"
        End If

        If chkIndexes.Value Then
            For Each dbIdx In dbTbl.Indexes
                Print #nHandle, ""
                Print #nHandle, sTab & "Set dbIdx = dbTbl.CreateIndex(" & Chr$(34) & dbIdx.Name & Chr$(34) & ")"
                Print #nHandle, sTab & "With dbIdx"
                If dbIdx.Primary Then
                    Print #nHandle, sTab & sTab & ".Primary = True"
                End If
                If dbIdx.Unique Then
                    Print #nHandle, sTab & sTab & ".Unique = True"
                End If

                For Each dbFld In dbIdx.Fields
                    Set fldLU = dbTbl.Fields(dbFld.Name)

                    If (fldLU.Type = dbText) Or (fldLU.Type = dbMemo) Then
                        Print #nHandle, sTab & sTab & ".Fields.Append .CreateField(" & Chr$(34) & dbFld.Name & Chr$(34) & "," & GetDAOConstant(fldLU.Type) & "," & fldLU.Size & ")"
                    Else
                        Print #nHandle, sTab & sTab & ".Fields.Append .CreateField(" & Chr$(34) & dbFld.Name & Chr$(34) & "," & GetDAOConstant(fldLU.Type) & ")"
                    End If
                Next dbFld

                Print #nHandle, sTab & "End With"
                Print #nHandle, sTab & "dbTbl.Indexes.Append dbIdx"
            Next dbIdx
        End If
    End If

    Print #nHandle, "End Sub"
    
    If bScriptPerTable Then
        Close #nHandle
        nHandle = 0
    End If
    Exit Sub

DAO_ScriptVBCodeError:
    MsgBox Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
    Dim nValue As Integer
    
    Me.Caption = Me.Caption & " (" & App.Major & "." & App.Minor & "." & App.Revision & ")"
    
    If GetSetting(App.EXEName, "Database", "ConnectionMethod", "ADO") = "ADO" Then
        optADO.Value = True
    Else
        optDAO.Value = True
    End If
    
    LoadScriptTypes
         
    '#######################################################################################
    'Script Tab
    chkViewSystemTables.Value = Val(GetSetting(App.EXEName, "Objects", "ViewSystemTables", vbUnchecked))
    chkTables.Value = Val(GetSetting(App.EXEName, "Objects", "Tables", vbChecked))
    chkIndexes.Value = Val(GetSetting(App.EXEName, "Objects", "Indexes", vbChecked))
    chkFieldAttributes.Value = Val(GetSetting(App.EXEName, "Objects", "FieldAttributes", vbChecked))
    chkFieldProperties.Value = Val(GetSetting(App.EXEName, "Objects", "FieldProperties", vbUnchecked))
    
    nValue = Val(GetSetting(App.EXEName, "Script", "Format", 0))
    lstFormat.ListIndex = nValue
    
    chkCreateScriptPerTable.Value = Val(GetSetting(App.EXEName, "Script", "CreateScriptPerTable", vbChecked))
    chkCreateDatabaseModule.Value = Val(GetSetting(App.EXEName, "Script", "CreateDatabaseModule", vbChecked))
    chkCreateProjectFile.Value = Val(GetSetting(App.EXEName, "Script", "CreateProjectFile", vbUnchecked))
    
    txtFolder.Text = GetSetting(App.EXEName, "Script", "Folder", "")
    
    '#######################################################################################
    'Options Tab
    txtADOReference.Text = GetSetting(App.EXEName, "Options", "ADOReference", DEF_ADOREF)
    txtADOXReference.Text = GetSetting(App.EXEName, "Options", "ADOXReference", DEF_ADOXREF)
    txtDAOReference.Text = GetSetting(App.EXEName, "Options", "DAOReference", DEF_DAOREF)
    
    txtSQLCollate.Text = GetSetting(App.EXEName, "Options", "SQLCollate", DEF_COLLATE)
    
    SetState
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.EXEName, "Database", "ConnectionMethod", IIf(optADO.Value, "ADO", "DAO")
    
    SaveSetting App.EXEName, "Objects", "ViewSystemTables", chkViewSystemTables.Value
    SaveSetting App.EXEName, "Objects", "Tables", chkTables.Value
    SaveSetting App.EXEName, "Objects", "Indexes", chkIndexes.Value
    SaveSetting App.EXEName, "Objects", "FieldAttributes", chkFieldAttributes.Value
    SaveSetting App.EXEName, "Objects", "FieldProperties", chkFieldProperties.Value
    
    SaveSetting App.EXEName, "Script", "Format", lstFormat.ListIndex
    SaveSetting App.EXEName, "Script", "CreateScriptPerTable", chkCreateScriptPerTable.Value
    SaveSetting App.EXEName, "Script", "CreateDatabaseModule", chkCreateDatabaseModule.Value
    SaveSetting App.EXEName, "Script", "CreateProjectFile", chkCreateProjectFile.Value
    SaveSetting App.EXEName, "Script", "Folder", txtFolder.Text
End Sub

Private Function FormatName(ByVal sName As String) As String
    Mid$(sName, 1, 1) = UCase$(Mid$(sName, 1, 1))
    If InStr(sName, " ") > 0 Then
        sName = Replace(sName, " ", "")
    End If
    
    FormatName = sName
End Function

Private Function GetADOConstant(lValue As Long) As String
    '////////////////////////////////////////////////////////////////////////////////////
    'Return the "text" version of an ADO DataType
    '////////////////////////////////////////////////////////////////////////////////////

    Select Case lValue
    Case adTinyInt
        GetADOConstant = "adTinyInt"
    Case adSmallInt
        GetADOConstant = "adSmallInt"
    Case adInteger
        GetADOConstant = "adInteger"
    Case adBigInt
        GetADOConstant = "adBigInt"
    Case adUnsignedTinyInt
        GetADOConstant = "adUnsignedTinyInt"
    Case adUnsignedSmallInt
        GetADOConstant = "adUnsignedSmallInt"
    Case adUnsignedInt
        GetADOConstant = "adUnsignedInt"
    Case adUnsignedBigInt
        GetADOConstant = "adUnsignedBigInt"
    Case adSingle
        GetADOConstant = "adSingle"
    Case adDouble
        GetADOConstant = "adDouble"
    Case adCurrency
        GetADOConstant = "adCurrency"
    Case adDecimal
        GetADOConstant = "adDecimal"
    Case adNumeric
        GetADOConstant = "adNumeric"
    Case adBoolean
        GetADOConstant = "adBoolean"
    Case adUserDefined
        GetADOConstant = "adUserDefined"
    Case adVariant
        GetADOConstant = "adVariant"
    Case adGUID
        GetADOConstant = "adGuid"
    Case adDate
        GetADOConstant = "adDate"
    Case adDBDate
        GetADOConstant = "adDate"
    Case adDBTime
        GetADOConstant = "adDBTime"
    Case adDBTimeStamp
        GetADOConstant = "adDBTimestamp"
    Case adBSTR
        GetADOConstant = "adBSTR"
    Case adChar
        GetADOConstant = "adChar"
    Case adVarChar
        GetADOConstant = "adVarChar"
    Case adLongVarChar
        GetADOConstant = "adLongVarChar"
    Case adWChar
        GetADOConstant = "adWChar"
    Case adVarWChar
        GetADOConstant = "adVarWChar"
    Case adLongVarWChar
        GetADOConstant = "adLongVarWChar"
    Case adBinary
        GetADOConstant = "adBinary"
    Case adVarBinary
        GetADOConstant = "adVarBinary"
    Case adLongVarBinary
        GetADOConstant = "adLongVarBinary"
    Case Else
        GetADOConstant = CStr(lValue)
    End Select
End Function

Private Function GetDAOConstant(lValue As Long) As String
    '////////////////////////////////////////////////////////////////////////////////////
    'Return the "text" version of an DAO DataType
    '////////////////////////////////////////////////////////////////////////////////////
    
    Select Case lValue
    Case dbBigInt
        GetDAOConstant = "dbBigInt"
    Case dbBinary
        GetDAOConstant = "dbBinary"
    Case dbBoolean
        GetDAOConstant = "dbBoolean"
    Case dbByte
        GetDAOConstant = "dbByte"
    Case dbChar
        GetDAOConstant = "dbByte"
    Case dbCurrency
        GetDAOConstant = "dbCurrency"
    Case dbDate
        GetDAOConstant = "dbDate"
    Case dbDecimal
        GetDAOConstant = "dbDecimal"
    Case dbDouble
        GetDAOConstant = "dbDouble"
    Case dbFloat
        GetDAOConstant = "dbFloat"
    Case dbGUID
        GetDAOConstant = "dbGUID"
    Case dbInteger
        GetDAOConstant = "dbInteger"
    Case dbLong
        GetDAOConstant = "dbLong"
    Case dbLongBinary
        GetDAOConstant = "dbLongBinary"
    Case dbMemo
        GetDAOConstant = "dbMemo"
    Case dbNumeric
        GetDAOConstant = "dbNumeric"
    Case dbSingle
        GetDAOConstant = "dbSingle"
    Case dbText
        GetDAOConstant = "dbText"
    Case dbTime
        GetDAOConstant = "dbTime"
    Case dbTimeStamp
        GetDAOConstant = "dbTimeStamp"
    Case dbVarBinary
        GetDAOConstant = "dbVarBinary"
    Case Else
        GetDAOConstant = CStr(lValue)
    End Select
End Function

Private Function GetJetDDLType(DataType As Long) As String
    '////////////////////////////////////////////////////////////////////////////////////
    'Returns the DAO DataType used when creating Database Objects with SQL statements
    '////////////////////////////////////////////////////////////////////////////////////

    Select Case DataType
    Case dbBinary
        GetJetDDLType = "BINARY"
    Case dbBoolean
        GetJetDDLType = "BIT"
    Case dbByte
        GetJetDDLType = "BYTE"
    Case dbCurrency
        GetJetDDLType = "CURRENCY"
    Case dbDate
        GetJetDDLType = "DATETIME"
    Case dbDouble
        GetJetDDLType = "DOUBLE"
    Case dbInteger
        GetJetDDLType = "SHORT"
    Case dbLong
        GetJetDDLType = "LONG"
    Case dbMemo
        GetJetDDLType = "LONGTEXT"
    Case dbSingle
        GetJetDDLType = "SINGLE"
    Case dbText
        GetJetDDLType = "TEXT"
    Case dbTime
        GetJetDDLType = "DATETIME"
    End Select
End Function

Private Function IsJetDB() As Boolean
    If optADO.Value Then
    Else
        IsJetDB = (InStr(LCase$(mDB.Connect), "jet") > 0)
    End If
End Function

Private Sub LoadScriptTypes()
    With lstFormat
        .Clear
        .AddItem "SQL Server - Query Analyser"
        .ItemData(.NewIndex) = SCRIPT_QUERY_ANALYSER
        
        If optADO.Value Then
            .AddItem "ADO Objects (VB6 Code)"
            .ItemData(.NewIndex) = SCRIPT_VBCODE
        Else
            .AddItem "DAO Objects (VB6 Code)"
            .ItemData(.NewIndex) = SCRIPT_VBCODE
            .AddItem "DAO Execute (VB6 Code)"
            .ItemData(.NewIndex) = SCRIPT_VBCODE
        End If
        
        .AddItem "Class Module (VB6 Code)"
        .ItemData(.NewIndex) = SCRIPT_VBCLASS_MODULE
        
        .ListIndex = 0
    End With
End Sub

Private Sub lstFormat_Click()
    With lstFormat
        chkDropTable.Enabled = (.ItemData(.ListIndex) = SCRIPT_QUERY_ANALYSER)
        chkCreateScriptPerTable.Enabled = (.ItemData(.ListIndex) <> SCRIPT_VBCLASS_MODULE)
        chkCreateDatabaseModule.Enabled = (.ItemData(.ListIndex) <> SCRIPT_QUERY_ANALYSER)
        chkCreateProjectFile.Enabled = (.ItemData(.ListIndex) <> SCRIPT_QUERY_ANALYSER)
    End With
End Sub

Private Sub optADO_Click()
    LoadScriptTypes
End Sub

Private Sub optDAO_Click()
    LoadScriptTypes
End Sub

Private Sub SetState()
    '////////////////////////////////////////////////////////////////////////////////////
    'Set Enabled State of controls depending on Connection State
    '////////////////////////////////////////////////////////////////////////////////////
    
    Dim bActive As Boolean
    
    If optDAO.Value Then
        bActive = Not (mDB Is Nothing)
    Else
        bActive = Not (mCN Is Nothing)
    End If
    
    cmdTagAll.Enabled = bActive
    cmdInvert.Enabled = bActive
    cmdGenerate.Enabled = bActive
    
    If Not bActive Then
        cmdConnect.Caption = "Connect"
        optDAO.Enabled = True
        optADO.Enabled = True
        lstTables.Clear
    Else
        cmdConnect.Caption = "Disconnect"
        optDAO.Enabled = False
        optADO.Enabled = False
    End If
End Sub

Private Function TblName(sTablename As String) As String
    Dim nPos As Integer
    
    nPos = InStr(sTablename, ".")
    If nPos > 0 Then
        TblName = Mid$(sTablename, nPos + 1)
    Else
        TblName = sTablename
    End If
End Function

Private Sub WriteClassHeader(nHandle As Integer, sClassName As String)
    Print #nHandle, "VERSION 1.0 CLASS"
    Print #nHandle, "BEGIN"
    Print #nHandle, "  MultiUse = -1  'True"
    Print #nHandle, "  Persistable = 0  'NotPersistable"
    Print #nHandle, "  DataBindingBehavior = 0  'vbNone"
    Print #nHandle, "  DataSourceBehavior = 0   'vbNone"
    Print #nHandle, "  MTSTransactionMode = 0   'NotAnMTSObject"
    Print #nHandle, "End"
    Print #nHandle, "Attribute VB_Name = " & Chr$(34) & PREFIX_CLASS & sClassName & Chr$(34)
    Print #nHandle, "Attribute VB_GlobalNameSpace = False"
    Print #nHandle, "Attribute VB_Creatable = True"
    Print #nHandle, "Attribute VB_PredeclaredId = False"
    Print #nHandle, "Attribute VB_Exposed = False"
    Print #nHandle, "Option Explicit"
    Print #nHandle, ""
End Sub

Private Sub WriteModuleHeader(nHandle As Integer, sModuleName As String)
    Print #nHandle, "Attribute VB_Name = " & Chr$(34) & PREFIX_MODULE & sModuleName & Chr$(34)
    Print #nHandle, "Option Explicit"
    Print #nHandle, ""
End Sub
