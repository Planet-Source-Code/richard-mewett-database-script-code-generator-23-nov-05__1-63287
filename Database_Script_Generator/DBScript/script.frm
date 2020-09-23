VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmScript 
   Caption         =   "Database Script Generator"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   825
   ClientWidth     =   9510
   Icon            =   "script.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   9510
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picScripting 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2798
      ScaleHeight     =   945
      ScaleWidth      =   3885
      TabIndex        =   16
      Top             =   2925
      Visible         =   0   'False
      Width           =   3915
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   30
         TabIndex        =   18
         Top             =   570
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblObject 
         Height          =   225
         Left            =   60
         TabIndex        =   17
         Top             =   90
         Width           =   3765
      End
   End
   Begin VB.CheckBox chkHighlight 
      Caption         =   "Highlight Keywords"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   6510
      Width           =   1995
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1515
      Left            =   540
      TabIndex        =   0
      Top             =   2340
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   2672
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&General"
      TabPicture(0)   =   "script.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblOutput"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdOutput"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtFolder"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkFieldProperties"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkIndexes"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkTables"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CommonDialog1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboFormat"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkFieldAttributes"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cboOutput"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "&Query Analyser"
      TabPicture(1)   =   "script.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkDropTable"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Visual Basic"
      TabPicture(2)   =   "script.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkSubPerTable"
      Tab(2).Control(1)=   "chkCreateDatabaseModule"
      Tab(2).ControlCount=   2
      Begin VB.CheckBox chkCreateDatabaseModule 
         Caption         =   "Create Database Module"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   2385
      End
      Begin VB.CheckBox chkSubPerTable 
         Caption         =   "Create Sub Per Table"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   -74880
         TabIndex        =   1
         Top             =   720
         Width           =   2385
      End
      Begin VB.ComboBox cboOutput 
         Height          =   315
         Left            =   3780
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   810
         Width           =   3345
      End
      Begin VB.CheckBox chkFieldAttributes 
         Caption         =   "Field Attributes"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1545
      End
      Begin VB.ComboBox cboFormat 
         Height          =   315
         Left            =   3780
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   3345
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7260
         Top             =   90
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CheckBox chkTables 
         Caption         =   "Tables"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1545
      End
      Begin VB.CheckBox chkIndexes 
         Caption         =   "Indexes"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1545
      End
      Begin VB.CheckBox chkFieldProperties 
         Caption         =   "Field Properties"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1545
      End
      Begin VB.TextBox txtFolder 
         Height          =   285
         Left            =   3780
         TabIndex        =   12
         Top             =   1140
         Width           =   5055
      End
      Begin VB.CommandButton cmdOutput 
         Caption         =   "..."
         Height          =   285
         Left            =   8880
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1140
         Width           =   375
      End
      Begin VB.CheckBox chkDropTable 
         Caption         =   "DROP TABLE Statement"
         Height          =   195
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   2385
      End
      Begin VB.Label lblOutput 
         AutoSize        =   -1  'True
         Caption         =   "Output"
         Height          =   195
         Left            =   2730
         TabIndex        =   7
         Top             =   870
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Script Format"
         Height          =   195
         Left            =   2730
         TabIndex        =   4
         Top             =   540
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Output Folder"
         Height          =   195
         Left            =   2730
         TabIndex        =   11
         Top             =   1200
         Width           =   960
      End
   End
   Begin RichTextLib.RichTextBox RTF 
      Height          =   4755
      Left            =   2190
      TabIndex        =   15
      Top             =   1650
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   8387
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"script.frx":0496
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000010&
      Caption         =   "Generate"
      Default         =   -1  'True
      Height          =   345
      Left            =   6900
      TabIndex        =   20
      Top             =   6450
      Width           =   1245
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   345
      Left            =   8190
      TabIndex        =   21
      Top             =   6450
      Width           =   1245
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4755
      Left            =   90
      TabIndex        =   14
      Top             =   1650
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   8387
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
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRTB As RTBColourParser
Private mCancelled As Boolean
Public Function GetADOConstant(lValue As Long) As String
'    Select Case lValue
'    Case adTinyInt
'        GetADOConstant = "adTinyInt"
'    Case adSmallInt
'        GetADOConstant = "adSmallInt"
'    Case adInteger
'        GetADOConstant = "adInteger"
'    Case adBigInt
'        GetADOConstant = "adBigInt"
'    Case adUnsignedTinyInt
'        GetADOConstant = "adUnsignedTinyInt"
'    Case adUnsignedSmallInt
'        GetADOConstant = "adUnsignedSmallInt"
'    Case adUnsignedInt
'        GetADOConstant = "adUnsignedInt"
'    Case adUnsignedBigInt
'        GetADOConstant = "adUnsignedBigInt"
'    Case adSingle
'        GetADOConstant = "adSingle"
'    Case adDouble
'        GetADOConstant = "adDouble"
'    Case adCurrency
'        GetADOConstant = "adCurrency"
'    Case adDecimal
'        GetADOConstant = "adDecimal"
'    Case adNumeric
'        GetADOConstant = "adNumeric"
'    Case adBoolean
'        GetADOConstant = "adBoolean"
'    Case adUserDefined
'        GetADOConstant = "adUserDefined"
'    Case adVariant
'        GetADOConstant = "adVariant"
'    Case adGuid
'        GetADOConstant = "adGuid"
'    Case adDate
'        GetADOConstant = "adDate"
'    Case adDBDate
'        GetADOConstant = "adDate"
'    Case adDBTime
'        GetADOConstant = "adDBTime"
'    Case adDBTimestamp
'        GetADOConstant = "adDBTimestamp"
'    Case adBSTR
'        GetADOConstant = "adBSTR"
'    Case adChar
'        GetADOConstant = "adChar"
'    Case adVarChar
'        GetADOConstant = "adVarChar"
'    Case adLongVarChar
'        GetADOConstant = "adLongVarChar"
'    Case adWChar
'        GetADOConstant = "adWChar"
'    Case adVarWChar
'        GetADOConstant = "adVarWChar"
'    Case adLongVarWChar
'        GetADOConstant = "adLongVarWChar"
'    Case adBinary
'        GetADOConstant = "adBinary"
'    Case adVarBinary
'        GetADOConstant = "adVarBinary"
'    Case adLongVarBinary
'        GetADOConstant = "adLongVarBinary"
'    End Select
End Function

Private Sub ScriptQA(tbl As DAO.TableDef, nHandle As Integer, bScriptPerTable As Boolean)
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim bMemo As Boolean
    Dim sBuffer As String
    Dim sText As String
    Dim sTab As String

    On Local Error GoTo ScriptQAError
    sTab = Space$(4)

    If bScriptPerTable Then
        nHandle = FreeFile
        Open txtFolder.Text & tbl.Name & ".sql" For Output As #nHandle
        
        ListView1.ListItems.Add , , tbl.Name & ".sql"
    End If

    Print #nHandle, "-- " & String(60, "*")
                
    If chkDropTable.Value Then
        Print #nHandle, "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[" & tbl.Name & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
        Print #nHandle, "drop table [dbo].[" & tbl.Name & "]"
        Print #nHandle, "GO"
        Print #nHandle, ""
    End If
    
    If chkTables.Value Then
        Print #nHandle, "CREATE TABLE [dbo].[" & tbl.Name & "] ("
        
        bMemo = False
        sBuffer = ""
        
        For Each fld In tbl.Fields
            If (fld.Attributes And dbAutoIncrField) = dbAutoIncrField Then
                sText = "[" & fld.Name & "] [int] IDENTITY (1,1) NOT NULL"
            Else
                Select Case fld.Type
                Case dbText
                    sText = "[" & fld.Name & "] [nvarchar] (" & fld.Size & ") COLLATE Latin1_General_CI_AS"
                Case dbInteger
                    sText = "[" & fld.Name & "] [smallint]"
                Case dbLong
                    sText = "[" & fld.Name & "] [int]"
                Case dbCurrency
                    sText = "[" & fld.Name & "] [money]"
                Case dbSingle, dbDouble
                    sText = "[" & fld.Name & "] [float]"
                Case dbDate
                    sText = "[" & fld.Name & "] [smalldatetime]"
                Case dbBoolean
                    sText = "[" & fld.Name & "] [bit]"
                
                Case dbMemo
                    bMemo = True
                    sText = "[" & fld.Name & "] [ntext] COLLATE Latin1_General_CI_AS"
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
            If idx.Primary Then
                Print #nHandle, "ALTER TABLE [dbo].[" & tbl.Name & "] WITH NOCHECK ADD"
                Print #nHandle, "CONSTRAINT [PK_" & tbl.Name & "] PRIMARY KEY  CLUSTERED"
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
                Print #nHandle, "CREATE NONCLUSTERED INDEX [IK_" & idx.Name & "] ON [dbo].[" & tbl.Name & "]"
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
                    sText = "CONSTRAINT [DF_" & tbl.Name & "_" & fld.Name & "] DEFAULT '' FOR [" & fld.Name & "]"
                Case dbInteger, dbLong, dbCurrency, dbSingle, dbDouble
                    If Val(fld.DefaultValue) <> 0 Then
                        sText = "CONSTRAINT [DF_" & tbl.Name & "_" & fld.Name & "] DEFAULT (" & fld.DefaultValue & ") FOR [" & fld.Name & "]"
                    Else
                        sText = "CONSTRAINT [DF_" & tbl.Name & "_" & fld.Name & "] DEFAULT (0) FOR [" & fld.Name & "]"
                    End If
                Case dbBoolean
                    If Val(fld.DefaultValue) <> 0 Then
                        sText = "CONSTRAINT [DF_" & tbl.Name & "_" & fld.Name & "] DEFAULT (1) FOR [" & fld.Name & "]"
                    Else
                        sText = "CONSTRAINT [DF_" & tbl.Name & "_" & fld.Name & "] DEFAULT (0) FOR [" & fld.Name & "]"
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
            Print #nHandle, "ALTER TABLE [dbo].[" & tbl.Name & "] WITH NOCHECK ADD"
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




Private Sub ScriptADO(dbTbl As DAO.TableDef, nHandle As Integer, bScriptPerTable As Boolean, lScriptCount As Long)
    Dim dbFld As DAO.Field
    Dim fldLU As DAO.Field
    Dim dbIdx As DAO.Index
    Dim sBuffer As String
    Dim sText As String
    Dim sTab As String
    
    On Local Error GoTo ScriptADOError
    sTab = Space$(4)

    If bScriptPerTable Then
        lScriptCount = 0
        Print #nHandle, "Public Sub Create_" & Replace(dbTbl.Name, " ", "_") & "(dbCat As ADOX.Catalog)"
        
        nHandle = FreeFile
        Open txtFolder.Text & dbTbl.Name & ".bas" For Output As #nHandle
        
        ListView1.ListItems.Add , , dbTbl.Name & ".bas"
    ElseIf chkSubPerTable.Value Then
        lScriptCount = 0
        Print #nHandle, "Public Sub Create_" & Replace(dbTbl.Name, " ", "_") & "(dbCat As ADOX.Catalog)"
    End If

    If lScriptCount = 0 Then
        Print #nHandle, sTab & "Dim dbCol As ADOX.Column"
        Print #nHandle, sTab & "Dim dbIdx As ADOX.Index"
        Print #nHandle, sTab & "Dim dbTbl As ADOX.Table"
        
        Print #nHandle, ""
        Print #nHandle, sTab & "'Source Database: " & dbs.Name
        Print #nHandle, ""
    End If

    Print #nHandle, sTab & "'" & String(80, "*")
    Print #nHandle, sTab & "'Code to generate Objects for Table: " & dbTbl.Name

    If chkTables.Value Then
        Print #nHandle, sTab & "Set dbTbl = New ADOX.Table"
        Print #nHandle, sTab & "With dbTbl"
        Print #nHandle, sTab & sTab & "Set .ParentCatalog = dbCat"
        Print #nHandle, sTab & sTab & ".Name = " & Chr$(34) & dbTbl.Name & Chr$(34) & ""
        Print #nHandle, ""
        
        For Each dbFld In dbTbl.Fields
            If (dbFld.Type = dbText) Then
                Print #nHandle, sTab & sTab & ".Columns.Append " & Chr$(34) & dbFld.Name & Chr$(34) & "," & MapDAOConstant(dbFld.Type) & "," & dbFld.Size
            Else
                Print #nHandle, sTab & sTab & ".Columns.Append " & Chr$(34) & dbFld.Name & Chr$(34) & "," & MapDAOConstant(dbFld.Type)
            End If
            
            If chkFieldAttributes.Value And dbFld.Attributes <> 0 Then
                Print #nHandle, sTab & sTab & ".Columns(" & Chr$(34) & dbFld.Name & Chr$(34) & ").Attributes = " & dbFld.Attributes
            End If
            
            If chkFieldProperties.Value Then
                If dbFld.AllowZeroLength Then
                    Print #nHandle, sTab & sTab & ".Columns(" & Chr$(34) & dbFld.Name & Chr$(34) & ").Properties(" & Chr$(34) & "Nullable" & Chr$(34) & ") = " & True
                End If
                If dbFld.Required Then
                    Print #nHandle, sTab & sTab & ".Columns(" & Chr$(34) & dbFld.Name & Chr$(34) & ").Properties(" & Chr$(34) & "Required" & Chr$(34) & ") = " & True
                End If
            End If
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
            If dbIdx.Primary Then
                Print #nHandle, sTab & sTab & ".PrimaryKey = True"
            End If
            If dbIdx.Unique Then
                Print #nHandle, sTab & sTab & ".Unique = True"
            End If

            For Each dbFld In dbIdx.Fields
                Set fldLU = dbTbl.Fields(dbFld.Name)

                If (fldLU.Type = dbText) Then
                    Print #nHandle, sTab & sTab & ".Columns.Append " & Chr$(34) & dbFld.Name & Chr$(34) & "," & MapDAOConstant(fldLU.Type) & "," & fldLU.Size
                Else
                    Print #nHandle, sTab & sTab & ".Columns.Append " & Chr$(34) & dbFld.Name & Chr$(34) & "," & MapDAOConstant(fldLU.Type)
                End If
            Next dbFld

            Print #nHandle, sTab & "End With"
            Print #nHandle, sTab & "dbTbl.Indexes.Append dbIdx"
        Next dbIdx
    End If

    If bScriptPerTable Then
        Print #nHandle, "End Sub"
        Close #nHandle
        nHandle = 0
    ElseIf chkSubPerTable.Value Then
        Print #nHandle, "End Sub"
    Else
        Print #nHandle, ""
        lScriptCount = lScriptCount + 1
    End If
    Exit Sub

ScriptADOError:
    MsgBox Err.Description, vbCritical
    Exit Sub
End Sub






Private Sub ScriptDAO(dbTbl As DAO.TableDef, nHandle As Integer, bScriptPerTable As Boolean, bSQL As Boolean, lScriptCount As Long)
    Dim dbFld As DAO.Field
    Dim fldLU As DAO.Field
    Dim dbIdx As DAO.Index
    Dim sBuffer As String
    Dim sText As String
    Dim sTab As String
    
    On Local Error GoTo ScriptDAOError
    
    sTab = Space$(4)
    
    If bScriptPerTable Then
        lScriptCount = 0
        Print #nHandle, "Public Sub Create_" & Replace(dbTbl.Name, " ", "_") & "(DB As DAO.Database)"
        
        nHandle = FreeFile
        Open txtFolder.Text & dbTbl.Name & ".bas" For Output As #nHandle
        
        ListView1.ListItems.Add , , dbTbl.Name & ".bas"
    ElseIf chkSubPerTable.Value Then
        lScriptCount = 0
        Print #nHandle, "Public Sub Create_" & Replace(dbTbl.Name, " ", "_") & "(DB As DAO.Database)"
    End If

    If lScriptCount = 0 Then
        If bSQL Then
            Print #nHandle, sTab & "Dim SQL As String"
        Else
            Print #nHandle, sTab & "Dim dbFld As DAO.Field"
            Print #nHandle, sTab & "Dim dbIdx As DAO.Index"
            Print #nHandle, sTab & "Dim dbTbl As DAO.TableDef"
        End If
        
        Print #nHandle, ""
        Print #nHandle, sTab & "'Source Database: " & dbs.Name
        Print #nHandle, ""
    End If

    Print #nHandle, sTab & "'" & String(80, "*")
    Print #nHandle, sTab & "'Code to generate Objects for Table: " & dbTbl.Name

    If bSQL Then
        If chkTables.Value Then
            Print #nHandle, sTab & "SQL = " & Chr$(34) & "CREATE TABLE " & dbTbl.Name & " (" & Chr$(34)
    
            sBuffer = ""
            For Each dbFld In dbTbl.Fields
                If (dbFld.Type = dbText) Then
                    sText = sTab & "SQL = SQL & " & Chr$(34) & dbFld.Name & " TEXT (" & dbFld.Size & ")"
                Else
                    sText = sTab & "SQL = SQL & " & Chr$(34) & dbFld.Name & " " & GetSQLType(dbFld.Type)
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

    If bScriptPerTable Then
        Print #nHandle, "End Sub"
        Close #nHandle
        nHandle = 0
    ElseIf chkSubPerTable.Value Then
        Print #nHandle, "End Sub"
    Else
        Print #nHandle, ""
        lScriptCount = lScriptCount + 1
    End If
    Exit Sub

ScriptDAOError:
    MsgBox Err.Description, vbCritical
    Exit Sub
End Sub







Public Function GetSQLType(DataType As Long) As String
    Select Case DataType
    Case dbBinary
        GetSQLType = "BINARY"
    Case dbBoolean
        GetSQLType = "BIT"
    Case dbByte
        GetSQLType = "BYTE"
    Case dbCurrency
        GetSQLType = "CURRENCY"
    Case dbDate
        GetSQLType = "DATETIME"
    Case dbDouble
        GetSQLType = "DOUBLE"
    Case dbInteger
        GetSQLType = "SHORT"
    Case dbLong
        GetSQLType = "LONG"
    Case dbMemo
        GetSQLType = "LONGTEXT"
    Case dbSingle
        GetSQLType = "SINGLE"
    Case dbText
        GetSQLType = "TEXT"
    Case dbTime
        GetSQLType = "DATETIME"
    End Select
End Function



Public Function MapDAOConstant(lValue As Long) As String
    Select Case lValue
    Case dbBigInt
        MapDAOConstant = "adBigInt"
    Case dbBinary
        MapDAOConstant = "adBinary"
    Case dbBoolean
        MapDAOConstant = "adBoolean"
    Case dbByte
        MapDAOConstant = "adByte"
    Case dbChar
        MapDAOConstant = "adByte"
    Case dbCurrency
        MapDAOConstant = "adCurrency"
    Case dbDate
        MapDAOConstant = "adDate"
    Case dbDecimal
        MapDAOConstant = "adDecimal"
    Case dbDouble
        MapDAOConstant = "adDouble"
    Case dbFloat
        MapDAOConstant = "adFloat"
    Case dbGUID
        MapDAOConstant = "adGUID"
    Case dbInteger
        MapDAOConstant = "adInteger"
    Case dbLong
        MapDAOConstant = "adLong"
    Case dbLongBinary
        MapDAOConstant = "adLongVarBinary"
    Case dbMemo
        MapDAOConstant = "adLongVarWChar"
    Case dbNumeric
        MapDAOConstant = "adNumeric"
    Case dbSingle
        MapDAOConstant = "adSingle"
    Case dbText
        MapDAOConstant = "adVarChar"
    Case dbTime
        MapDAOConstant = "adTime"
    Case dbTimeStamp
        MapDAOConstant = "adTimeStamp"
    Case dbVarBinary
        MapDAOConstant = "adVarBinary"
    Case Else
        MapDAOConstant = CStr(lValue)
    End Select
End Function




Public Function GetDAOConstant(lValue As Long) As String
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





Private Sub cboFormat_Click()
    Dim bEnabled(1) As Boolean
    
    Select Case cboFormat.ListIndex
    Case 0
        bEnabled(0) = True
    Case 4
    Case Else
        bEnabled(0) = True
        bEnabled(1) = True
    End Select
    
    lblOutput.Enabled = bEnabled(0)
    cboOutput.Enabled = bEnabled(0)
    
    'chkCreateDatabaseModule.Enabled = bEnabled(0)
    'chkSubPerTable.Enabled = bEnabled(0)
End Sub


Private Sub chkHighlight_Click()
    If Not ListView1.SelectedItem Is Nothing Then
        ListView1_ItemClick ListView1.SelectedItem
    End If
End Sub

Private Sub cmdBrowse_Click()
    Dim sFolder As String
    
    sFolder = GetFolder(Me, "Scripts Folder:")
    If Len(sFolder) > 0 Then
        txtFolder.Text = sFolder
    End If
End Sub


Private Sub cmdClose_Click()
    If cmdClose.Caption = "Cancel" Then
        mCancelled = True
    Else
        Unload Me
    End If
End Sub


Private Sub cmdOK_Click()
    Dim tbl As DAO.TableDef
    Dim lvItem As ListItem
    Dim lCount As Long
    Dim lScriptCount As Long
    Dim nH As Integer
    Dim nHandle As Integer
    Dim nPos As Integer
    
    On Local Error GoTo ScriptTerminate
    
    ListView1.ListItems.Clear
    RTF.Text = ""
    
    If (cboOutput.ListIndex = 1) And cboOutput.Enabled Then
        With CommonDialog1
            .CancelError = True
            .Flags = cdlOFNHideReadOnly
            
            If cboFormat.ListIndex = 0 Then
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
                txtFolder.Text = Left$(.Filename, nPos)
                ListView1.ListItems.Add , , mID$(.Filename, nPos + 1)
            Else
                ListView1.ListItems.Add , , .Filename
            End If
            
            nHandle = FreeFile
            Open .Filename For Output As #nHandle
        End With
    End If
    
    If Len(txtFolder.Text) > 0 Then
        If Right$(txtFolder.Text, 1) <> "\" Then
            txtFolder.Text = txtFolder.Text & "\"
        End If
    End If
        
    Select Case cboFormat.ListIndex
    Case 0
        mRTB.EditorMode = emDefault
        mRTB.LoadKeyWords App.Path & "\QAKeyWords.txt"
        
    Case 4
        mRTB.EditorMode = emVB
        mRTB.LoadKeyWords App.Path & "\VBKeyWords.txt"
        
    Case Else
        mRTB.EditorMode = emVB
        mRTB.LoadKeyWords App.Path & "\VBKeyWords.txt"
        
        If (chkCreateDatabaseModule.Value = vbChecked) Then
            ListView1.ListItems.Add 1, , "CreateDatabase.bas"
            
            nH = FreeFile
            Open txtFolder.Text & "CreateDatabase.bas" For Output As #nH
            Print #nH, "Public Sub CreateDatabase(sDatabase As String)"
            
            If cboFormat.ListIndex = 1 Then
                Print #nH, Space$(4) & "Dim dbCat As ADOX.Catalog"
                Print #nH, ""
                Print #nH, Space$(4) & "Set dbCat = New ADOX.Catalog"
                Print #nH, Space$(4) & "dbCat.Create " & Chr$(34) & "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr$(34) & " & sDatabase"
                Print #nH, ""
                Print #nH, Space$(4) & "'[Code to work with database]"
                Print #nH, ""
                Print #nH, Space$(4) & "Set dbCat = Nothing"
            Else
                Print #nH, Space$(4) & "Dim db As DAO.Database"
                Print #nH, ""
                Print #nH, Space$(4) & "Set db = DBEngine.CreateDatabase(sDatabase, dbLangGeneral, dbVersion40)"
                Print #nH, ""
                Print #nH, Space$(4) & "'[Code to work with database]"
                Print #nH, ""
                Print #nH, Space$(4) & "Set db = Nothing"
            End If
            
            Print #nH, "End Sub"
            Close #nH
        End If
        
        If (cboOutput.ListIndex = 1) And (chkSubPerTable.Value = vbUnchecked) Then
            lScriptCount = 1
            
            If cboFormat.ListIndex = 1 Then
                Print #nHandle, "Public Sub CreateTables (dbCat As ADOX.Catalog)"
                Print #nHandle, Space$(4) & "Dim dbCol As ADOX.Column"
                Print #nHandle, Space$(4) & "Dim dbIdx As ADOX.Index"
                Print #nHandle, Space$(4) & "Dim dbTbl As ADOX.Table"
            Else
                Print #nHandle, "Public Sub CreateTables (DB As DAO.Database)"
                Print #nHandle, Space$(4) & "Dim dbFld As DAO.Field"
                Print #nHandle, Space$(4) & "Dim dbIdx As DAO.Index"
                Print #nHandle, Space$(4) & "Dim dbTbl As DAO.TableDef"
            End If
            
            Print #nHandle, ""
            Print #nHandle, Space$(4) & "'Source Database: " & dbs.Name
            Print #nHandle, ""
        End If
    
    End Select
    
    mCancelled = False
    Screen.MousePointer = vbHourglass
    cmdClose.Caption = "Cancel"
    
    lCount = frmDBObjects.CheckedCount()
    
    lblObject.Caption = ""
    ProgressBar1.Max = lCount
    ProgressBar1.Value = ProgressBar1.Min
    picScripting.Visible = True
    
    Me.Refresh
    
    For Each lvItem In frmDBObjects.lvT.ListItems
        If lvItem.Selected Then
            lblObject.Caption = "Table " & lvItem.Text
            lblObject.Refresh
            
            Set tbl = dbs.TableDefs(lvItem.Text)
            
            Select Case cboFormat.ListIndex
            Case 0
                ScriptQA tbl, nHandle, (cboOutput.ListIndex = 0)
            Case 1
                ScriptADO tbl, nHandle, (cboOutput.ListIndex = 0), lScriptCount
            Case 2
                ScriptDAO tbl, nHandle, (cboOutput.ListIndex = 0), False, lScriptCount
            Case 3
                ScriptDAO tbl, nHandle, (cboOutput.ListIndex = 0), True, lScriptCount
            Case 4
                ScriptClass tbl, nHandle
            End Select
            
            ProgressBar1.Value = ProgressBar1.Value + 1
            DoEvents
        End If
    Next lvItem
    
    Select Case cboFormat.ListIndex
    Case 1, 2, 3
        If (cboOutput.ListIndex = 1) And (chkSubPerTable.Value = vbUnchecked) Then
            Print #nHandle, "End Sub"
        End If

    End Select
    
    
ScriptTerminate:
    If nHandle > 0 Then
        Close #nHandle
    End If
    
    With ListView1
        If .ListItems.Count > 0 Then
            .ListItems(1).Selected = True
            ListView1_ItemClick .SelectedItem
        End If
    End With
    
    Screen.MousePointer = vbDefault
    cmdClose.Caption = "Close"
    picScripting.Visible = False
    
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical
    End If
End Sub

Private Sub Form_Load()
    Dim nValue As Integer
    
    Set mRTB = New RTBColourParser
    
    With ListView1
        .ColumnHeaders.Add , , "File", 2000
        .View = lvwReport
    End With
    
    With cboFormat
        .AddItem "Query Analyser (SQL Server)"
        .AddItem "ADO Objects (VB Code)"
        .AddItem "DAO Objects (VB Code)"
        .AddItem "DAO Execute (VB Code)"
        .AddItem "Class Module (VB Code)"
    End With
    
    With cboOutput
        .AddItem "Create Script per Table"
        .AddItem "Create Single Script"
    End With
    
    chkTables.Value = Val(GetSetting(App.EXEName, "Script", "Tables", vbChecked))
    chkIndexes.Value = Val(GetSetting(App.EXEName, "Script", "Indexes", vbChecked))
    chkFieldAttributes.Value = Val(GetSetting(App.EXEName, "Script", "FieldAttributes", vbChecked))
    chkFieldProperties.Value = Val(GetSetting(App.EXEName, "Script", "FieldProperties", vbUnchecked))
    
    nValue = Val(GetSetting(App.EXEName, "Script", "Format", 0))
    cboFormat.ListIndex = nValue
    
    nValue = Val(GetSetting(App.EXEName, "Script", "Output", 0))
    cboOutput.ListIndex = nValue
    
    txtFolder.Text = GetSetting(App.EXEName, "Script", "Folder", "")
    
    chkDropTable.Value = Val(GetSetting(App.EXEName, "Script", "DropTable", vbChecked))
    
    chkCreateDatabaseModule.Value = Val(GetSetting(App.EXEName, "Script", "CreateDatabaseModule", vbUnchecked))
    chkSubPerTable.Value = Val(GetSetting(App.EXEName, "Script", "Public SubPerTable", vbChecked))
    
    chkHighlight.Value = Val(GetSetting(App.EXEName, "Script", "Highlight", vbChecked))
End Sub


Private Function FormatName(ByVal sName As String) As String
    Mid$(sName, 1, 1) = UCase$(mID$(sName, 1, 1))
    If InStr(sName, " ") > 0 Then
        sName = Replace(sName, " ", "")
    End If
    
    FormatName = sName
End Function


Public Function GetVBType(DataType As Long) As String
    Select Case DataType
    Case dbBinary
        GetVBType = "Variant"
    Case dbBoolean
        GetVBType = "Boolean"
    Case dbByte
        GetVBType = "Byte"
    Case dbCurrency
        GetVBType = "Currency"
    Case dbDate
        GetVBType = "Date"
    Case dbDouble
        GetVBType = "Double"
    Case dbInteger
        GetVBType = "Integer"
    Case dbLong
        GetVBType = "Long"
    Case dbMemo
        GetVBType = "String"
    Case dbSingle
        GetVBType = "Single"
    Case dbText
        GetVBType = "String"
    Case dbTime
        GetVBType = "Date"
    End Select
End Function
Private Sub ScriptClass(dbTbl As DAO.TableDef, nHandle As Integer)
    Dim dbFld As DAO.Field
    Dim sComment As String
    Dim sField As String
    Dim sTab As String
    
    On Local Error GoTo ScriptClassError
    
    sTab = Space$(4)

    nHandle = FreeFile
    Open txtFolder.Text & dbTbl.Name & ".cls" For Output As #nHandle
    
    ListView1.ListItems.Add , , dbTbl.Name & ".cls"
    
    Print #nHandle, ""
    Print #nHandle, "'Source Database: " & dbs.Name
    Print #nHandle, "'Source Table: " & dbTbl.Name
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
        
        Print #nHandle, "Private m" & sField & " As " & GetVBType(dbFld.Type) & sComment
    Next dbFld
    
    Print #nHandle, ""
    Print #nHandle, "Public Sub LoadData (db as DAO.Database)"
    Print #nHandle, Space$(4) & "Dim rs As DAO.RecordSet"
    Print #nHandle, Space$(4) & "Dim SQL As String"
    Print #nHandle, ""
    Print #nHandle, Space$(4) & "'[Code To define SQL]"
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
    Print #nHandle, Space$(4) & "'[Code To define SQL]"
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
        Print #nHandle, "Public Property Get " & sField & " () As " & GetVBType(dbFld.Type)
        Print #nHandle, sTab & sField & " = m" & sField
        Print #nHandle, "End Property"
        Print #nHandle, ""
        Print #nHandle, "Public Property Let " & sField & " (vData As " & GetVBType(dbFld.Type) & ")"
        Print #nHandle, sTab & "m" & sField & " = vData"
        Print #nHandle, "End Property"
        Print #nHandle, ""
    Next dbFld
    
    Close #nHandle
    nHandle = 0
    Exit Sub

ScriptClassError:
    MsgBox Err.Description, vbCritical
    Exit Sub
End Sub










Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        SSTab1.Width = Me.Width - 285
        
        ListView1.Height = Me.Height - 2475
        RTF.Height = Me.Height - 2475
        RTF.Width = Me.Width - 2385
        
        chkHighlight.Top = Me.Height - 720
        cmdOK.Top = Me.Height - 780
        cmdClose.Top = cmdOK.Top
        
        cmdOK.Left = Me.Width - 2760
        cmdClose.Left = Me.Width - 1440
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.EXEName, "Script", "Tables", chkTables.Value
    SaveSetting App.EXEName, "Script", "Indexes", chkIndexes.Value
    SaveSetting App.EXEName, "Script", "FieldAttributes", chkFieldAttributes.Value
    SaveSetting App.EXEName, "Script", "FieldProperties", chkFieldProperties.Value
    
    SaveSetting App.EXEName, "Script", "Format", cboFormat.ListIndex
    SaveSetting App.EXEName, "Script", "Output", cboOutput.ListIndex
    SaveSetting App.EXEName, "Script", "Folder", txtFolder.Text
    
    SaveSetting App.EXEName, "Script", "DropTable", chkDropTable.Value
    
    SaveSetting App.EXEName, "Script", "CreateDatabaseModule", chkCreateDatabaseModule.Value
    SaveSetting App.EXEName, "Script", "Public SubPerTable", chkSubPerTable.Value
    
    SaveSetting App.EXEName, "Script", "Highlight", chkHighlight.Value
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If chkHighlight.Value Then
        mRTB.LoadFile txtFolder.Text & Item.Text, RTF
    Else
        With RTF
            .LoadFile txtFolder.Text & Item.Text
            
            .SelStart = 1
            .SelLength = Len(.Text)
            .SelColor = vbBlack
        End With
    End If
End Sub


