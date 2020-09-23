VERSION 5.00
Begin VB.Form frmDBConnection 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Database Connection Builder"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBrowse 
      Cancel          =   -1  'True
      Caption         =   "Browse..."
      Height          =   345
      Left            =   4200
      TabIndex        =   11
      Top             =   990
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1695
      Left            =   60
      TabIndex        =   15
      Top             =   2220
      Width           =   4035
      Begin VB.OptionButton optOLEDB 
         Caption         =   "OLE DB Connection String"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1380
         Width           =   2235
      End
      Begin VB.OptionButton optODBC 
         Caption         =   "ODBC Connection String"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1140
         Value           =   -1  'True
         Width           =   2235
      End
      Begin VB.TextBox txtUID 
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   270
         Width           =   2175
      End
      Begin VB.TextBox txtPWD 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   630
         Width           =   2175
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "UID:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   330
         Width           =   330
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "PWD:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   690
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "(optional)"
         Height          =   195
         Left            =   3210
         TabIndex        =   17
         Top             =   330
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(optional)"
         Height          =   195
         Left            =   3210
         TabIndex        =   19
         Top             =   690
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Method"
      Height          =   2085
      Left            =   60
      TabIndex        =   12
      Top             =   60
      Width           =   4035
      Begin VB.ComboBox cboDatabase 
         Height          =   315
         Left            =   1380
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1650
         Width           =   2535
      End
      Begin VB.ComboBox cboSQLServer 
         Height          =   315
         Left            =   1380
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1290
         Width           =   2535
      End
      Begin VB.OptionButton optDirect 
         Caption         =   "Direct"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1020
         Width           =   2205
      End
      Begin VB.ComboBox cboDSN 
         Height          =   315
         Left            =   390
         TabIndex        =   1
         Top             =   540
         Width           =   3525
      End
      Begin VB.OptionButton optDSN 
         Caption         =   "DSN"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Database:"
         Height          =   195
         Left            =   390
         TabIndex        =   14
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SQL Server:"
         Height          =   195
         Left            =   390
         TabIndex        =   13
         Top             =   1350
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4200
      TabIndex        =   10
      Top             =   570
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   9
      Top             =   150
      Width           =   1245
   End
End
Attribute VB_Name = "frmDBConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Simple GUI for building an ODBC/OLEDB connection string.

'Credits: Vlad Vissoultchev for the Enumeration of SQL Servers and associated databases

'################################################################################

Private Const SQL_ERROR                     As Integer = -1
Private Const SQL_INVALID_HANDLE            As Integer = -2
Private Const SQL_NEED_DATA                 As Integer = 99
Private Const SQL_NO_DATA_FOUND             As Integer = 100
Private Const SQL_SUCCESS                   As Integer = 0
Private Const SQL_SUCCESS_WITH_INFO         As Integer = 1
'--- for SQLSetConnectOption
Private Const SQL_ATTR_LOGIN_TIMEOUT        As Long = 103
Private Const SQL_ATTR_CONNECTION_TIMEOUT   As Long = 113
Private Const SQL_ATTR_QUERY_TIMEOUT        As Long = 0
Private Const SQL_COPT_SS_BASE              As Long = 1200
Private Const SQL_COPT_SS_INTEGRATED_SECURITY As Long = (SQL_COPT_SS_BASE + 3) ' Force integrated security on login
Private Const SQL_COPT_SS_BASE_EX           As Long = 1240
Private Const SQL_COPT_SS_BROWSE_CACHE_DATA As Long = (SQL_COPT_SS_BASE_EX + 5) ' Determines if we should cache browse info. Used when returned buffer is greater then ODBC limit (32K)
'--- param type
Private Const SQL_IS_UINTEGER               As Integer = (-5)
Private Const SQL_IS_INTEGER                As Integer = (-6)
'--- for SQL_COPT_SS_INTEGRATED_SECURITY
Private Const SQL_IS_OFF                    As Long = 0
Private Const SQL_IS_ON                     As Long = 1
'--- for SQL_COPT_SS_BROWSE_CACHE_DATA
Private Const SQL_CACHE_DATA_NO             As Long = 0
Private Const SQL_CACHE_DATA_YES            As Long = 1
'--- for SQLSetEnvAttr
Private Const SQL_ATTR_ODBC_VERSION         As Long = 200
Private Const SQL_OV_ODBC3                  As Long = 3

Private Const SQL_FETCH_NEXT As Long = 1

Private Declare Function SQLDataSources Lib "odbc32.dll" (ByVal hEnv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer

Private Declare Function SQLAllocEnv Lib "odbc32.dll" (phEnv As Long) As Integer
Private Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal hEnv As Long, phDbc As Long) As Integer
Private Declare Function SQLSetEnvAttr Lib "odbc32" (ByVal EnvironmentHandle As Long, ByVal Attrib As Long, Value As Any, ByVal StringLength As Long) As Integer
Private Declare Function SQLBrowseConnect Lib "odbc32.dll" (ByVal hDbc As Long, ByVal szConnStrIn As String, ByVal cbConnStrIn As Integer, ByVal szConnStrOut As String, ByVal cbConnStrOutMax As Integer, pcbConnStrOut As Integer) As Integer
Private Declare Function SQLDisconnect Lib "odbc32.dll" (ByVal hDbc As Long) As Integer
Private Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hDbc As Long) As Integer
Private Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal hEnv As Long) As Integer
Private Declare Function SQLSetConnectOption Lib "odbc32.dll" (ByVal ConnectionHandle As Long, ByVal Option_ As Integer, ByVal Value As Long) As Integer
Private Declare Function SQLGetConnectOption Lib "odbc32.dll" (ByVal ConnectionHandle As Long, ByVal Option_ As Integer, Value As Long) As Integer
Private Declare Function SQLError Lib "odbc32.dll" (ByVal EnvironmentHandle As Long, ByVal ConnectionHandle As Long, ByVal StatementHandle As Long, ByVal Sqlstate As String, NativeError As Long, ByVal MessageText As String, ByVal BufferLength As Integer, TextLength As Integer) As Integer
'--- ODBC 3.0
Private Declare Function SQLSetConnectAttr Lib "odbc32" Alias "SQLSetConnectAttrA" (ByVal ConnectionHandle As Long, ByVal Attrib As Long, Value As Any, ByVal StringLength As Long) As Integer
Private Declare Function SQLGetConnectAttr Lib "odbc32" Alias "SQLGetConnectAttrA" (ByVal ConnectionHandle As Long, ByVal Attrib As Long, Value As Any, ByVal BufferLength As Long, StringLength As Long) As Integer

Private Const STR_NO_USER_DBS           As String = "<No user databases>"

'################################################################################
Private mConnectionString As String
Private Sub GetDSNsAndDrivers()
    Dim nCount As Integer
    Dim i As Integer
    Dim sDSNItem As String * 1024
    Dim sDRVItem As String * 1024
    Dim sDSN As String
    Dim sDRV As String
    Dim iDSNLen As Integer
    Dim iDRVLen As Integer
    Dim lHenv As Long         'handle to the environment

    On Error Resume Next
    
    'get the DSNs
    If SQLAllocEnv(lHenv) <> -1 Then
        Do Until i <> SQL_SUCCESS
            sDSNItem = Space$(1024)
            sDRVItem = Space$(1024)
            i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
            sDSN = Left$(sDSNItem, iDSNLen)
            sDRV = Left$(sDRVItem, iDRVLen)
                
            If sDSN <> Space(iDSNLen) Then
                cboDSN.AddItem sDSN
            End If
        Loop
    End If
    
    'For nCount = cboDSN.LBound To cboDSN.UBound
    '    cboDSN(nCount).AddItem "(None)", 0
    '    cboDSN(nCount).ListIndex = 0
    'Next nCount
End Sub




Public Property Get ConnectionString() As String
    ConnectionString = mConnectionString
End Property

Public Property Let ConnectionString(ByVal NewValue As String)
    mConnectionString = NewValue
    
    ParseString mConnectionString
End Property

Private Sub ParseString(sConnection As String)
    Dim nCount As Integer
    Dim nPos As Integer
    Dim sData() As String
    
    sData() = Split(sConnection, ";")
    
    For nCount = LBound(sData) To UBound(sData)
        nPos = InStr(sData(nCount), "=")
        If nPos > 0 Then
            Select Case UCase$(Left$(sData(nCount), nPos - 1))
            Case "DSN"
                cboDSN.Text = Mid$(sData(nCount), nPos + 1)
            Case "SERVER"
                cboSQLServer.Text = Mid$(sData(nCount), nPos + 1)
            Case "DATABASE"
                cboDatabase.Text = Mid$(sData(nCount), nPos + 1)
            Case "UID"
                txtUID.Text = Mid$(sData(nCount), nPos + 1)
            Case "PWD"
                txtPWD.Text = Mid$(sData(nCount), nPos + 1)
            End Select
        End If
    Next nCount
End Sub

Private Function EnumSqlDbs(sServer As String, Optional sUser As String, Optional sPass As String) As Variant
    Const CONN_STR      As String = "DRIVER={SQL Server};SERVER=%1;UID=%2;PWD=%3;"
    Const PREFIX        As String = "Database={"
    Const SUFFIX        As String = "}"
    Dim sConnStr        As String
    
    EnumSqlDbs = pvBrowseConnect(Replace(Replace(Replace(CONN_STR, "%1", sServer), "%2", sUser), "%3", sPass), PREFIX, SUFFIX, Len(sUser) = 0)
End Function


Private Sub cboSQLServer_Click()
    Dim vDb As Variant
    
    cboDatabase.Clear
    For Each vDb In EnumSqlDbs(cboSQLServer.Text, txtUID.Text, txtPWD.Text)
        cboDatabase.AddItem vDb
    Next
End Sub


Private Sub cmdBrowse_Click()
    Dim mSQL As DAO.Database
    Dim sDB As String
    
    On Local Error GoTo CancelError
    
    Set mSQL = DBEngine.OpenDatabase(sDB)
    
    ParseString mSQL.Connect
    
    mSQL.Close
    Set mSQL = Nothing
    
CancelError:
    Exit Sub
End Sub


Private Sub cmdCancel_Click()
    mConnectionString = ""
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    If optDirect.Value Then
        If Len(cboSQLServer.Text) = 0 Then
            MsgBox "You must specify a Server name.", vbExclamation
            cboSQLServer.SetFocus
            Exit Sub
        End If
        If Len(cboDatabase.Text) = 0 Then
            MsgBox "You must specify a Database name.", vbExclamation
            cboDatabase.SetFocus
            Exit Sub
        End If
        
        If optODBC.Value Then
            mConnectionString = "DRIVER=SQL Server;SERVER=" & cboSQLServer.Text & ";DATABASE=" & cboDatabase.Text
        Else
            mConnectionString = "Provider=SQLOLEDB;Data Source=" & cboSQLServer.Text & ";Initial Catalog=" & cboDatabase.Text
        End If
    Else
        mConnectionString = "DSN=" & cboDSN.Text
    End If
        
    If Len(txtUID.Text) > 0 Then
        mConnectionString = mConnectionString & ";UID=" & txtUID.Text
    End If
    If Len(txtPWD.Text) > 0 Then
        mConnectionString = mConnectionString & ";PWD=" & txtPWD.Text
    End If
    
    If optOLEDB.Value Then
        mConnectionString = Replace(mConnectionString, ";UID", ";User Id")
        mConnectionString = Replace(mConnectionString, ";PWD", ";Password")
    End If
    
    Me.Hide
End Sub


Private Sub Form_Load()
    optDirect.Value = True
    
    GetDSNsAndDrivers
    GetSQLServers
End Sub

Private Sub GetSQLServers()
    Dim vSrv As Variant
    
    For Each vSrv In EnumSqlServers
        cboSQLServer.AddItem vSrv
    Next
End Sub
Private Function EnumSqlServers() As Variant
    Const CONN_STR      As String = "DRIVER={SQL Server}"
    Const PREFIX        As String = "Server={"
    Const SUFFIX        As String = "}"
    
    EnumSqlServers = pvBrowseConnect(CONN_STR, PREFIX, SUFFIX)
End Function


Private Function pvBrowseConnect(sConnStr As String, sPrefix As String, sSuffix As String, Optional ByVal bItegrated As Boolean)
    Const FUNC_NAME     As String = "pvBrowseConnect"
    Dim rc              As Integer
    Dim hEnv            As Long
    Dim hDbc            As Long
    Dim sBuffer         As String
    Dim nReqBufSize     As Integer
    Dim lStart          As Long
    Dim lEnd            As Long
    Dim dwSec           As Long
    Dim lStrLen         As Long

    '--- init environment
    rc = SQLAllocEnv(hEnv)
    rc = SQLSetEnvAttr(hEnv, SQL_ATTR_ODBC_VERSION, ByVal SQL_OV_ODBC3, SQL_IS_INTEGER)
    '--- init conn
    rc = SQLAllocConnect(hEnv, hDbc)
    '--- timeouts to ~5 secs
    rc = SQLSetConnectOption(hDbc, SQL_ATTR_CONNECTION_TIMEOUT, 3)
    rc = SQLSetConnectOption(hDbc, SQL_ATTR_LOGIN_TIMEOUT, 3)
    '--- integrated security
    If bItegrated Then
        rc = SQLSetConnectOption(hDbc, SQL_COPT_SS_INTEGRATED_SECURITY, SQL_IS_ON)
    End If
    '--- improve performance
    rc = SQLSetConnectOption(hDbc, SQL_COPT_SS_BROWSE_CACHE_DATA, SQL_CACHE_DATA_YES)
    '--- initial buffer size
    nReqBufSize = 1000
    '--- repeat getting info until buffer gets large enough
    Do
        sBuffer = String(nReqBufSize + 1, 0)
        rc = SQLBrowseConnect(hDbc, sConnStr, Len(sConnStr), sBuffer, Len(sBuffer), nReqBufSize)
    Loop While rc = SQL_NEED_DATA And nReqBufSize >= Len(sBuffer)
    '--- if ok -> parse buffer
    If rc = SQL_SUCCESS Or rc = SQL_NEED_DATA Then
        '--- find prefix
        lStart = InStr(1, sBuffer, sPrefix)
        If lStart > 0 Then
            lStart = lStart + Len(sPrefix)
            '--- find suffix
            lEnd = InStr(lStart, sBuffer, sSuffix)
            If lEnd > 0 Then
                lEnd = lEnd - Len(sSuffix) + 1
                '--- success
                pvBrowseConnect = Split(Mid(sBuffer, lStart, lEnd - lStart), ",")
            End If
        Else
            Err.Raise vbObjectError, "ODBC", pvGetError(rc, hEnv, hDbc, 0)
        End If
    End If
    '--- disconnect
    rc = SQLDisconnect(hDbc)
    '--- free handles
    rc = SQLFreeConnect(hDbc)
    rc = SQLFreeEnv(hEnv)
    '--- on failure -> return Array(0 To -1)
    If Not IsArray(pvBrowseConnect) Then
        pvBrowseConnect = Split("")
    End If
End Function


Private Function pvGetError(ByVal rc As Long, ByVal hEnv As Long, ByVal hDbc As Long, ByVal hStm As Long) As String
    Dim sSqlState       As String * 5
    Dim lNativeError    As Long
    Dim sMsg            As String * 512
    Dim nTextLength     As Integer
    
    SQLError hEnv, hDbc, hStm, sSqlState, lNativeError, sMsg, Len(sMsg), nTextLength
    pvGetError = "ODBC Result: 0x" & Hex(rc) & vbCrLf & vbCrLf & Left(sMsg, nTextLength)
End Function


Private Sub optDirect_Click()
    cboDSN.Enabled = False
    cboSQLServer.Enabled = True
    cboDatabase.Enabled = True
End Sub


Private Sub optDSN_Click()
    cboDSN.Enabled = True
    cboSQLServer.Enabled = False
    cboDatabase.Enabled = False
End Sub


