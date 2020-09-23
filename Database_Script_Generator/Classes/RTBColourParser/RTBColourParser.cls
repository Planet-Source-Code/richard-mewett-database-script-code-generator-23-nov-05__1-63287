VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RTBColourParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"KeyWords"
Option Explicit

'#############################################################################################################################
'Title:     RTBColourParser (A Class for color highlighting a Rich Text Box)
'Author:    Richard Mewett
'Created:   01/09/04
'Version:   1.0.1 (8th September 2004)

'Copyright © 2004 Richard Mewett. All rights reserved.

'Credits: Some code inspired by Will Bardens RTF articles on PSC

'#############################################################################################################################
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Enum RTBEditorMode
    emDefault = 0
    emVB = 1
    emVC = 2
End Enum

Private Const VB_COMMENT_COLOUR = &H8000&
Private Const VB_COMMENT_TAG = "'"
Private Const VC_COMMENT_COLOUR = &H8000&
Private Const VC_COMMENT_TAG = "//"
Private Const VCPP_COMMENT_TAG = "/*"

Private mKeyWords As KeyWords
Private mRTBEditorMode As RTBEditorMode

Private mSeperators() As String

Public Event BeginLoading(Max As Long)
Public Event Loading(Value As Long)
Public Event EndLoading()
Public Property Set KeyWords(ByVal vData As KeyWords)
    Set mKeyWords = vData
End Property


Public Property Get KeyWords() As KeyWords
    If mKeyWords Is Nothing Then
        Set KeyWords = New KeyWords
    End If
    
    Set KeyWords = mKeyWords
End Property




Public Sub LoadKeyWords(ByVal Filename As String)
    Dim nHandle As Integer
    Dim nPos As Integer
    Dim sData As String
    Dim sValue As String
    
    On Local Error GoTo LoadKeyWordsError
    
    Set KeyWords = New KeyWords
    
    nHandle = FreeFile
    Open Filename For Input As #nHandle
    
    Do Until EOF(nHandle)
        Line Input #nHandle, sData
        
        nPos = InStr(sData, vbTab)
        If nPos > 0 Then
            sValue = Mid$(sData, nPos + 1)
            If Len(sValue) > 0 Then
                mKeyWords.Add Left$(sData, nPos - 1), Val(sValue)
            End If
        End If
    Loop
    
    Close #nHandle
    Exit Sub

LoadKeyWordsError:
    Exit Sub
End Sub


Public Sub LoadFile(ByVal Filename As String, RTB As RichTextBox)
    Dim lCount As Long
    Dim lPos(1) As Long
    Dim sData() As String
    Dim sText As String
    
    Dim dStart As Single
    
    On Local Error GoTo LoadFileError
    
    LockWindowUpdate RTB.hwnd

    With RTB
        .LoadFile Filename, 1
        
        sText = .Text
        If (Len(sText) > 0) And (mKeyWords.Count > 0) Then
            .SelStart = 1
            .SelLength = Len(sText)
            .SelColor = vbBlack
            
            lPos(0) = 1
            
            RaiseEvent BeginLoading(Len(sText))
            
            dStart = Timer
            
            'sData() = Split(sText, vbCrLf, , vbBinaryCompare)
            'For lCount = LBound(sData()) To UBound(sData())
            'Next lCount
            
            Do While lPos(0) < Len(sText)
                lPos(1) = InStr(lPos(0) + 1, sText, vbCrLf)
                If lPos(1) = 0 Then
                    lPos(1) = Len(sText)
                End If
                FormatText RTB, lPos(0), lPos(1)
                
                lCount = lCount + 1
                If (lCount Mod 10) = 0 Then
                    RaiseEvent Loading(lPos(0))
                    DoEvents
                End If
                
                lPos(0) = lPos(1) + Len(vbCrLf)
            Loop
            
            RaiseEvent EndLoading
        End If
        
        .SelStart = 0
    End With
    
LoadFileError:
    LockWindowUpdate 0
    Exit Sub
End Sub


Public Sub SaveKeyWords(ByVal Filename As String)
    Dim KW As KeyWord
    Dim nHandle As Integer
    
    On Local Error GoTo SaveKeyWordsError
    
    nHandle = FreeFile
    Open Filename For Output As #nHandle
    
    For Each KW In mKeyWords
        Print #nHandle, KW.Text & vbTab & KW.Colour
    Next KW
    
    Close #nHandle
    Exit Sub

SaveKeyWordsError:
    Exit Sub
End Sub




Public Sub FormatText(RTB As RichTextBox, ByVal lStart As Long, ByVal lEnd As Long, Optional LockRedraw As Boolean)
    Dim lColour As Long
    Dim lPos As Long
    Dim nWord As Integer
    Dim bParseText As Boolean
    Dim sWords() As String
    Dim sLine As String
    Dim sText As String
    
    If LockRedraw Then
        LockWindowUpdate RTB.hwnd
    End If

    With RTB
        sLine = Trim$(Mid$(.Text, lStart, lEnd - lStart))
        If sLine <> vbCrLf Then
            sText = .Text
        
            If (Left$(sLine, 2) = vbCrLf) And (Right$(sLine, 2) = vbCrLf) Then
                sLine = Mid$(sLine, 3, Len(sLine) - 4)
            ElseIf (Left$(sLine, 2) = vbCrLf) Then
                sLine = Mid$(sLine, 3)
            ElseIf (Right$(sLine, 2) = vbCrLf) Then
                sLine = Left$(sLine, Len(sLine) - 2)
            End If
            
            Select Case mRTBEditorMode
            Case emVB
                sLine = FilterVBString(sLine)
            End Select
            
            sWords = Tokenise(sLine)
            
            lPos = 1
            For nWord = LBound(sWords) To UBound(sWords)
                If Len(sWords(nWord)) > 0 Then
                    Select Case mRTBEditorMode
                    Case emVB
                        If (Left$(sWords(nWord), 1) = VB_COMMENT_TAG) Then
                            bParseText = False
                            .SelStart = InStr(lStart, sText, sWords(nWord)) - 1
                            .SelLength = Len(sWords(nWord))
                            .SelColor = VB_COMMENT_COLOUR
                        Else
                            bParseText = True
                        End If
                    
                    Case emVC
                        If (Left$(sWords(nWord), 2) = VC_COMMENT_TAG) Or (Left$(sWords(nWord), 2) = VCPP_COMMENT_TAG) Then
                            bParseText = False
                            .SelStart = InStr(lStart, sText, sWords(nWord)) - 1
                            .SelLength = Len(sWords(nWord))
                            .SelColor = VC_COMMENT_COLOUR
                        Else
                            bParseText = True
                        End If
                    
                    Case Else
                        bParseText = True
                    
                    End Select
                    
                    If bParseText Then
                        lColour = mKeyWords(sWords(nWord)).Colour
                        If lColour Then
                            .SelStart = InStr(lStart + lPos - 1, sText, sWords(nWord)) - 1
                            .SelLength = Len(sWords(nWord))
                            .SelColor = lColour
                        End If
                    End If
                End If
                
                lPos = lPos + Len(sWords(nWord))
            Next nWord
        End If
    End With
    
    If LockRedraw Then
        LockWindowUpdate 0
    End If
End Sub



Private Function FilterVBString(ByVal sText As String) As String
    Dim nComment As Integer
    Dim nQuote(1) As Integer

    nComment = InStr(sText, VB_COMMENT_TAG)
    nQuote(0) = InStr(sText, Chr$(34))
    If (nComment = 0) Or (nQuote(0) < nComment) Then
        Do While nQuote(0) > 0
            nQuote(1) = InStr(nQuote(0) + 1, sText, Chr$(34))
            If nQuote(1) > 0 Then
                sText = Left$(sText, nQuote(0) - 1) & Mid$(sText, nQuote(1) + 1)
                nComment = InStr(nQuote(1) + 1, sText, VB_COMMENT_TAG)
                nQuote(0) = InStr(nQuote(1) + 1, sText, Chr$(34))
                If nQuote(0) > nComment Then
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop
    End If
    
    FilterVBString = sText
End Function

Private Function Tokenise(ByVal sText As String) As String()
    Dim lPos As Long
    Dim lSeperator As Long
    Dim lSplit As Long
    Dim lTextLength As Long
    Dim nCount As Integer
    Dim nWords As Integer
    Dim sWords() As String
    
    ReDim sWords(100)
    
    lPos = 1
    lTextLength = Len(sText)
    
    Do
        lSplit = lTextLength + 1
        For nCount = 0 To UBound(mSeperators)
            lSeperator = InStr(lPos, sText, mSeperators(nCount))
            If (lSeperator > 0) And (lSeperator < lSplit) Then
                lSplit = lSeperator
            End If
        Next nCount
        
        If lSplit = lTextLength + 1 Then
            sWords(nWords) = Mid$(sText, lPos)
        Else
            Select Case mRTBEditorMode
            Case emVB
                If Mid$(sText, lSplit, 1) = VB_COMMENT_TAG Then
                    ReDim Preserve sWords(nWords + 1)
                    
                    sWords(nWords) = Mid$(sText, lPos, lSplit - lPos)
                    sWords(nWords = 1) = Mid$(sText, lSplit)
                    Tokenise = sWords
                    Exit Function
                End If
                
            Case emVC
                If (Mid$(sText, lSplit, 2) = VC_COMMENT_TAG) Or Mid$(sText, lSplit, 2) = (VCPP_COMMENT_TAG) Then
                    ReDim Preserve sWords(nWords + 1)
                    
                    sWords(nWords) = Mid$(sText, lPos, lSplit - lPos)
                    sWords(nWords = 1) = Mid$(sText, lSplit)
                    Tokenise = sWords
                    Exit Function
                End If
            
            End Select
            
            sWords(nWords) = Mid$(sText, lPos, lSplit - lPos)
        End If
        
        nWords = nWords + 1
        If nWords > UBound(sWords) Then
            ReDim Preserve sWords(nWords + 10)
        End If
    
        lPos = lSplit + 1
        
        If lPos >= lTextLength Then
            Exit Do
        End If
    Loop

    If nWords > 0 Then
        ReDim Preserve sWords(nWords)
    Else
        ReDim sWords(0)
    End If
    Tokenise = sWords
End Function



Public Property Get EditorMode() As RTBEditorMode
    EditorMode = mRTBEditorMode
End Property

Public Property Let EditorMode(ByVal vNewValue As RTBEditorMode)
    mRTBEditorMode = vNewValue
    
    Select Case mRTBEditorMode
    Case emDefault
        ReDim mSeperators(1)
        mSeperators(0) = " "
        mSeperators(1) = "."
        
    Case emVB
        ReDim mSeperators(8)
        mSeperators(0) = " "
        mSeperators(1) = "."
        mSeperators(2) = ","
        mSeperators(3) = "="
        mSeperators(4) = "("
        mSeperators(5) = ")"
        mSeperators(6) = "<"
        mSeperators(7) = ">"
        mSeperators(8) = VB_COMMENT_TAG
    
    Case emVC
        ReDim mSeperators(10)
        mSeperators(0) = " "
        mSeperators(1) = "."
        mSeperators(2) = ","
        mSeperators(3) = "="
        mSeperators(4) = "("
        mSeperators(5) = ")"
        mSeperators(6) = "<"
        mSeperators(7) = ">"
        mSeperators(8) = vbTab
        mSeperators(9) = VC_COMMENT_TAG
        mSeperators(10) = VCPP_COMMENT_TAG
    
    End Select
End Property

Private Sub Class_Initialize()
    EditorMode = emDefault
End Sub



