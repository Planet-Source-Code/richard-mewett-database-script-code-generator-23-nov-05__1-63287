VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeyWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"KeyWord"
Attribute VB_Ext_KEY = "Member0" ,"KeyWord"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(Text As String, Colour As Long) As KeyWord
    Dim objNewMember As KeyWord
    Set objNewMember = New KeyWord
    
    objNewMember.Text = Text
    objNewMember.Colour = Colour
    
    mCol.Add objNewMember, Text
    
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Sub Clear()
    Set mCol = New Collection
End Sub

Public Property Get Item(vntIndexKey As Variant) As KeyWord
Attribute Item.VB_UserMemId = 0
    On Local Error GoTo ItemError
    
    Set Item = mCol(vntIndexKey)
    Exit Property

ItemError:
    Set Item = New KeyWord
    Exit Property
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

