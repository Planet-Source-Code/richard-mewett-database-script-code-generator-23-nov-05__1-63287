VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeyWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mColour As Long
Private mText As String

Public Property Get Text() As String
    Text = mText
End Property

Public Property Let Text(ByVal vNewValue As String)
    mText = vNewValue
End Property

Public Property Get Colour() As Long
    Colour = mColour
End Property

Public Property Let Colour(ByVal vNewValue As Long)
    mColour = vNewValue
End Property
