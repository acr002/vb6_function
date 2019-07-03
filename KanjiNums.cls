VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KanjiNums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"KanjiNum"
Attribute VB_Ext_KEY = "Member0" ,"KanjiNum"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-----------------------------------------------------------[date: 2018.04.11]
Option Explicit

'***********************************************
Private mCol As Collection
'***********************************************

Public Function Add(ByVal a_buf As String, ByVal a_num As Long, byval a_s_type as string, Optional sKey As String) As KanjiNum
  Dim objNewMember As KanjiNum
  Set objNewMember = New KanjiNum
  objNewMember.buf = a_buf
  objNewMember.num = a_num
  objNewMember.s_type = a_s_type
  If Len(sKey) = 0 Then
    mCol.Add objNewMember
  Else
    mCol.Add objNewMember, sKey
  End If
  Set Add = objNewMember
  Set objNewMember = Nothing
End Function
'-----------------------------------------------------------------------------

Public Property Get Item(vntIndexKey As Variant) As KanjiNum
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property
'-----------------------------------------------------------------------------

Public Property Get Count() As Long
  Count = mCol.Count
End Property
'-----------------------------------------------------------------------------

Public Sub Remove(vntIndexKey As Variant)
  mCol.Remove vntIndexKey
End Sub
'-----------------------------------------------------------------------------

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mCol.[_NewEnum]
End Property
'-----------------------------------------------------------------------------

Private Sub Class_Initialize()
  Set mCol = New Collection
End Sub
'-----------------------------------------------------------------------------

Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub
'-----------------------------------------------------------------------------

