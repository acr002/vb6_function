VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KanjiNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-----------------------------------------------------------[date: 2018.04.11]
Option Explicit

'***********************************************
Private m_buf As String
Private m_num As Long
Private m_s_type As String
'***********************************************

Public Property Let s_type(ByVal vData As String)
  m_s_type = vData
End Property
'-----------------------------------------------------------------------------

Public Property Get s_type() As String
  s_type = m_s_type
End Property
'-----------------------------------------------------------------------------

Public Property Let num(ByVal vData As Long)
  m_num = vData
End Property
'-----------------------------------------------------------------------------
Public Property Get num() As Long
  num = m_num
End Property
'-----------------------------------------------------------------------------
Public Property Let buf(ByVal vData As String)
  m_buf = vData
End Property
'-----------------------------------------------------------------------------
Public Property Get buf() As String
  buf = m_buf
End Property
'-----------------------------------------------------------------------------

