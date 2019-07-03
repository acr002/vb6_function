'-----------------------------------------------------------[date: 2019.07.03]
Attribute VB_Name = "Module1"
Option Explicit

'***********************************************
' 2019.07.03(��)
'***********************************************
Private Enum numbers
  zero = 0
  N1 = 1
  N2 = 2
  N3 = 3
  N4 = 4
  N5 = 5
  N6 = 6
  N7 = 7
  N8 = 8
  N9 = 9
End Enum
'***********************************************

Public Sub main()
  Dim ii     As Variant
  Dim col    As Collection
  Dim t_path As String
  t_path = ThisWorkbook.Path & "\02 collect files.bas"
  Set col = col_file_contens(t_path)
  Debug.Print col.Count
  Set col = col_file_contens(t_path, flag_all:=True)
  Debug.Print col.Count
  'For Each ii In col
  '  Debug.Print ii
  'Next ii
End Sub
'-----------------------------------------------------------------------------

' 2019.07.03(��)
' �e�L�X�g�t�@�C���̒��g���i�[�����R���N�V������Ԃ��܂��B
' ��������zero�̍s�͏����܂��B
' �R�����g�s���܂߂邩�ǂ����͑��������w�肵�Ă��������B
Private Function col_file_contens(a_path As String, Optional flag_all As Boolean = False) As Collection
  Dim col      As Collection
  Dim ff_in    As Long
  Dim buf_base As String
  Dim buf      As String
  Set col = New Collection
  If Len(Dir(a_path)) Then
    ff_in = FreeFile()
    Open a_path For Input As #ff_in
    Do Until EOF(ff_in)
      Line Input #ff_in, buf_base
      buf = Trim(buf_base)
      If Len(buf) Then
        If flag_all Then
          col.Add buf_base
        Else
          Select Case Mid(buf, 1, 1)
            Case "*", "'"
            Case Else
              col.Add buf_base
          End Select
        End If
      End If
    Loop
    Close #ff_in
  End If
  Set col_file_contens = col
End Function
'-----------------------------------------------------------------------------

