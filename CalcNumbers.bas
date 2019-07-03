Attribute VB_Name = "CalcNumbers"
'-----------------------------------------------------------[date: 2019.07.03]
Option Explicit

'***********************************************
Private sig As KanjiNums
'***********************************************

Public Sub main()
  Form1.Show
End Sub
'-----------------------------------------------------------------------------

Public Sub cc()
  dim buf as string
  Dim b_buf As Variant
  Dim col As Collection
  Call load_kanjinum
  Set col = New Collection
  col.Add "�S�O�\�񖜎O�܏\"
  col.Add "�ܐ�l�S�O�\��"
  col.Add "�ܐ�l�S�O�\��"
  col.Add "�ܐ�l�O��"
  col.Add "����"
  col.Add "�l�l�l"
  col.Add "�ꖜ��"
  col.Add "��Z�Z�Z��"
  col.Add "���O�l�ܘZ����"
  col.Add "��܏\��"
  col.Add "��ܓ�"
  col.Add "��Z�S�Z��"
  col.Add "��Z�S�Z��\"
  col.Add "��Z�S��\"
  col.Add "��O�S��"
  col.Add "��O�S��"
  col.Add "���"
  col.Add "���ܓ�"
  col.Add "�T��ܓ�"
  col.Add "�T��R�Q��"
  col.Add "�R���\��"
  col.Add "��Z�O"
  col.Add "1"
  col.Add "21"
  col.Add "��"
  For Each b_buf In col
    Debug.Print b_buf, format(calc_number(b_buf), "#,0")
    buf = buf & vbcrlf & b_buf & space(N1) & format(calc_number(b_buf), "#,0")
  Next b_buf
  msgbox buf
End Sub
'-----------------------------------------------------------------------------

Public Function calc_number(ByVal a_buf As String) As Long
  Dim i   As Long
  Dim b_sig As KanjiNum
  Dim buf As String
  Dim tbuf As String
  Dim na  As Long
  Dim nb  As Long
  Dim nc  As Long
  'buf = StrConv(Trim(a_buf), vbNarrow)
  buf = Trim(a_buf)
  For i = N1 To Len(buf)
    tbuf = Mid(buf, i, N1)
    For Each b_sig In sig
      If tbuf = b_sig.buf Then
        select case b_sig.s_type
          case "A"
            na = (na * N10) + b_sig.num
          case "B"
            If na = zero Then
              na = N1
            End If
            na = na * b_sig.num
            nb = nb + na
            na = zero
          case "C"
            ' 2018.04.13(��) comments out
            'if na = zero then
            '  na = N1
            'end if
            nc = nc + (na + nb) * b_sig.num
            na = zero
            nb = zero
          case else
        end select
        Exit For
      End If
    Next b_sig
    'If Val(tbuf) Then
    '  na = (na * N10) + Val(tbuf)
    'else
    '  if tbuf = "0" then
    '    na = na * N10
    '  end if
    'End If
  Next i
  calc_number = na + nb + nc
End Function
'-----------------------------------------------------------------------------

Private Sub load_kanjinum()
  Set sig = New KanjiNums
  ' A(��������A�������킹��v�f)
  ' B(�������v�f)
  ' C(���m��v�f)
  sig.Add "��", 1, "A"
  sig.Add "��", 2, "A"
  sig.Add "�O", 3, "A"
  sig.Add "�l", 4, "A"
  sig.Add "��", 5, "A"
  sig.Add "�Z", 6, "A"
  sig.Add "��", 7, "A"
  sig.Add "��", 8, "A"
  sig.Add "��", 9, "A"
  sig.Add "�E", 10, "B"
  sig.Add "�\", 10, "B"
  sig.Add "�S", 100, "B"
  sig.Add "��", 1000, "B"
  sig.Add "��", 10000, "C"
  sig.Add "��", 10000, "C"
  sig.Add "��", 100000000, "C"
  sig.Add "��", 0, "A"
  sig.Add "�Z", 0, "A"
  sig.Add "��", 1, "A"
  sig.Add "��", 2, "A"
  sig.Add "�Q", 3, "A"
  sig.Add "��", 5, "A"
  sig.Add "�P", 1, "A"
  sig.Add "�Q", 2, "A"
  sig.Add "�R", 3, "A"
  sig.Add "�S", 4, "A"
  sig.Add "�T", 5, "A"
  sig.Add "�U", 6, "A"
  sig.Add "�V", 7, "A"
  sig.Add "�W", 8, "A"
  sig.Add "�X", 9, "A"
  sig.Add "0", 0, "A"
  sig.Add "1", 1, "A"
  sig.Add "2", 2, "A"
  sig.Add "3", 3, "A"
  sig.Add "4", 4, "A"
  sig.Add "5", 5, "A"
  sig.Add "6", 6, "A"
  sig.Add "7", 7, "A"
  sig.Add "8", 8, "A"
  sig.Add "9", 9, "A"
  'sig.Add "��", 1000000000000#, C
  'sig.Add "��", 1E+16, C
End Sub
'-----------------------------------------------------------------------------

