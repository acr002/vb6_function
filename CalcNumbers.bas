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
  col.Add "•SO\“ρ–Oά\"
  col.Add "άηl•SO\“ρ–"
  col.Add "άηl•SO\“ρ"
  col.Add "άηlO“ρ"
  col.Add "“ρ“ρ“ρ"
  col.Add "lll"
  col.Add "κ–‹γ"
  col.Add "κZZZ‹γ"
  col.Add "κ“ρOlάZµ”ª"
  col.Add "ηά\“ρ"
  col.Add "ηά“ρ"
  col.Add "κZ•SZ‹γ"
  col.Add "κZ•SZ‹γ\"
  col.Add "κZ•S‹γ\"
  col.Add "ηO•S–"
  col.Add "ηO•S“ρ–"
  col.Add "κκ"
  col.Add "κηά“ρ"
  col.Add "‚Tηά“ρ"
  col.Add "‚Tη‚R‚Q–"
  col.Add "‚R–\ά"
  col.Add "“ρZO"
  col.Add "1"
  col.Add "21"
  col.Add "ά"
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
            ' 2018.04.13(‹ΰ) comments out
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
  ' A(…’²®γA‘«‚µ‡‚ν‚Ή‚ι—v‘f)
  ' B(…’²®—v‘f)
  ' C(…m’θ—v‘f)
  sig.Add "κ", 1, "A"
  sig.Add "“ρ", 2, "A"
  sig.Add "O", 3, "A"
  sig.Add "l", 4, "A"
  sig.Add "ά", 5, "A"
  sig.Add "Z", 6, "A"
  sig.Add "µ", 7, "A"
  sig.Add "”ª", 8, "A"
  sig.Add "‹γ", 9, "A"
  sig.Add "E", 10, "B"
  sig.Add "\", 10, "B"
  sig.Add "•S", 100, "B"
  sig.Add "η", 1000, "B"
  sig.Add "–", 10000, "C"
  sig.Add "δέ", 10000, "C"
  sig.Add "‰­", 100000000, "C"
  sig.Add "—λ", 0, "A"
  sig.Add "Z", 0, "A"
  sig.Add "λ", 1, "A"
  sig.Add "“σ", 2, "A"
  sig.Add "Q", 3, "A"
  sig.Add "ή", 5, "A"
  sig.Add "‚P", 1, "A"
  sig.Add "‚Q", 2, "A"
  sig.Add "‚R", 3, "A"
  sig.Add "‚S", 4, "A"
  sig.Add "‚T", 5, "A"
  sig.Add "‚U", 6, "A"
  sig.Add "‚V", 7, "A"
  sig.Add "‚W", 8, "A"
  sig.Add "‚X", 9, "A"
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
  'sig.Add "’›", 1000000000000#, C
  'sig.Add "‹", 1E+16, C
End Sub
'-----------------------------------------------------------------------------

