VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_DebugTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
#Const NO_TRACE = 1
Sub ExpandALL()
    Me.Cells.EntireRow.Hidden = False
End Sub
Sub CollapseALL()
    Dim r As Range
    Dim i As Long
    ' すべて非表示
    Me.UsedRange.EntireRow.Hidden = True
    ' ヘッダ行を再表示
    Set r = Me.Cells(1, 1)
    r.EntireRow.Hidden = False
    Set r = r.Offset(1, 0)
    For i = 2 To Me.UsedRange.Row + Me.UsedRange.Rows.Count - 1
        Set r = Me.Cells(i, 1)
        ' レベル1だけ再表示
        If r.Value = 1 Then r.Cells.EntireRow.Hidden = False
    Next
    Me.Cells(1, 1).Select
End Sub
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    Dim r As Range  ' Range
    Dim i As Long   ' Index
    Dim lv As Long  ' Stack Level
    Dim cnt As Long ' Row Counter
    If Target.Range.Offset(1, 0).EntireRow.Hidden Then
    ' 非表示の行のうち，レベルが1大きい行のみを再表示する。
        Set r = Target.Range.Offset(1, 0)
        lv = Target.Range.EntireRow.Cells(1, 1).Value
        Do While r.EntireRow.Hidden
            If r.EntireRow.Cells(1).Value = lv + 1 Then r.EntireRow.Hidden = False
            Set r = r.Offset(1, 0)
        Loop
    Else
    ' レベルが自分より大きい行を非表示にする。
        Set r = Target.Range.Offset(1, 0)
        lv = Target.Range.EntireRow.Cells(1, 1).Value
        cnt = 0
        Do While r.EntireRow.Cells(1).Value > lv
            cnt = cnt + 1
            Set r = r.Offset(1, 0)
        Loop
        If cnt > 0 Then
            Target.Range.Offset(1, 0).Resize(cnt).EntireRow.Hidden = True
        End If
    End If
End Sub

