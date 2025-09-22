Attribute VB_Name = "MdWorksheet"
Option Explicit

Sub CreateHeaderNames(ByVal ws As Worksheet)
    Dim dic As New Dictionary
    Dim nm As Name, nameStr As String
    For Each nm In ws.Names
        If nm.RefersToRange.Parent.Name = ws.Name Then
            nameStr = nm.Name
            nameStr = Mid(nameStr, InStr(1, nameStr, "!") + 1)
            dic.Add nm.Name, True
        End If
    Next
    ' ワークシートの1行目を走査
    Dim rg As Range
    For Each rg In ws.Rows(1).Cells
        nameStr = Trim(rg.Value)
        If nameStr = "" Then Exit For
        ' 空のセルはスキップ
        If Len(nameStr) > 0 Then
            ' ディクショナリに存在しない場合のみNameを作成
            If Not dic.Exists(nameStr) Then
                On Error Resume Next
                ' Nameを作成
                ws.Names.Add Name:=nameStr, RefersTo:=rg
                If Err.Number = 0 Then
                    ws.Names(nameStr).Comment = nameStr
                    dic.Add nameStr, True
                End If
                On Error GoTo 0
            End If
        End If
    Next
End Sub
