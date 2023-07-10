Attribute VB_Name = "ModuleGlobal"
Option Explicit

' リポジトリ親フォルダ
'Public Const parentDir As String = "C:¥Users¥pckohei¥source¥repos¥VBA"
Public Const ParentDir As String = "C:¥Users¥1784¥Source¥Repos¥VBA"

' エラー出力
Public Sub OutputError(errPlace As String, Optional errNote As String)
    Dim msg As String
    msg = vbCrLf
    msg = "日時：" & Format$(Now(), "yyyy/mm/dd hh:nn:ss") & vbCrLf
    msg = msg & "ソース：" & Err.Source & vbCrLf
    msg = msg & "ブック名：" & ActiveWorkbook.Name & vbCrLf
    msg = msg & "場所：" & errPlace & vbCrLf
    msg = msg & "備考：" & errNote & vbCrLf
    msg = msg & "エラー番号：" & Err.Number & vbCrLf
    msg = msg & "エラー内容：" & Err.Description & vbCrLf
    Debug.Print msg
End Sub

Public Sub 指定フォルダが無ければ作る(dirPath As String)

    Dim fso As Object:   Set fso = CreateObject("Scripting.FileSystemObject")
    Dim flds As Variant:    flds = Split(dirPath, "¥")

    Dim i As Integer, fld As String
    For i = 0 To UBound(flds)
        fld = fld & flds(i) & "¥"
        If fld = "¥¥" Then
            i = i + 1
            fld = fld & flds(i) & "¥"
        ElseIf Not fso.FolderExists(fld) Then
            fso.CreateFolder fld
        End If
    Next i

    Set fso = Nothing

End Sub

