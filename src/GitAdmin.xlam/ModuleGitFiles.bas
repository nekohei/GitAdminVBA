Attribute VB_Name = "ModuleGitFiles"
Option Explicit

' Git設定ファイルの内容がコメントで埋め込まれているモジュール
Private Const ContentsModuleName As String = "ModuleGitFilesContents"

' 標準でBOM付きになる為、BOM除去
Public Sub GenerateUTF8(txt As String, filePath As String)

    Dim binData As Variant, sm As Object
    
    Set sm = CreateObject("ADODB.Stream")
    With sm
        .Type = 2               ' 文字列
        .Charset = "utf-8"      ' 文字コード指定
        .Open
        .WriteText txt          ' 文字列を書き込む
        .Position = 0           ' streamの先頭に移動
        .Type = 1               ' バイナリー
        .Position = 3           ' streamの先頭から3バイトをスキップ
        binData = .Read         ' バイナリー取得
        .Close: Set sm = Nothing
    End With
    
    Set sm = CreateObject("ADODB.Stream")
    With sm
        .Type = 1               ' バイナリー
        .Open
        .Write binData          ' バイナリーデータを書き込む
        .SaveToFile filePath, 2 ' 保存
        .Close: Set sm = Nothing
    End With

End Sub

Public Sub GenerateGitFiles()
    
    If Application.VBE.ActiveVBProject Is Nothing Then Exit Sub
    
    Dim reposName As String: reposName = ActiveWorkbook.BuiltinDocumentProperties(5).Value
    Dim xBookName As String: xBookName = ActiveWorkbook.Name
    Dim rootDir As String: rootDir = ParentDir & "¥" & reposName
    Call 指定フォルダが無ければ作る(rootDir & "¥.vscode")
    Dim srcDir As String: srcDir = rootDir & "¥src¥" & xBookName
    Call 指定フォルダが無ければ作る(srcDir)
    Dim filePath As String: filePath = srcDir & "¥"
    
    Dim txt As String, fileName As String
    With ThisWorkbook.VBProject.VBComponents(ContentsModuleName).CodeModule
        Dim i As Long, iLine As String
        For i = 1 To .CountOfLines
            iLine = .Lines(i, 1)
            ' 対象はコメント行のみ
            If UCase(Left(Trim(iLine), 3)) = "REM" Then
                fileName = Trim(Mid(Trim(iLine), 4))
                txt = ""
            ElseIf Left(Trim(iLine), 1) = "'" Then
                ' 先頭シングルクォーテーションの除去
                iLine = Mid(RTrim(iLine), 2)
                txt = txt & iLine & vbCrLf
            End If
            ' 空白行または終端行ならテキスト作成
            If (Trim(iLine) = "" Or i = .CountOfLines) And Trim(txt) <> "" Then
                If fileName = "settings.json" Then
                    Call GenerateUTF8(txt, rootDir & "¥.vscode¥" & fileName)
                Else
                    Call GenerateUTF8(txt, filePath & fileName)
                End If
            End If
        Next
    End With

End Sub
