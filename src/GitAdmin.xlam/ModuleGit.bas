Attribute VB_Name = "ModuleGit"
Option Explicit

' ルートの親フォルダ
Public Const ParentDir As String = "Source¥Repos¥VBA_TEST"
' Git設定ファイルの内容が埋め込まれているモジュール名
Private Const ContentsModuleName As String = "ModuleGitFilesContents"
' 作業用一時ファイル名
Private Const TempFileName = "TempGitOutput"
' GitCmd引数用
Public Enum GitCommand
    Stage = 1
    Commit = 2
    Push = 3
End Enum

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

Public Sub CreateReposDir()
    
    If SetReposName = -1 Then Exit Sub
    
    ' リポジトリフォルダ作成
    Dim reposDir As String: reposDir = GetRootDir
    If reposDir = "" Then Exit Sub
    Call CreateDirIfThereNo(reposDir)

    ' サブフォルダ作成
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder reposDir & "¥.vscode"
    fso.CreateFolder reposDir & "¥bin"
    fso.CreateFolder reposDir & "¥src"
    Dim srcDir As String: srcDir = reposDir & "¥src¥" & ActiveWorkbook.Name
    fso.CreateFolder srcDir
    Set fso = Nothing

    ' Git設定ファイル作成
    Call GenerateGitFiles(reposDir, srcDir)
    
End Sub

' 引数のフォルダパスが存在しない場合に作る
Public Sub CreateDirIfThereNo(dirPath As String)

    Dim fso As Object:   Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dirs As Variant:    dirs = Split(dirPath, "¥")

    Dim i As Integer, dr As String
    For i = 0 To UBound(dirs)
        dr = dr & dirs(i) & "¥"
        If dr = "¥¥" Then
            i = i + 1
            dr = dr & dirs(i) & "¥"
        ElseIf Not fso.FolderExists(dr) Then
            fso.CreateFolder dr
        End If
    Next

    Set fso = Nothing

End Sub

' 引数の文字列でUTF-8テキストファイル作成（標準でBOM付きになる為、BOM除去）
Public Sub GenerateUTF8(txt As String, filePath As String)

    Dim binData As Variant, sm As Object
    
    Set sm = CreateObject("ADODB.Stream")
    With sm
        .Type = 2                ' 文字列
        .Charset = "utf-8"       ' 文字コード指定
        .Open                    ' 開く
        .WriteText txt           ' 文字列を書き込む
        .Position = 0            ' streamの先頭に移動
        .Type = 1                ' バイナリー
        .Position = 3            ' streamの先頭から3バイトをスキップ
        binData = .Read          ' バイナリー取得
        .Close: Set sm = Nothing ' 閉じて解放
    End With
    
    Set sm = CreateObject("ADODB.Stream")
    With sm
        .Type = 1                ' バイナリー
        .Open                    ' 開く
        .Write binData           ' バイナリーデータを書き込む
        .SaveToFile filePath, 2  ' 保存
        .Close: Set sm = Nothing ' 閉じて解放
    End With

End Sub

' Git設定ファイルの作成
Public Sub GenerateGitFiles(rootDir As String, srcDir As String)
    
    If Application.VBE.ActiveVBProject Is Nothing Then Exit Sub
    
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
            ' 空白行または最終行ならテキスト作成
            If (Trim(iLine) = "" Or i = .CountOfLines) And Trim(txt) <> "" Then
                If fileName = "settings.json" Then
                    Call GenerateUTF8(txt, rootDir & "¥.vscode¥" & fileName)
                Else
                    Call GenerateUTF8(txt, srcDir & "¥" & fileName)
                End If
            End If
        Next
    End With

End Sub

' xBookブックをsrcフォルダにExport
Private Sub ExportCodeModules(ByVal xBook As Workbook, ByVal srcDir As String)
    On Error GoTo Catch
    Dim vbPjt As VBIDE.VBProject: Set vbPjt = xBook.VBProject
    
    Dim vbCmp As VBIDE.VBComponent
    For Each vbCmp In vbPjt.VBComponents
        Select Case vbCmp.Type
            Case vbext_ct_StdModule
                vbCmp.Export srcDir & "¥" & vbCmp.Name & ".bas"
            Case vbext_ct_MSForm
                vbCmp.Export srcDir & "¥" & vbCmp.Name & ".frm"
            Case vbext_ct_ClassModule
                vbCmp.Export srcDir & "¥" & vbCmp.Name & ".cls"
            Case vbext_ct_Document
                vbCmp.Export srcDir & "¥" & vbCmp.Name & ".dcm"
        End Select
    Next
    GoTo Finally
Catch:
    OutputError "ExportCodeModules"
Finally:
    ' 何もしない
End Sub
 
' コマンドプロンプトでcmd引数を実行
Private Function RunCmd(cmd As String, Optional showInt As Integer = 0, Optional toWait As Boolean = True) As String

    Dim tmpPath As String: tmpPath = Environ$("temp") & "¥" & TempFileName
    Call GenerateUTF8(" ", tmpPath)
    
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Call wsh.Run("cmd /c " & cmd & " > " & tmpPath, showInt, toWait)

    Dim sm As Object: Set sm = CreateObject("ADODB.Stream")
    sm.Type = 2
    sm.Charset = "utf-8"
    sm.Open
    sm.LoadFromFile tmpPath
    Dim tmp As String: tmp = sm.ReadText
    sm.Close: Set sm = Nothing
    
    RunCmd = tmp

End Function

' Gitコマンドを実行
Public Sub GitCmd(cmd As GitCommand, Optional arg As String = Empty)
    On Error GoTo Catch
    Dim rootDir As String: rootDir = GetRootDir
    If rootDir = "" Then
        MsgBox "リポジトリ名が登録されていません。", vbInformation
        Exit Sub
    End If
    Call ChDir(rootDir)
    Dim tmpPath As String: tmpPath = Environ$("temp") & "¥TempGitOutput"
    Call GenerateUTF8("_", tmpPath)
    
    Dim rt As String
    Select Case cmd
    Case Stage
        If MsgBox(ActiveWorkbook.Name & " の変更をステージします。" & vbLf & vbLf & _
                  ActiveWorkbook.Name & " の保存とエクスポートを伴います。", vbInformation + vbOKCancel) = vbOK Then
            Application.DisplayAlerts = False
            If ActiveWorkbook.Path = rootDir & "¥bin" Then
                MsgBox "binフォルダ内の" & ActiveWorkbook.Name & "を開いたままステージ出来ません。" & vbLf & _
                       "ステージはキャンセルされました。, vbInformation"
                GoTo Finally
            Else
                ActiveWorkbook.Save
                Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
                Call fso.CopyFile(ActiveWorkbook.FullName, rootDir & "¥bin¥" & ActiveWorkbook.Name, True)
                Call Decombine
                rt = RunCmd("git add ." & tmpPath)
            End If
        Else
            GoTo Finally
        End If
    Case Commit
        If arg = Empty Then
            arg = InputBox("コミットのメッセージを入力してください。")
            If arg = "" Then GoTo Finally
            If MsgBox("""" & arg & """" & vbLf & vbLf & "このメッセージでコミットします。", vbInformation + vbOKCancel) = vbCancel Then
                GoTo Finally
            End If
        End If
        rt = RunCmd("git commit -m """ & arg & """")
    Case Push
        Dim mBranch As String
        If arg = Empty Then mBranch = "main"
        rt = RunCmd("git push origin " & mBranch)
    End Select
    Debug.Print rt
    GoTo Finally
Catch:
    OutputError "GitCmd"
Finally:
    Application.DisplayAlerts = True
End Sub

' ---メニュー用 ここから------------------------------------
'
Public Sub GitStage()
    Call GitCmd(Stage)
End Sub

Public Sub GitCommit()
    Call GitCmd(Commit)
End Sub

Public Sub GitPush()
    Call GitCmd(Push)
End Sub
'
' ---メニュー用 ここまで------------------------------------

' ルートディレクトリを返す
Private Function GetRootDir() As String
    Dim reposName As String
    reposName = ActiveWorkbook.BuiltinDocumentProperties(5).Value
    If reposName = "" Then
        GetRootDir = ""
    Else
        GetRootDir = Environ$("USERPROFILE") & "¥" & ParentDir & "¥" & reposName
    End If
End Function

' srcフォルダを指定してActiveWorkbookをExport
Public Sub Decombine()
    Dim srcPath As String: srcPath = GetRootDir & "¥src¥" & ActiveWorkbook.Name
    Call CreateDirIfThereNo(srcPath)
    Call ExportCodeModules(ActiveWorkbook, srcPath)
End Sub

' リポジトリ名をブックプロパティのコメント欄に記録
Public Function SetReposName() As Integer
    Dim reposName As String: reposName = ActiveWorkbook.BuiltinDocumentProperties(5).Value
    If reposName = "" Then
        reposName = InputBox("リポジトリ名を英字で入力してください。")
        If reposName = "" Then
            SetReposName = -1
            Exit Function
        End If
        If CheckReposName(reposName) = "" Then
            MsgBox "無効なリポジトリ名です。", vbInformation
            SetReposName = -1
            Exit Function
        End If
        ActiveWorkbook.BuiltinDocumentProperties(5).Value = reposName
        SetReposName = 1
    Else
        SetReposName = 0
    End If
End Function

' リポジトリ名に禁止文字が使われていないかどうかチェック
Private Function CheckReposName(ByVal stg As String) As String
    Dim i As Integer
    For i = 1 To Len(stg)
        Select Case Asc(Mid(stg, i, 1))
        Case 0 To 127
            If InStr("/¥@‾ ", Mid(stg, i, 1)) > 0 Or _
                (i = Len(stg) And Mid(stg, i, 1) = ".") Then _
                    GoTo Invalid
        Case Else
            GoTo Invalid
        End Select
    Next
    CheckReposName = stg
    Exit Function
Invalid:
    CheckReposName = ""
End Function

