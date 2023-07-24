Attribute VB_Name = "ModuleGit"
Option Explicit

' ルートの親フォルダ
Public Const ParentDir As String = "Source¥Repos¥VBA"
' Git設定ファイルの内容が埋め込まれているモジュール名
Private Const ContentsModuleName As String = "ModuleGitFilesContents"
' 作業用一時ファイル名
Private Const TempFileName = "TempGitOutput"
' GitCmd引数用
Public Enum GitCommand
    Status
    Stage
    Commit
    Push
End Enum

' エラー出力
Public Sub OutputError(errPlace As String, Optional errNote As String)
    Dim msg As String
    msg = vbCrLf
    msg = "日時：" & Format$(Now(), "yyyy/mm/dd hh:nn:ss") & vbCrLf
    msg = msg & "ソース：" & Err.Source & vbCrLf
    msg = msg & "ブック名：" & ActiveWorkbook.name & vbCrLf
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
    If Not fso.FolderExists(reposDir & "¥.vscode") Then fso.CreateFolder reposDir & "¥.vscode"
    If Not fso.FolderExists(reposDir & "¥bin") Then fso.CreateFolder reposDir & "¥bin"
    If Not fso.FolderExists(reposDir & "¥src") Then fso.CreateFolder reposDir & "¥src"
    Dim srcDir As String: srcDir = reposDir & "¥src¥" & ActiveWorkbook.name
    If Not fso.FolderExists(srcDir) Then fso.CreateFolder srcDir

    ' Git設定ファイル作成
    Call GenerateGitFiles(reposDir, srcDir)
    
    ' binフォルダにブックをコピー
    ActiveWorkbook.Save
    Call fso.CopyFile(ActiveWorkbook.FullName, reposDir & "¥bin¥" & ActiveWorkbook.name, True)
    Set fso = Nothing
    
    ' srcフォルダにCodeModuleをExport
    Call Decombine
    
    MsgBox ActiveWorkbook.name & " 用のリポジトリの準備ができました。", vbInformation

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
                vbCmp.Export srcDir & "¥" & vbCmp.name & ".bas"
            Case vbext_ct_MSForm
                vbCmp.Export srcDir & "¥" & vbCmp.name & ".frm"
            Case vbext_ct_ClassModule
                vbCmp.Export srcDir & "¥" & vbCmp.name & ".cls"
            Case vbext_ct_Document
                vbCmp.Export srcDir & "¥" & vbCmp.name & ".dcm"
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
    On Error GoTo Catch
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim msgPath As String: msgPath = Environ$("temp") & "¥gitTmp.log"
    Dim errPath As String: errPath = Environ$("temp") & "¥gitErr.log"
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim rt As Long: rt = wsh.Run("cmd /c " & cmd & " > " & msgPath & " 2> " & errPath, showInt, toWait)
    Dim tmpPath As String
    If rt = 0 Then
        tmpPath = msgPath
    Else
        tmpPath = errPath
    End If
    Dim sm As Object: Set sm = CreateObject("ADODB.Stream")
    sm.Type = 2
    sm.Charset = "utf-8"
    sm.Open
    sm.LoadFromFile tmpPath
    Dim msg As String: msg = sm.ReadText & "/// " & CurDir & " ///"
    sm.Close: Set sm = Nothing
    GoTo Finally
Catch:
    OutputError "RunCmd"
Finally:
    If fso.FileExists(msgPath) Then fso.DeleteFile msgPath
    If fso.FileExists(errPath) Then fso.DeleteFile errPath
    Set fso = Nothing
    RunCmd = msg
End Function

Private Function RunPowerShell(cmd As String, Optional showInt As Integer = 0, Optional toWait As Boolean = False) As String
    
    Dim tmpPath As String: tmpPath = Environ$("temp") & "¥" & TempFileName
    Call GenerateUTF8(" ", tmpPath)
    
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    '  -NoLogo (見出しを出さない)
    '  -ExecutionPolicy RemoteSigned (実行権限を設定)
    '  -Command (PowerShellのコマンドレット構文を記載）
    Call wsh.Run("powershell -NoLogo -ExecutionPolicy RemoteSigned -Command """ & cmd & _
                 " | Out-File -filePath " & tmpPath & " -encoding utf8""", showInt, toWait)
                 
    Dim sm As Object: Set sm = CreateObject("ADODB.Stream")
    sm.Type = 2
    sm.Charset = "utf-8"
    sm.Open
    sm.LoadFromFile tmpPath
    Dim tmp As String: tmp = sm.ReadText
    sm.Close: Set sm = Nothing
    
    RunPowerShell = tmp

End Function

Sub Cls()
    If Application.VBE.ActiveWindow.Caption = "イミディエイト" Then
        Application.SendKeys "^a", False
        Application.SendKeys "{Del}", False
    End If
End Sub

Sub CheckEncoding()
    Dim en As Object: Set en = CreateObject("System.Text.UTF8Encoding")
'    Dim sjis As Object: Set sjis = en.GetEncoding("shift_jis")
    Dim bin As Variant: bin = en.GetBytes_4("依頼NO.")
    Dim deco As Object: Set deco = en.GetDecoder

End Sub

Function ExtractProjectName(fullURL As String) As String
    Dim startCut As Integer: startCut = InStrRev(fullURL, "/") + 1
    Dim endCut   As Integer: endCut = InStr(1, fullURL, ".git")
    ExtractProjectName = Mid(fullURL, startCut, endCut - startCut)
End Function

' Gitコマンドを実行
Public Function GitCmd(cmd As GitCommand, Optional arg As String = Empty, Optional isPowerShell As Boolean = False) As Integer
    On Error GoTo Catch
    Dim rootDir As String: rootDir = GetRootDir
    If rootDir = "" Then
        MsgBox """" & ActiveWorkbook.name & """" & vbLf & vbLf & "リポジトリ名が登録されていません。", vbInformation
        Exit Function
    End If
    
    Call ChDir(rootDir)
    
    Dim rt As String
    Select Case cmd
    Case Status
        If isPowerShell Then
            rt = RunPowerShell("git status")
        Else
            rt = RunCmd("git status")
        End If
    Case Stage
        If MsgBox(ActiveWorkbook.name & " の変更をステージします。" & vbLf & vbLf & _
                  ActiveWorkbook.name & " の保存とエクスポートを伴います。", vbInformation + vbOKCancel) = vbOK Then
            Application.DisplayAlerts = False
            If ActiveWorkbook.path = rootDir & "¥bin" Then
                MsgBox "binフォルダ内の" & ActiveWorkbook.name & "を開いたままステージ出来ません。" & vbLf & _
                       "ステージはキャンセルされました。, vbInformation"
                GoTo Finally
            Else
                ActiveWorkbook.Save
                Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
                Call fso.CopyFile(ActiveWorkbook.FullName, rootDir & "¥bin¥" & ActiveWorkbook.name, True)
                Call Decombine
                If isPowerShell Then
                    rt = RunPowerShell("git add .")
                Else
                    rt = RunCmd("git add .")
                End If
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
        If isPowerShell Then
            rt = RunPowerShell("git commit -m """ & arg & """")
        Else
            rt = RunCmd("git commit -m """ & arg & """")
        End If
    Case Push
        Dim mBranch As String
        If arg = Empty Then mBranch = "main"
        If isPowerShell Then
            rt = RunPowerShell("git push origin " & mBranch)
        Else
            rt = RunCmd("git push origin " & mBranch)
        End If
    End Select
    Debug.Print rt
    GoTo Finally
Catch:
    OutputError "GitCmd"
Finally:
    Application.DisplayAlerts = True
End Function

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
    Dim srcPath As String: srcPath = GetRootDir & "¥src¥" & ActiveWorkbook.name
    Call CreateDirIfThereNo(srcPath)
    Call ExportCodeModules(ActiveWorkbook, srcPath)
End Sub

' リポジトリ名をブックのプロパティのコメント欄に記録
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

Sub ImportDocument(path As String, xlBook As Workbook)
    Dim compos As VBComponents
    Set compos = xlBook.VBProject.VBComponents
    
    Dim impCompo As VBComponent
    Set impCompo = compos.Import(path)
    
    Dim origCompo As VBComponent
    Dim cname As String, bname As String
    cname = impCompo.name
    bname = GetNameFromPath(path) ' Assuming you have a function to get name from path
    
    If cname <> bname Then
        Set origCompo = compos.Item(bname)
    Else
        Dim sht As Worksheet
        Set sht = xlBook.Worksheets.Add()
        Set origCompo = compos.Item(sht.CodeName)
        
        Dim tmpname As String
        tmpname = "ImportTemp"
        While ComponentExists(compos, tmpname)
            tmpname = tmpname & "1"
        Wend
        
        impCompo.name = tmpname
        origCompo.name = cname
    End If
    
    Dim imod As CodeModule, omod As CodeModule
    Set imod = impCompo.CodeModule
    Set omod = origCompo.CodeModule
    omod.DeleteLines 1, omod.CountOfLines
    omod.AddFromString imod.Lines(1, imod.CountOfLines)
    
    compos.Remove impCompo
End Sub

Function GetNameFromPath(path As String) As String
    ' Function to get name from path
    GetNameFromPath = Mid(path, InStrRev(path, "¥") + 1, Len(path))
End Function

Function ComponentExists(compos As VBComponents, name As String) As Boolean
    Dim c As VBComponent
    On Error Resume Next
    Set c = compos.Item(name)
    If Err.Number = 0 Then
        ComponentExists = True
    Else
        ComponentExists = False
    End If
    On Error GoTo 0
End Function

Public Sub GitInit()
    On Error GoTo Catch
    Dim urlStr As String: urlStr = InputBox("リモートリポジトリのURLを入力してください。")
    If urlStr = "" Then GoTo Finally
    Dim reposName As String: reposName = ExtractProjectName(urlStr)
    ActiveWorkbook.BuiltinDocumentProperties(5).Value = reposName
    Dim reposPath As String: reposPath = Environ$("USERPROFILE") & "¥" & ParentDir & "¥" & reposName
    Call ChDir(reposPath)
            
    Dim cmdStr As String: cmdStr = "echo # & " & reposName & " >> README.md"
    Dim rt As Long: rt = RunCmd(cmdStr)
    rt = GitCmd(Stage)
    rt = GitCmd(Commit, "リポジトリ開始")
    cmdStr = "git branch -M main"
    rt = RunCmd(cmdStr)
    cmdStr = "git remote add origin " & urlStr
    rt = RunCmd(cmdStr)
    rt = GitCmd(Push)
    GoTo Finally
Catch:
    OutputError "GitInit"
Finally:
    
End Sub
