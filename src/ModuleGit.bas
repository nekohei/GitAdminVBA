Attribute VB_Name = "ModuleGit"
Option Explicit

' ルートの親フォルダ
Public Const parentDir As String = "Source¥Repos¥VBA"
' Git設定ファイルの内容が埋め込まれているモジュール名
Private Const ContentsModuleName As String = "ModuleGitFilesContents"

' GitCmd引数用
Public Enum GitCommand
    Init
    Status
    Stage
    Commit
    Push
End Enum

' エラー出力
Public Sub PrintErr(ByVal errObj As ErrObject, Optional errPlace As String, Optional errNote As String)
    Dim msg As String
    msg = _
        vbLf & _
        "DateTime : " & Format$(Now(), "yyyy-mm-dd hh:nn:ss") & vbLf & _
        "Source   : " & errObj.Source & vbLf & _
        "BookName : " & ActiveWorkbook.Name & vbLf & _
        "Place    : " & errPlace & vbLf & _
        "Note     : " & errNote & vbLf & _
        "Number   : " & errObj.Number & vbLf & _
        "Message  : " & errObj.Description & vbLf
    Debug.Print msg
End Sub

Public Sub CreateNewRepository()
    
    ' リポジトリ名を設定する
    Dim xBook As Workbook: Set xBook = ActiveWorkbook
    Dim repoName As String: repoName = SetAndThenGetReposName(xBook.Name)
    If repoName = "" Then Exit Sub
    
    ' ローカルリポジトリフォルダ作成
    Dim repoDir As String: repoDir = GetRootDir(xBook.Name)
    If repoDir = "" Then Exit Sub
    Call CreateDirIfThereNo(repoDir)

    ' ローカルリポジトリフォルダ内のサブフォルダ作成
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(repoDir & "¥.vscode") Then fso.CreateFolder repoDir & "¥.vscode"
    If Not fso.FolderExists(repoDir & "¥bin") Then fso.CreateFolder repoDir & "¥bin"
    If Not fso.FolderExists(repoDir & "¥bin¥old") Then fso.CreateFolder repoDir & "¥bin¥old"
    If Not fso.FolderExists(repoDir & "¥src") Then fso.CreateFolder repoDir & "¥src"
    Dim bookName As String: bookName = GetShortBookName(xBook.Name)
    Dim srcDir As String: srcDir = repoDir & "¥src"
    If Not fso.FolderExists(srcDir) Then fso.CreateFolder srcDir

    ' Git設定ファイル作成
    Call GenerateGitFiles(repoDir, repoDir)
    
    ' ブックを保存してbinフォルダにコピー
    xBook.Save
    Call fso.CopyFile(xBook.FullName, repoDir & "¥bin¥" & xBook.Name, True)
    Set fso = Nothing
    
    ' srcフォルダにCodeModuleをExport
    Call Decombine(xBook.Name)
    
    '登録済みアカウントごとにリモートリポジトリを作成
    Dim accounts() As String: accounts = GetAccountList()
    Dim successCount As Long: successCount = 0
    Dim acIdx As Long
    For acIdx = 0 To UBound(accounts)
        If CreateRemoteRepos(bookName, repoName, accounts(acIdx)) Then
            successCount = successCount + 1
        End If
    Next acIdx
    If successCount > 0 Then
        Call GitCmd(Init)
        MsgBox bookName & " 用のリポジトリの準備ができました。", vbInformation
    End If
    
    Set xBook = Nothing
End Sub

Public Function CreateRemoteRepos(ByVal bookName As String, ByVal repoName As String, ByVal accountName As String) As Boolean
    On Error GoTo Catch
    Dim rt As Boolean: rt = False
    
    ' トークンを取得
    Dim token As String: token = GetTokenFromRegistry(accountName)
    If Trim(token) = "" Then
        MsgBox "個人用アクセストークンを登録してください。", vbInformation
        CreateRemoteRepos = rt
        Exit Function
    End If

    ' HTTPオブジェクトを生成
    Dim http     As Object: Set http = CreateObject("MSXML2.XMLHTTP")

    ' GitHub APIのURL
    Dim url      As String: url = "https://api.github.com/user/repos"
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "token " & token

    ' JSONリクエストボディを作成（アクセス修飾子はprivate）
    Dim jsonBody As String: jsonBody = "{""name"":""" & repoName & """, ""private"": true}"

    ' リクエストを送信
    Call http.send(jsonBody)

    ' 結果を表示
    If http.Status = 201 Then
        Dim json    As Object: Set json = JsonConverter.ParseJson(http.responseText)
        Dim repoUrl As String: repoUrl = json("html_url")
        MsgBox accountName & " のリモートリポジトリが作成されました。" & vbCr & vbCr & repoUrl, vbInformation
        rt = True
    Else
        MsgBox "リモートリポジトリの作成に失敗しました。" & vbCr & vbCr & _
            "Status: " & http.Status & vbCr & http.responseText, vbExclamation
        rt = False
    End If
    GoTo Finally
Catch:
    MsgBox Err.Description, vbExclamation
    rt = False
Finally:
    Set http = Nothing
    Set json = Nothing
    CreateRemoteRepos = rt
End Function

' 引数のフォルダパスが存在しない場合に作る
Public Sub CreateDirIfThereNo(ByVal dirPath As String)

    Dim fso  As Object:  Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dirs As Variant: dirs = Split(dirPath, "¥")

    Dim i As Integer, dr As String
    For i = 0 To UBound(dirs)
        dr = dr & dirs(i) & "¥"
        ' 共有ネットワークパスを考慮
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
Public Sub GenerateUTF8(ByVal txt As String, ByVal FilePath As String, Optional ByVal withBom As Boolean = True)

    Dim binData As Variant, sm1 As Object, sm2 As Object
    
    Set sm1 = CreateObject("ADODB.Stream")
    With sm1
        .Type = 2                 ' 文字列
        .Charset = "utf-8"        ' 文字コード指定
        .Open                     ' 開く
        .WriteText txt            ' 文字列を書き込む
        .Position = 0             ' streamの先頭に移動
        .Type = 1                 ' バイナリー
        If withBom Then
            .Position = 3         ' streamの先頭から3バイトをスキップ
        End If
        binData = .Read           ' バイナリー取得
        .Close: Set sm1 = Nothing ' 閉じて解放
    End With
    
    Set sm2 = CreateObject("ADODB.Stream")
    With sm2
        .Type = 1                 ' バイナリー
        .Open                     ' 開く
        .Write binData            ' バイナリーデータを書き込む
        .SaveToFile FilePath, 2   ' 保存
        .Close: Set sm2 = Nothing ' 閉じて解放
    End With

End Sub

' Git設定ファイルの作成
Public Sub GenerateGitFiles(ByVal rootDir As String, ByVal srcDir As String)
    
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
 
' コマンドプロンプトでcmd引数を実行
' 標準出力は文字化けする
Private Function RunCmd(ByVal cmd As String, Optional ByVal showInt As Integer = 0, Optional ByVal toWait As Boolean = True) As String
    On Error GoTo Catch
    Dim fso     As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim msgPath As String: msgPath = Environ$("temp") & "¥gitTmp.log"
    Dim errPath As String: errPath = Environ$("temp") & "¥gitErr.log"
    Dim wsh     As Object: Set wsh = CreateObject("WScript.Shell")
    Dim rt      As Long:   rt = wsh.Run("cmd /c " & cmd & " > " & msgPath & " 2> " & errPath, showInt, toWait)
    
    Dim msg As String
    If rt = 0 Then
        msg = "(正常終了 - " & CurDir & ") " & vbCr & cmd
    Else
        msg = "(異常終了 - " & CurDir & ") " & vbCr & cmd
    End If
    
    Dim msgStream As Object: Set msgStream = CreateObject("ADODB.Stream")
    msgStream.Type = 2
    msgStream.Charset = "utf-8"
    msgStream.Open
    msgStream.LoadFromFile msgPath
    Dim msgText As String
    msgText = msgStream.ReadText
    msgStream.Close: Set msgStream = Nothing
    
    Dim errStream As Object: Set errStream = CreateObject("ADODB.Stream")
    errStream.Type = 2
    errStream.Charset = "utf-8"
    errStream.Open
    errStream.LoadFromFile errPath
    Dim errText As String
    errText = errStream.ReadText
    errStream.Close: Set errStream = Nothing
    
    Select Case True
    Case Trim(msgText) = "" And Trim(errText) = ""
        ' 何もしない
    Case Trim(msgText) <> "" And Trim(errText) = ""
        msg = msg & vbCr & msgText
    Case Trim(msgText) = "" And Trim(errText) <> ""
        msg = msg & vbCr & errText
    Case Trim(msgText) <> "" And Trim(errText) <> ""
        msg = msg & vbCr & msgText & vbCr & errText
    End Select
    
    GoTo Finally
Catch:
    PrintErr Err, "RunCmd"
Finally:
    If fso.FileExists(msgPath) Then fso.DeleteFile msgPath
    If fso.FileExists(errPath) Then fso.DeleteFile errPath
    Set fso = Nothing
    Set wsh = Nothing
    RunCmd = msg
End Function

' Gitコマンドを実行
Public Function GitCmd(ByVal cmd As GitCommand, Optional ByVal arg As String = Empty, Optional ByVal isPowerShell As Boolean = False) As Integer
    On Error GoTo Catch
    Dim xBook As Workbook: Set xBook = ActiveWorkbook
    Dim rootDir As String: rootDir = GetRootDir(xBook.Name)
    If rootDir = "" Then
        MsgBox """" & xBook.Name & """" & vbLf & vbLf & "リポジトリ名が登録されていません。", vbInformation
        GoTo Finally
    End If
    
    Call ChDir(rootDir)
    
    Dim rt As String
    Select Case cmd
    Case Init
        Dim bookName As String: bookName = GetShortBookName(xBook.Name)
        Dim initRepoName As String: initRepoName = GetSetting("Excel", bookName, "RepositoryName")
        Dim initAccounts() As String: initAccounts = GetPushAccounts(xBook.Name)
        If initRepoName = "" Or UBound(initAccounts) < 0 Then
            MsgBox "リモートリポジトリを作成してください。", vbInformation
            GoTo Finally
        End If
        rt = RunCmd("git init")
        rt = rt & vbCr & RunCmd("git add .")
        rt = rt & vbCr & RunCmd("git commit -m ""リポジトリ開始""")
        rt = rt & vbCr & RunCmd("git branch -M main")
        Dim initOriginUrl As String
        initOriginUrl = "https://github.com/" & initAccounts(0) & "/" & initRepoName & ".git"
        rt = rt & vbCr & RunCmd("git remote add origin " & initOriginUrl)
        Dim initIdx As Long
        For initIdx = 0 To UBound(initAccounts)
            Dim initPat As String: initPat = GetTokenFromRegistry(initAccounts(initIdx))
            If initPat <> "" Then
                Dim initPushUrl As String
                initPushUrl = "https://" & initPat & "@github.com/" & initAccounts(initIdx) & "/" & initRepoName & ".git"
                rt = rt & vbCr & RunCmd("git push " & initPushUrl & " main")
            End If
        Next initIdx
    Case Status
        rt = RunCmd("git status")
    Case Stage
        If MsgBox(xBook.Name & " の変更をステージします。" & vbLf & vbLf & _
                  xBook.Name & " の保存とエクスポートを伴います。", vbInformation + vbOKCancel) = vbOK Then
            Application.DisplayAlerts = False
            If xBook.path = rootDir & "¥bin" Then
                MsgBox "binフォルダ内の" & xBook.Name & "を開いたままステージ出来ません。" & vbLf & _
                       "ステージはキャンセルされました。", vbInformation
                GoTo Finally
            Else
                'xBook.Save
                Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
                Call fso.CopyFile(xBook.FullName, rootDir & "¥bin¥" & xBook.Name, True)
                Call Decombine(xBook.Name)
                rt = RunCmd("git add .")
                Set fso = Nothing
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
        If arg = Empty Then mBranch = "main" Else mBranch = arg
        Dim pushBookName As String: pushBookName = GetShortBookName(xBook.Name)
        Dim pushRepoName As String: pushRepoName = GetSetting("Excel", pushBookName, "RepositoryName")
        If pushRepoName = "" Then
            MsgBox "リポジトリが登録されていません。", vbInformation
            GoTo Finally
        End If
        Dim pushAccounts() As String: pushAccounts = GetPushAccounts(xBook.Name)
        Dim pushIdx As Long
        For pushIdx = 0 To UBound(pushAccounts)
            Dim pushPat As String: pushPat = GetTokenFromRegistry(pushAccounts(pushIdx))
            If pushPat <> "" Then
                Dim pushUrl As String
                pushUrl = "https://" & pushPat & "@github.com/" & pushAccounts(pushIdx) & "/" & pushRepoName & ".git"
                rt = rt & vbCr & RunCmd("git push " & pushUrl & " " & mBranch)
            End If
        Next pushIdx
    End Select
    If Right(rt, 1) <> vbLf Then rt = rt & vbLf
    Debug.Print rt
    GoTo Finally
Catch:
    PrintErr Err, "GitCmd"
Finally:
    Application.DisplayAlerts = True
    Set xBook = Nothing
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
Private Function GetRootDir(ByVal bookName As String) As String
    bookName = GetShortBookName(bookName)
    Dim repoName As String: repoName = GetSetting("Excel", bookName, "RepositoryName")
    If repoName = "" Then
        GetRootDir = ""
        Exit Function
    End If
    GetRootDir = Environ$("USERPROFILE") & "¥" & parentDir & "¥" & repoName
End Function

' srcフォルダを指定してActiveWorkbookをExport
Public Sub Decombine(ByVal bookName As String, Optional ByVal includeBookName As Boolean = False)
    Dim rootDir As String: rootDir = GetRootDir(bookName)
    If rootDir = "" Then Exit Sub
    Dim srcPath As String
    If includeBookName Then
        srcPath = rootDir & "¥src¥" & GetShortBookName(bookName)
    Else
        srcPath = rootDir & "¥src"
    End If
    Call CreateDirIfThereNo(srcPath)
    Dim srcDic As New Dictionary
    Dim vbPjt As VBProject: Set vbPjt = Workbooks(bookName).VBProject
    Dim vbCmps As VBComponents: Set vbCmps = vbPjt.VBComponents
    Dim vbCmp As VBIDE.VBComponent
    For Each vbCmp In vbCmps
        Dim fName As String: fName = ""
        Dim fPath As String: fPath = ""
        Select Case vbCmp.Type
            Case vbext_ct_StdModule
                fName = vbCmp.Name & ".bas"
                fPath = srcPath & "¥" & fName
            Case vbext_ct_MSForm
                fName = vbCmp.Name & ".frm"
                fPath = srcPath & "¥" & fName
            Case vbext_ct_ClassModule
                fName = vbCmp.Name & ".cls"
                fPath = srcPath & "¥" & fName
            Case vbext_ct_Document
                fName = vbCmp.Name & ".dcm"
                fPath = srcPath & "¥" & fName
            Case Else
                GoTo Continue
        End Select
        vbCmp.Export fPath
        ConvertUTF8 fPath
        ' プロジェクトに無いモジュールをフォルダから削除する準備
        If Not srcDic.Exists(fName) Then
            Call srcDic.Add(fName, fPath)
            If Right(fName, 3) = "frm" Then
                fName = vbCmp.Name & ".frx"
                fPath = srcPath & "¥" & fName
                If Not srcDic.Exists(fName) Then _
                    Call srcDic.Add(fName, fPath)
            End If
        End If
Continue:
        Set vbCmp = Nothing
    Next
    Set vbCmps = Nothing
    Set vbPjt = Nothing
    ' プロジェクトに無いモジュールファイルを削除
    Dim fso As New FileSystemObject
    Dim fld As Folder
    Set fld = fso.GetFolder(srcPath)
    Dim f As File
    For Each f In fld.Files
        If Not srcDic.Exists(f.Name) Then f.Delete
    Next
End Sub

Private Sub ConvertUTF8(ByVal srcPath As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    ' テキストファイルを開く
    Dim sm As Object: Set sm = fso.OpenTextFile(srcPath)
    ' ファイルの内容を読み込む
    Dim txt As String: txt = sm.ReadAll
    Set sm = Nothing
    Set fso = Nothing
    ' UTF-8でテキストファイルに保存
    Call GenerateUTF8(txt, srcPath)
End Sub

Private Sub DecombineEx()
    ' 新しいExcelインスタンスを作成
    Dim newExcel As Excel.Application
    Set newExcel = CreateObject("Excel.Application")
    ' ブックを読み取り専用で開く
    Dim bookName As String: bookName = ActiveWorkbook.FullName
    Dim xBook As Workbook
    Set xBook = newExcel.Workbooks.Open(bookName, ReadOnly:=True)
    
    Dim rootDir As String: rootDir = GetRootDir(bookName)
    If rootDir = "" Then Exit Sub
    Dim srcPath As String
    srcPath = rootDir & "¥src"
    Call CreateDirIfThereNo(srcPath)
    Dim vbPjt As VBProject: Set vbPjt = xBook.VBProject
    Dim vbCmps As VBComponents: Set vbCmps = vbPjt.VBComponents
    Dim vbCmp As VBIDE.VBComponent
    For Each vbCmp In vbCmps
        Dim fPath As String: fPath = ""
        Select Case vbCmp.Type
            Case vbext_ct_StdModule
                fPath = srcPath & "¥" & vbCmp.Name & ".bas"
            Case vbext_ct_MSForm
                fPath = srcPath & "¥" & vbCmp.Name & ".frm"
            Case vbext_ct_ClassModule
                fPath = srcPath & "¥" & vbCmp.Name & ".cls"
            Case vbext_ct_Document
                fPath = srcPath & "¥" & vbCmp.Name & ".dcm"
            Case Else
                GoTo Continue
        End Select
        vbCmp.Export fPath
        ConvertUTF8 fPath
Continue:
        Set vbCmp = Nothing
    Next
    xBook.Close False
    newExcel.Quit
    Set vbCmps = Nothing
    Set vbPjt = Nothing
End Sub

' リポジトリ名をレジストリに記録
Public Function SetAndThenGetReposName(ByVal bookName As String) As String
    bookName = GetShortBookName(bookName)
    Dim repoName As String: repoName = GetSetting("Excel", bookName, "RepositoryName")
    If repoName = "" Then
        repoName = InputBox("リポジトリ名を英字で入力してください。")
        If repoName = "" Then
            SetAndThenGetReposName = ""
            Exit Function
        End If
        If Not IsValidRepoName(repoName) Then
            MsgBox "リポジトリ名は無効です。" & vbCr & vbCr & _
                "リポジトリ名は英字で始まり、小文字、数字、ハイフン、アンダースコア、" & vbCr & _
                "ピリオドを含めることができ、最大256文字までです。" & vbCr & _
                "連続するハイフン、アンダースコアは使用できません。", vbInformation
            SetAndThenGetReposName = ""
            Exit Function
        End If
        Call SaveSetting("Excel", bookName, "RepositoryName", repoName)
    End If
    SetAndThenGetReposName = repoName
End Function

' 最初のアンダースコアから最後のドットまでの間の文字列を消去する
Private Function GetShortBookName(ByVal bookName As String) As String
    
    ' アンスコの位置
    Dim unsPos As Integer: unsPos = InStr(bookName, "_")
    ' ドットの位置
    Dim dotPos As Integer: dotPos = InStrRev(bookName, ".")

    If unsPos = 0 Or dotPos = 0 Then
        ' アンスコまたはドットが見つからない場合、元のファイル名を返す
        GetShortBookName = bookName
    Else
        ' アンスコの前の部分と、最後のドットの後の部分を結合
        GetShortBookName = Left(bookName, unsPos - 1) & Mid(bookName, dotPos)
    End If
End Function

' レジストリにトークンを登録
Public Sub RegisterToken()
    On Error GoTo Catch
    Dim accountName As String
    accountName = InputBox("GitHubのアカウント名を入力してください。")
    If accountName = "" Then Exit Sub
    Dim keyStr As String
    keyStr = InputBox("GitHubの個人アクセストークンを入力してください。")
    If keyStr = "" Then Exit Sub
    Call SaveSetting("GitHub", accountName, "Classic", keyStr)
    On Error Resume Next
    DeleteSetting "GitHub", "Token"
    On Error GoTo 0
    MsgBox accountName & " のトークンを登録しました。", vbInformation
    Exit Sub
Catch:
    MsgBox Err.Description, vbExclamation
End Sub

' レジストリからトークンを得る
Public Function GetTokenFromRegistry(ByVal accountName As String) As String
    GetTokenFromRegistry = GetSetting("GitHub", accountName, "Classic")
End Function

'登録済みアカウント一覧を返す（reg query によるレジストリキー列挙）
Public Function GetAccountList() As String()
    On Error GoTo Catch
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim oExec As Object
    Set oExec = wsh.Exec("cmd /c reg query ""HKCU¥Software¥VB and VBA Program Settings¥GitHub""")
    Dim output As String
    Do While Not oExec.StdOut.AtEndOfStream
        output = output & oExec.StdOut.ReadLine() & vbLf
    Loop
    Set wsh = Nothing

    Dim baseKey As String
    baseKey = "HKEY_CURRENT_USER¥Software¥VB and VBA Program Settings¥GitHub¥"
    Dim lines() As String: lines = Split(output, vbLf)
    Dim result() As String
    ReDim result(UBound(lines))
    Dim idx As Long: idx = 0
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim line As String: line = Trim(lines(i))
        If Left(line, Len(baseKey)) = baseKey Then
            result(idx) = Mid(line, Len(baseKey) + 1)
            idx = idx + 1
        End If
    Next i
    If idx = 0 Then
        GetAccountList = Array()
    Else
        ReDim Preserve result(idx - 1)
        GetAccountList = result
    End If
    Exit Function
Catch:
    GetAccountList = Array()
End Function

'ブックのプッシュ先アカウントリストをレジストリから取得（未設定時は全アカウント）
Public Function GetPushAccounts(ByVal bookName As String) As String()
    bookName = GetShortBookName(bookName)
    Dim setting As String
    setting = GetSetting("Excel", bookName, "PushAccounts")
    If Trim(setting) = "" Then
        '未設定の場合は全アカウント（後方互換）
        GetPushAccounts = GetAccountList()
    Else
        Dim parts() As String: parts = Split(setting, ",")
        Dim i As Long
        For i = 0 To UBound(parts)
            parts(i) = Trim(parts(i))
        Next i
        GetPushAccounts = parts
    End If
End Function

'ブックのプッシュ先アカウントをレジストリに登録
Public Sub RegisterPushAccounts()
    On Error GoTo Catch
    Dim xBook As Workbook: Set xBook = ActiveWorkbook
    Dim bookName As String: bookName = GetShortBookName(xBook.Name)
    Dim allAccounts() As String: allAccounts = GetAccountList()
    If UBound(allAccounts) < 0 Then
        MsgBox "登録済みのGitHubアカウントがありません。", vbInformation
        Exit Sub
    End If
    Dim current As String: current = GetSetting("Excel", bookName, "PushAccounts")
    Dim prompt As String
    prompt = "プッシュ先アカウントをカンマ区切りで入力してください。" & vbLf & vbLf & _
             "登録済みアカウント: " & Join(allAccounts, ", ") & vbLf & vbLf & _
             "現在の設定: " & IIf(current = "", "（未設定 = 全アカウント）", current)
    Dim ans As String: ans = InputBox(prompt, "プッシュ先アカウント設定", current)
    If StrPtr(ans) = 0 Then Exit Sub  'キャンセル
    Call SaveSetting("Excel", bookName, "PushAccounts", ans)
    If Trim(ans) = "" Then
        MsgBox "プッシュ先設定を解除しました（全アカウントにプッシュします）。", vbInformation
    Else
        MsgBox bookName & " のプッシュ先: " & ans, vbInformation
    End If
    Exit Sub
Catch:
    MsgBox Err.Description, vbExclamation
End Sub


' レジストリからトークン用のキーを削除
Public Sub DeleteToken()
    On Error GoTo Catch
    Dim accounts() As String: accounts = GetAccountList()
    If UBound(accounts) < 0 Then
        MsgBox "登録されているアカウントはありません。", vbInformation
        Exit Sub
    End If
    Dim accountName As String
    accountName = InputBox("削除するアカウント名を入力してください。" & vbLf & vbLf & "登録済み：" & Join(accounts, ", "))
    If accountName = "" Then Exit Sub
    DeleteSetting "GitHub", accountName
    MsgBox accountName & " のトークンを削除しました。", vbInformation
    Exit Sub
Catch:
    MsgBox Err.Description, vbExclamation
End Sub

' GitHubのリポジトリ名の有効性をチェックする
Function IsValidRepoName(ByVal repoName As String) As Boolean
    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")

    ' リポジトリ名は英字で始まり、指定された文字のみ含む、最大256文字
    With regEx
        .Pattern = "^[a-zA-Z][-a-zA-Z0-9_.]*$"
        .IgnoreCase = False
        .Global = False
    End With

    ' リポジトリ名が空、長すぎる、または正規表現に一致しない場合は無効
    If Len(repoName) = 0 Or Len(repoName) > 256 Or Not regEx.Test(repoName) Then
        IsValidRepoName = False
    Else
        ' 連続するハイフン、アンダースコアをチェック
        If InStr(repoName, "--") > 0 Or InStr(repoName, "__") > 0 Then
            IsValidRepoName = False
        Else
            IsValidRepoName = True
        End If
    End If
    
    Set regEx = Nothing
End Function

Function GetRepositoryURL(ByVal repoPath As String) As String
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim cmd As String: cmd = "cmd /c cd /d """ & repoPath & """ & git remote get-url origin"
    Dim wEx As Object: Set wEx = wsh.Exec(cmd)

    Dim rt As String
    Do While Not wEx.StdOut.AtEndOfStream
        rt = rt & wEx.StdOut.ReadLine() & vbNewLine
    Loop
    Dim rtStr() As String: rtStr = Split(rt, vbNewLine)
    GetRepositoryURL = rtStr(0)
End Function


' srcフォルダを指定してActiveWorkbookをExport
Public Sub DecombinePortable(ByVal bookName As String, Optional rootDir As String, Optional ByVal includeBookName As Boolean = False)
    If rootDir = "" Then rootDir = InputBox("リポジトリのパスを指定")
    If rootDir = "" Then Exit Sub
    Dim srcPath As String
    If includeBookName Then
        srcPath = rootDir & "¥src¥" & GetShortBookName(bookName)
    Else
        srcPath = rootDir & "¥src"
    End If
    Call CreateDirIfThereNo(srcPath)
    Dim srcDic As New Dictionary
    Dim vbPjt As VBProject: Set vbPjt = Workbooks(bookName).VBProject
    Dim vbCmps As VBComponents: Set vbCmps = vbPjt.VBComponents
    Dim vbCmp As VBIDE.VBComponent
    For Each vbCmp In vbCmps
        Dim fName As String: fName = ""
        Dim fPath As String: fPath = ""
        Select Case vbCmp.Type
            Case vbext_ct_StdModule
                fName = vbCmp.Name & ".bas"
                fPath = srcPath & "¥" & fName
            Case vbext_ct_MSForm
                fName = vbCmp.Name & ".frm"
                fPath = srcPath & "¥" & fName
            Case vbext_ct_ClassModule
                fName = vbCmp.Name & ".cls"
                fPath = srcPath & "¥" & fName
            Case vbext_ct_Document
                fName = vbCmp.Name & ".dcm"
                fPath = srcPath & "¥" & fName
            Case Else
                GoTo Continue
        End Select
        vbCmp.Export fPath
        ConvertUTF8 fPath
        ' プロジェクトに無いモジュールをフォルダから削除する準備
        If Not srcDic.Exists(fName) Then
            Call srcDic.Add(fName, fPath)
            If Right(fName, 3) = "frm" Then
                fName = vbCmp.Name & ".frx"
                fPath = srcPath & "¥" & fName
                If Not srcDic.Exists(fName) Then _
                    Call srcDic.Add(fName, fPath)
            End If
        End If
Continue:
        Set vbCmp = Nothing
    Next
    Set vbCmps = Nothing
    Set vbPjt = Nothing
    ' プロジェクトに無いモジュールファイルを削除
    Dim fso As New FileSystemObject
    Dim fld As Folder
    Set fld = fso.GetFolder(srcPath)
    Dim f As File
    For Each f In fld.Files
        If Not srcDic.Exists(f.Name) Then f.Delete
    Next
    MsgBox "展開が完了しました。", vbInformation
End Sub

