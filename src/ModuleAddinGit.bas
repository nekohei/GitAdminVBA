Attribute VB_Name = "ModuleAddinGit"
Option Explicit

' アドイン自体のGit管理用モジュール
' このアドイン自体をGitHubにプッシュする機能

Private Const AddinRepoName As String = "GitAdminVBA"
Private Const AddinRootPath As String = "C:¥Users¥1784¥Source¥Repos¥VBA¥GitAdminVBA"

' アドインのGitステータス確認
Public Sub AddinGitStatus(Optional ByVal dummy As Boolean = False)
    On Error GoTo Catch
    Call ChDir(AddinRootPath)
    Dim result As String
    result = RunCmd("git status")
    Debug.Print result
    MsgBox "アドインのGitステータス:" & vbCrLf & vbCrLf & result, vbInformation
    Exit Sub
Catch:
    PrintErr Err, "AddinGitStatus"
End Sub

' アドインの変更をステージング
Public Sub AddinGitStage(Optional ByVal dummy As Boolean = False)
    On Error GoTo Catch
    If MsgBox("アドイン自体の変更をステージしますか?" & vbCrLf & vbCrLf & _
              "この操作により、アドインのソースファイルがエクスポートされ、" & vbCrLf & _
              "変更がステージングエリアに追加されます。", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If

    ' アドインのソースをエクスポート
    Call ExportAddinToSrc

    ' 変更をステージング
    Call ChDir(AddinRootPath)
    Dim result As String
    result = RunCmd("git add .")
    Debug.Print result
    MsgBox "アドインの変更をステージしました。" & vbCrLf & vbCrLf & result, vbInformation
    Exit Sub
Catch:
    PrintErr Err, "AddinGitStage"
End Sub

' アドインの変更をコミット
Public Sub AddinGitCommit(Optional ByVal dummy As Boolean = False)
    On Error GoTo Catch
    Dim commitMessage As String
    commitMessage = InputBox("コミットメッセージを入力してください:", "アドインコミット")
    If commitMessage = "" Then Exit Sub

    If MsgBox("""" & commitMessage & """" & vbCrLf & vbCrLf & _
              "このメッセージでアドインの変更をコミットしますか?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If

    Call ChDir(AddinRootPath)
    Dim result As String
    result = RunCmd("git commit -m """ & commitMessage & """")
    Debug.Print result
    MsgBox "アドインの変更をコミットしました。" & vbCrLf & vbCrLf & result, vbInformation
    Exit Sub
Catch:
    PrintErr Err, "AddinGitCommit"
End Sub

' アドインの変更をプッシュ
Public Sub AddinGitPush(Optional ByVal dummy As Boolean = False)
    On Error GoTo Catch
    If MsgBox("アドインの変更をGitHubにプッシュしますか?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If

    Call ChDir(AddinRootPath)
    Dim result As String
    result = RunCmd("git push origin main")
    Debug.Print result
    MsgBox "アドインの変更をプッシュしました。" & vbCrLf & vbCrLf & result, vbInformation
    Exit Sub
Catch:
    PrintErr Err, "AddinGitPush"
End Sub

' アドインの完全なGitワークフロー実行
Public Sub AddinGitWorkflow(Optional ByVal dummy As Boolean = False)
    On Error GoTo Catch
    Dim commitMessage As String
    commitMessage = InputBox("コミットメッセージを入力してください:", "アドイン完全ワークフロー")
    If commitMessage = "" Then Exit Sub

    If MsgBox("以下の操作を実行します:" & vbCrLf & _
              "1. ソースファイルのエクスポート" & vbCrLf & _
              "2. 変更のステージング" & vbCrLf & _
              "3. コミット: """ & commitMessage & """" & vbCrLf & _
              "4. GitHubへのプッシュ" & vbCrLf & vbCrLf & _
              "実行しますか?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If

    ' ソースファイルエクスポート
    Call ExportAddinToSrc

    ' Git操作を実行
    Call ChDir(AddinRootPath)
    Dim result As String

    ' ステージング
    result = RunCmd("git add .")
    Debug.Print "Stage: " & result

    ' コミット
    result = result & vbCrLf & RunCmd("git commit -m """ & commitMessage & """")
    Debug.Print "Commit: " & result

    ' プッシュ
    result = result & vbCrLf & RunCmd("git push origin main")
    Debug.Print "Push: " & result

    MsgBox "アドインの完全なGitワークフローが完了しました。" & vbCrLf & vbCrLf & result, vbInformation
    Exit Sub
Catch:
    PrintErr Err, "AddinGitWorkflow"
End Sub

' アドインのソースファイルをsrcフォルダにエクスポート
Public Sub ExportAddinToSrc(Optional ByVal dummy As Boolean = False)
    On Error GoTo Catch
    Dim addinPath As String
    addinPath = ThisWorkbook.FullName

    ' アドインのVBProjectを取得
    Dim vbPjt As VBProject: Set vbPjt = ThisWorkbook.VBProject
    Dim vbCmps As VBComponents: Set vbCmps = vbPjt.VBComponents
    Dim vbCmp As VBIDE.VBComponent

    ' srcフォルダの確認・作成
    Dim srcPath As String: srcPath = AddinRootPath & "¥src"
    Call CreateDirIfThereNo(srcPath)

    ' 既存ファイルを削除する前に、現在のプロジェクトに存在するファイル名を収集
    Dim srcDic As New Dictionary

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

        ' ファイルをエクスポート
        vbCmp.Export fPath
        ConvertUTF8 fPath

        ' プロジェクトに存在するモジュールをフォルダから削除する処理
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

    ' プロジェクトに存在しないモジュールファイルを削除
    Dim fso As New FileSystemObject
    Dim fld As Folder
    Set fld = fso.GetFolder(srcPath)
    Dim f As File
    For Each f In fld.Files
        If Not srcDic.Exists(f.Name) Then f.Delete
    Next

    Set vbCmps = Nothing
    Set vbPjt = Nothing
    Set fso = Nothing
    Exit Sub
Catch:
    PrintErr Err, "ExportAddinToSrc"
End Sub

' アドインリポジトリの初期化
Public Sub InitializeAddinRepo(Optional ByVal dummy As Boolean = False)
    On Error GoTo Catch
    If MsgBox("アドインのGitリポジトリを初期化しますか?" & vbCrLf & vbCrLf & _
              "この操作により以下が実行されます:" & vbCrLf & _
              "1. git init" & vbCrLf & _
              "2. ソースファイルのエクスポート" & vbCrLf & _
              "3. 初回コミット" & vbCrLf & _
              "4. mainブランチの設定", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If

    Call ChDir(AddinRootPath)
    Dim result As String

    ' Git初期化
    result = RunCmd("git init")
    Debug.Print "Init: " & result

    ' ソースファイルエクスポート
    Call ExportAddinToSrc

    ' 初回コミット
    result = result & vbCrLf & RunCmd("git add .")
    result = result & vbCrLf & RunCmd("git commit -m ""GitAdminVBAアドインの初回コミット""")
    result = result & vbCrLf & RunCmd("git branch -M main")

    Debug.Print result
    MsgBox "アドインリポジトリの初期化が完了しました。" & vbCrLf & vbCrLf & result, vbInformation
    Exit Sub
Catch:
    PrintErr Err, "InitializeAddinRepo"
End Sub

' アドイン用GitHubリポジトリの作成
Public Sub CreateAddinGitHubRepo(Optional ByVal dummy As Boolean = False)
    On Error GoTo Catch

    ' トークンを取得
    Dim token As String: token = GetTokenFromRegistry()
    If Trim(token) = "" Then
        MsgBox "個人アクセストークンが登録されていません。", vbInformation
        Exit Sub
    End If

    If MsgBox("GitHubにアドイン用リポジトリ """ & AddinRepoName & """ を作成しますか?" & vbCrLf & vbCrLf & _
              "プライベートリポジトリとして作成されます。", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If

    ' HTTPオブジェクトを生成
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")

    ' GitHub APIのURL
    Dim url As String: url = "https://api.github.com/user/repos"
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "token " & token

    ' JSONリクエストボディを作成（アクセス性はprivate）
    Dim jsonBody As String
    jsonBody = "{""name"":""" & AddinRepoName & """, ""private"": true, ""description"": ""VBA Git Admin Add-in for Excel""}"

    ' リクエストを送信
    Call http.send(jsonBody)

    ' 結果を表示
    If http.Status = 201 Then
        Dim json As Object: Set json = JsonConverter.ParseJson(http.responseText)
        Dim repoUrl As String: repoUrl = json("html_url")

        ' リモートリポジトリのURLを設定してプッシュ
        Call ChDir(AddinRootPath)
        Dim result As String
        result = RunCmd("git remote add origin " & repoUrl)
        result = result & vbCrLf & RunCmd("git push -u origin main")

        Debug.Print result
        MsgBox "アドイン用リモートリポジトリが作成され、プッシュされました。" & vbCr & vbCr & _
               repoUrl & vbCr & vbCr & result, vbInformation
    Else
        MsgBox "リモートリポジトリの作成に失敗しました。" & vbCr & vbCr & _
            "Status: " & http.Status & vbCr & http.responseText, vbExclamation
    End If

    Set http = Nothing
    Set json = Nothing
    Exit Sub
Catch:
    PrintErr Err, "CreateAddinGitHubRepo"
End Sub