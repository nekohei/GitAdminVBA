# patch_module_git.py - バイナリ方式でModuleGit.basを改修
import sys

SRC = 'c:/Users/1784/claude/VBA/Excel/GitAdminVBA/src/ModuleGit.bas'

def j(text):
    """日本語文字列をCP932バイト列に変換"""
    return text.encode('cp932')

with open(SRC, 'rb') as f:
    src = f.read()

changes = []

# ============================================================
# 1. GetTokenFromRegistry: パラメータ追加
# ============================================================
changes.append((
    b'Public Function GetTokenFromRegistry() As String\r\n'
    b'    GetTokenFromRegistry = GetSetting("GitHub", "Token", "Classic")\r\n'
    b'End Function',

    b'Public Function GetTokenFromRegistry(ByVal accountName As String) As String\r\n'
    b'    GetTokenFromRegistry = GetSetting("GitHub", accountName, "Classic")\r\n'
    b'End Function\r\n'
    b'\r\n'
    + j('\'登録済みアカウント一覧を返す（WMI によるレジストリキー列挙）\r\n')
    + b'Public Function GetAccountList() As String()\r\n'
    b'    Dim oReg As Object\r\n'
    b'    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\default:StdRegProv")\r\n'
    b'    Dim arrSubKeys() As String\r\n'
    b'    Const HKCU As Long = &H80000001\r\n'
    b'    oReg.EnumKey HKCU, "Software\\VB and VBA Program Settings\\GitHub", arrSubKeys\r\n'
    b'    Set oReg = Nothing\r\n'
    b'    If IsNull(arrSubKeys) Or Not IsArray(arrSubKeys) Then\r\n'
    b'        GetAccountList = Array()\r\n'
    b'    Else\r\n'
    b'        GetAccountList = arrSubKeys\r\n'
    b'    End If\r\n'
    b'End Function'
))

# ============================================================
# 2. RegisterToken: アカウント名を入力させる形式に変更
# ============================================================
changes.append((
    b'Public Sub RegisterToken()\r\n'
    b'    On Error GoTo Catch\r\n'
    b'    Dim keyStr As String\r\n'
    + j('    keyStr = InputBox("GitHubの個人アクセストークンを入力してください。")\r\n')
    + b'    If keyStr = "" Then Exit Sub\r\n'
    b'    Call SaveSetting("GitHub", "Token", "Classic", keyStr)\r\n'
    + j('    MsgBox "GitHubの個人アクセストークンを登録しました。", vbInformation\r\n')
    + b'    Exit Sub\r\n'
    b'Catch:\r\n'
    b'    MsgBox Err.Description, vbExclamation\r\n'
    b'End Sub',

    b'Public Sub RegisterToken()\r\n'
    b'    On Error GoTo Catch\r\n'
    b'    Dim accountName As String\r\n'
    + j('    accountName = InputBox("GitHubのアカウント名を入力してください。")\r\n')
    + b'    If accountName = "" Then Exit Sub\r\n'
    b'    Dim keyStr As String\r\n'
    + j('    keyStr = InputBox("GitHubの個人アクセストークンを入力してください。")\r\n')
    + b'    If keyStr = "" Then Exit Sub\r\n'
    b'    Call SaveSetting("GitHub", accountName, "Classic", keyStr)\r\n'
    b'    On Error Resume Next\r\n'
    b'    DeleteSetting "GitHub", "Token"\r\n'
    b'    On Error GoTo 0\r\n'
    + j('    MsgBox accountName & " のトークンを登録しました。", vbInformation\r\n')
    + b'    Exit Sub\r\n'
    b'Catch:\r\n'
    b'    MsgBox Err.Description, vbExclamation\r\n'
    b'End Sub'
))

# ============================================================
# 3. DeleteToken: アカウント一覧から選択して削除
# ============================================================
changes.append((
    b'Public Sub DeleteToken()\r\n'
    b'    Call DeleteSetting("GitHub", "Token", "Classic")\r\n'
    b'End Sub',

    b'Public Sub DeleteToken()\r\n'
    b'    On Error GoTo Catch\r\n'
    b'    Dim accounts() As String: accounts = GetAccountList()\r\n'
    b'    If UBound(accounts) < 0 Then\r\n'
    + j('        MsgBox "登録されているアカウントはありません。", vbInformation\r\n')
    + b'        Exit Sub\r\n'
    b'    End If\r\n'
    b'    Dim accountName As String\r\n'
    + j('    accountName = InputBox("削除するアカウント名を入力してください。" & vbLf & vbLf & "登録済み：" & Join(accounts, ", "))\r\n')
    + b'    If accountName = "" Then Exit Sub\r\n'
    b'    DeleteSetting "GitHub", accountName\r\n'
    + j('    MsgBox accountName & " のトークンを削除しました。", vbInformation\r\n')
    + b'    Exit Sub\r\n'
    b'Catch:\r\n'
    b'    MsgBox Err.Description, vbExclamation\r\n'
    b'End Sub'
))

# ============================================================
# 4. CreateRemoteRepos: accountName パラメータ追加
# ============================================================
changes.append((
    b'Public Function CreateRemoteRepos(ByVal bookName As String, ByVal repoName As String) As Boolean',
    b'Public Function CreateRemoteRepos(ByVal bookName As String, ByVal repoName As String, ByVal accountName As String) As Boolean'
))
changes.append((
    b'    Dim token As String: token = GetTokenFromRegistry()\r\n',
    b'    Dim token As String: token = GetTokenFromRegistry(accountName)\r\n'
))
# GitCmd(Init)呼び出し・RepositoryURL保存を削除し、アカウント名付きメッセージに変更
changes.append((
    b'        Dim json    As Object: Set json = JsonConverter.ParseJson(http.responseText)\r\n'
    b'        Dim repoUrl As String: repoUrl = json("html_url")\r\n'
    b'        Call SaveSetting("Excel", bookName, "RepositoryURL", repoUrl)\r\n'
    b'        Call GitCmd(Init)\r\n'
    + j('        MsgBox "リモートリポジトリが作成されました。" & vbCr & vbCr & repoUrl, vbInformation\r\n'),

    b'        Dim json    As Object: Set json = JsonConverter.ParseJson(http.responseText)\r\n'
    b'        Dim repoUrl As String: repoUrl = json("html_url")\r\n'
    + j('        MsgBox accountName & " のリモートリポジトリが作成されました。" & vbCr & vbCr & repoUrl, vbInformation\r\n')
))

# ============================================================
# 5. CreateNewRepository: 全アカウントにループ対応
# ============================================================
changes.append((
    j('    \'リモートリポジトリの作成\r\n')
    + b'    If CreateRemoteRepos(bookName, repoName) Then\r\n'
    + j('        MsgBox bookName & " 用のリポジトリの初期化ができました。", vbInformation\r\n')
    + b'    End If',

    j('    \'登録済みアカウントごとにリモートリポジトリを作成\r\n')
    + b'    Dim accounts() As String: accounts = GetAccountList()\r\n'
    b'    Dim successCount As Long: successCount = 0\r\n'
    b'    Dim acIdx As Long\r\n'
    b'    For acIdx = 0 To UBound(accounts)\r\n'
    b'        If CreateRemoteRepos(bookName, repoName, accounts(acIdx)) Then\r\n'
    b'            successCount = successCount + 1\r\n'
    b'        End If\r\n'
    b'    Next acIdx\r\n'
    b'    If successCount > 0 Then\r\n'
    b'        Call GitCmd(Init)\r\n'
    + j('        MsgBox bookName & " 用のリポジトリの初期化ができました。", vbInformation\r\n')
    + b'    End If'
))

# ============================================================
# 6. GitCmd Case Init: RepositoryURL 依存をなくしアカウントループに
# ============================================================
changes.append((
    b'    Case Init\r\n'
    b'        Dim bookName As String: bookName = GetShortBookName(xBook.Name)\r\n'
    b'        Dim repoUrl  As String: repoUrl = GetSetting("Excel", bookName, "RepositoryURL")\r\n'
    b'        If repoUrl = "" Then\r\n'
    + j('            MsgBox "リモートリポジトリを作成してください。", vbInformation\r\n')
    + b'            GoTo Finally\r\n'
    b'        End If\r\n'
    b'        rt = RunCmd("git init")\r\n'
    b'        rt = rt & vbCr & RunCmd("git add .")\r\n'
    + j('        rt = rt & vbCr & RunCmd("git commit -m ""リポジトリ開始""")\r\n')
    + b'        rt = rt & vbCr & RunCmd("git branch -M main")\r\n'
    b'        rt = rt & vbCr & RunCmd("git remote add origin " & repoUrl)\r\n'
    b'        rt = rt & vbCr & RunCmd("git push -u origin main")\r\n',

    b'    Case Init\r\n'
    b'        Dim bookName As String: bookName = GetShortBookName(xBook.Name)\r\n'
    b'        Dim initRepoName As String: initRepoName = GetSetting("Excel", bookName, "RepositoryName")\r\n'
    b'        Dim initAccounts() As String: initAccounts = GetAccountList()\r\n'
    b'        If initRepoName = "" Or UBound(initAccounts) < 0 Then\r\n'
    + j('            MsgBox "リモートリポジトリを作成してください。", vbInformation\r\n')
    + b'            GoTo Finally\r\n'
    b'        End If\r\n'
    b'        rt = RunCmd("git init")\r\n'
    b'        rt = rt & vbCr & RunCmd("git add .")\r\n'
    + j('        rt = rt & vbCr & RunCmd("git commit -m ""リポジトリ開始""")\r\n')
    + b'        rt = rt & vbCr & RunCmd("git branch -M main")\r\n'
    b'        Dim initOriginUrl As String\r\n'
    b'        initOriginUrl = "https://github.com/" & initAccounts(0) & "/" & initRepoName & ".git"\r\n'
    b'        rt = rt & vbCr & RunCmd("git remote add origin " & initOriginUrl)\r\n'
    b'        Dim initIdx As Long\r\n'
    b'        For initIdx = 0 To UBound(initAccounts)\r\n'
    b'            Dim initPat As String: initPat = GetTokenFromRegistry(initAccounts(initIdx))\r\n'
    b'            If initPat <> "" Then\r\n'
    b'                Dim initPushUrl As String\r\n'
    b'                initPushUrl = "https://" & initPat & "@github.com/" & initAccounts(initIdx) & "/" & initRepoName & ".git"\r\n'
    b'                rt = rt & vbCr & RunCmd("git push " & initPushUrl & " main")\r\n'
    b'            End If\r\n'
    b'        Next initIdx\r\n'
))

# ============================================================
# 7. GitCmd Case Push: 全アカウントへ push
# ============================================================
changes.append((
    b'    Case Push\r\n'
    b'        Dim mBranch As String\r\n'
    b'        If arg = Empty Then mBranch = "main" Else mBranch = arg\r\n'
    b'        rt = RunCmd("git push origin " & mBranch)\r\n',

    b'    Case Push\r\n'
    b'        Dim mBranch As String\r\n'
    b'        If arg = Empty Then mBranch = "main" Else mBranch = arg\r\n'
    b'        Dim pushBookName As String: pushBookName = GetShortBookName(xBook.Name)\r\n'
    b'        Dim pushRepoName As String: pushRepoName = GetSetting("Excel", pushBookName, "RepositoryName")\r\n'
    b'        If pushRepoName = "" Then\r\n'
    + j('            MsgBox "リポジトリが登録されていません。", vbInformation\r\n')
    + b'            GoTo Finally\r\n'
    b'        End If\r\n'
    b'        Dim pushAccounts() As String: pushAccounts = GetAccountList()\r\n'
    b'        Dim pushIdx As Long\r\n'
    b'        For pushIdx = 0 To UBound(pushAccounts)\r\n'
    b'            Dim pushPat As String: pushPat = GetTokenFromRegistry(pushAccounts(pushIdx))\r\n'
    b'            If pushPat <> "" Then\r\n'
    b'                Dim pushUrl As String\r\n'
    b'                pushUrl = "https://" & pushPat & "@github.com/" & pushAccounts(pushIdx) & "/" & pushRepoName & ".git"\r\n'
    b'                rt = rt & vbCr & RunCmd("git push " & pushUrl & " " & mBranch)\r\n'
    b'            End If\r\n'
    b'        Next pushIdx\r\n'
))

# ============================================================
# 適用
# ============================================================
for i, (old, new) in enumerate(changes):
    if old in src:
        src = src.replace(old, new)
        print(f'変更 {i+1}: OK')
    else:
        print(f'変更 {i+1}: NOT FOUND')
        # デバッグ: 最初の50バイトを表示
        print(f'  探索: {repr(old[:80])}')

with open(SRC, 'wb') as f:
    f.write(src)

print('\n完了')
