# -*- coding: utf-8 -*-
"""CreateNewRepository にプッシュ先アカウント選択を追加するパッチ"""

fpath = r'c:/Users/1784/claude/VBA/Excel/GitAdminVBA/src/ModuleGit.bas'

with open(fpath, 'rb') as f:
    raw = f.read()
text = raw.decode('cp932')

# ============================================================
# 1. リポジトリ名設定後にプッシュ先アカウント選択を追加
#    挿入点: "If repoName = "" Then Exit Sub" の直後
# ============================================================
OLD1 = '    Dim repoName As String: repoName = SetAndThenGetReposName(xBook.Name)\r\n    If repoName = "" Then Exit Sub'
NEW1 = '''    Dim repoName As String: repoName = SetAndThenGetReposName(xBook.Name)
    If repoName = "" Then Exit Sub

    'プッシュ先アカウントを選択（スキップすると全アカウント）
    Call RegisterPushAccounts()'''
assert OLD1 in text, 'ERROR: target1 not found'
text = text.replace(OLD1, NEW1, 1)

# ============================================================
# 2. CreateNewRepository 内の GetAccountList() を GetPushAccounts() に変更
#    "accounts = GetAccountList()" は CreateNewRepository 内のもの（L67相当）
# ============================================================
OLD2 = "    '登録済みアカウントごとにリモートリポジトリを作成\r\n    Dim accounts() As String: accounts = GetAccountList()"
# OLD2が見つからない場合は別のパターンも試す
if OLD2 not in text:
    # コメントなしパターン
    OLD2 = "    Dim accounts() As String: accounts = GetAccountList()\r\n    Dim successCount As Long"
    NEW2 = "    Dim accounts() As String: accounts = GetPushAccounts(xBook.Name)\r\n    Dim successCount As Long"
    if OLD2 in text:
        text = text.replace(OLD2, NEW2, 1)
        print('Pattern2b matched')
    else:
        # 単純パターン
        import re
        # CreateNewRepository 内の GetAccountList 呼び出しを探す（最初の1箇所のみ）
        pattern = r'(    Dim accounts\(\) As String: accounts = )GetAccountList\(\)'
        match = re.search(pattern, text)
        if match:
            text = text[:match.start()] + text[match.start():match.end()].replace('GetAccountList()', 'GetPushAccounts(xBook.Name)') + text[match.end():]
            print('Pattern2c (regex) matched')
        else:
            print('ERROR: target2 not found')
else:
    NEW2 = "    '登録済みアカウントごとにリモートリポジトリを作成\r\n    Dim accounts() As String: accounts = GetPushAccounts(xBook.Name)"
    text = text.replace(OLD2, NEW2, 1)
    print('Pattern2a matched')

with open(fpath, 'wb') as f:
    f.write(text.encode('cp932'))

print('OK: patch2 applied')
