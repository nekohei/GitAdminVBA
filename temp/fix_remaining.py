# fix_remaining.py - 実バイト列で正確に置換

SRC = 'c:/Users/1784/claude/VBA/Excel/GitAdminVBA/src/ModuleGit.bas'

def j(text):
    return text.encode('cp932')

with open(SRC, 'rb') as f:
    src = f.read()

# 実バイト列から構築（デバッグ出力を元に）
# コメント行: '  リモートリポジトリの作成
comment  = b'    \' \x83\x8a\x83\x82\x81[\x83g\x83\x8a\x83|\x83W\x83g\x83\x8a\x82\xcc\x8d\xec\x90\xac\r\n'
# MsgBoxのメッセージ: 用のリポジトリの準備ができました。
msgtext  = b'\x97p\x82\xcc\x83\x8a\x83|\x83W\x83g\x83\x8a\x82\xcc\x8f\x80\x94\xf5\x82\xaa\x82\xc5\x82\xab\x82\xdc\x82\xb5\x82\xbd\x81B'

old = (
    comment
    + b'    If CreateRemoteRepos(bookName, repoName) Then\r\n'
    + b'        MsgBox bookName & " ' + msgtext + b'", vbInformation\r\n'
    + b'    End If'
)

new = (
    b'    \'' + j('登録済みアカウントごとにリモートリポジトリを作成') + b'\r\n'
    + b'    Dim accounts() As String: accounts = GetAccountList()\r\n'
    + b'    Dim successCount As Long: successCount = 0\r\n'
    + b'    Dim acIdx As Long\r\n'
    + b'    For acIdx = 0 To UBound(accounts)\r\n'
    + b'        If CreateRemoteRepos(bookName, repoName, accounts(acIdx)) Then\r\n'
    + b'            successCount = successCount + 1\r\n'
    + b'        End If\r\n'
    + b'    Next acIdx\r\n'
    + b'    If successCount > 0 Then\r\n'
    + b'        Call GitCmd(Init)\r\n'
    + b'        MsgBox bookName & " ' + msgtext + b'", vbInformation\r\n'
    + b'    End If'
)

if old in src:
    src = src.replace(old, new)
    print('CreateNewRepository: OK')
else:
    print('CreateNewRepository: NOT FOUND')

with open(SRC, 'wb') as f:
    f.write(src)

print('完了')
