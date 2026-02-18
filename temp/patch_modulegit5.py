# -*- coding: utf-8 -*-
"""RegisterToken の誤適用を修正するパッチ"""

fpath = r'c:/Users/1784/claude/VBA/Excel/GitAdminVBA/src/ModuleGit.bas'

with open(fpath, 'rb') as f:
    raw = f.read()
text = raw.decode('cp932')

# RegisterToken の開始位置を特定
func_start = text.find('Public Sub RegisterToken()')
assert func_start >= 0, 'RegisterToken not found'

# その後の End Function を探す（誤変更された末尾）
func_end = text.find('\nEnd Function', func_start)
assert func_end >= 0
func_end += len('\nEnd Function')

func_body = text[func_start:func_end]
print('=== 現在の RegisterToken ===')
for i, l in enumerate(func_body.split('\n')):
    print(str(i) + ': ' + repr(l))

# 修正: 誤挿入された行を削除し、正しい内容に戻す
# "    RegisterPushAccounts = True\r\n    Exit Function\r\nCatch:" を "    Exit Sub\r\nCatch:" に
# "    RegisterPushAccounts = False\r\nEnd Function" を "End Sub" に

# パターン1: RegisterPushAccounts = True の前後
# 改行コードの混在があるので regex を使う
import re

# "    RegisterPushAccounts = True[\r\n]+    Exit Function[\r\n]+Catch:"
#  → "    Exit Sub\r\nCatch:"
func_body = re.sub(
    r'    RegisterPushAccounts = True\r?\n    Exit Function\r?\n(Catch:)',
    r'    Exit Sub\r\n\1',
    func_body
)

# "    RegisterPushAccounts = False[\r\n]+End Function"
#  → "End Sub"
func_body = re.sub(
    r'    RegisterPushAccounts = False\r?\nEnd Function',
    'End Sub',
    func_body
)

print('\n=== 修正後の RegisterToken ===')
for i, l in enumerate(func_body.split('\n')):
    print(str(i) + ': ' + repr(l))

# ファイルに書き戻す
text = text[:func_start] + func_body + text[func_end:]

with open(fpath, 'wb') as f:
    f.write(text.encode('cp932'))

print('\nOK: patch5 applied')
