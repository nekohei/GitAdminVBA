# -*- coding: utf-8 -*-
"""RegisterPushAccounts 内の Exit Sub / End Sub を修正するパッチ"""

fpath = r'c:/Users/1784/claude/VBA/Excel/GitAdminVBA/src/ModuleGit.bas'

with open(fpath, 'rb') as f:
    raw = f.read()
text = raw.decode('cp932')

# RegisterPushAccounts の開始・終了位置を特定して範囲内で置換
func_start = text.find('Public Function RegisterPushAccounts() As Boolean')
assert func_start >= 0, 'Function not found'
# End Sub か End Function が来るまでの終端を探す
func_end = text.find('\nEnd Sub', func_start)
assert func_end >= 0, 'End Sub not found in function'
func_end += len('\nEnd Sub')

func_body = text[func_start:func_end]
print('=== 現在の関数本体 ===')
for i, l in enumerate(func_body.split('\n')):
    print(str(i) + ': ' + repr(l))

# 関数本体内で以下を置換:
# 1. "        Exit Sub" (allAccounts チェック内) → "        RegisterPushAccounts = False\r\n        Exit Function"
# 2. "    Exit Sub" (Catch の直前) → "    RegisterPushAccounts = True\r\n    Exit Function"
# 3. "End Sub" (末尾) → "    RegisterPushAccounts = False\r\nEnd Function"

# 注意: 文字列内で順番に適用する (先に固有のパターンから)

# 1. "        Exit Sub" (8スペース = allAccounts チェック内)
old1 = '        Exit Sub\r\n    End If'
new1 = '        RegisterPushAccounts = False\r\n        Exit Function\r\n    End If'
old1_lf = '        Exit Sub\n    End If'
new1_lf = '        RegisterPushAccounts = False\n        Exit Function\n    End If'

if old1 in func_body:
    func_body = func_body.replace(old1, new1, 1)
    print('1. 8sp Exit Sub fixed (CRLF)')
elif old1_lf in func_body:
    func_body = func_body.replace(old1_lf, new1_lf, 1)
    print('1. 8sp Exit Sub fixed (LF)')
else:
    print('WARNING: pattern1 not found')

# 2. "    Exit Sub\r\nCatch:" (4スペース = 正常終了後)
old2 = '    Exit Sub\r\nCatch:'
new2 = '    RegisterPushAccounts = True\r\n    Exit Function\r\nCatch:'
old2_lf = '    Exit Sub\nCatch:'
new2_lf = '    RegisterPushAccounts = True\n    Exit Function\nCatch:'

if old2 in func_body:
    func_body = func_body.replace(old2, new2, 1)
    print('2. 4sp Exit Sub→Function fixed (CRLF)')
elif old2_lf in func_body:
    func_body = func_body.replace(old2_lf, new2_lf, 1)
    print('2. 4sp Exit Sub→Function fixed (LF)')
else:
    print('WARNING: pattern2 not found')

# 3. "End Sub" → "    RegisterPushAccounts = False\r\nEnd Function"
old3 = '\nEnd Sub'
new3 = '\n    RegisterPushAccounts = False\r\nEnd Function'
old3_lf = '\nEnd Sub'
new3_lf = '\n    RegisterPushAccounts = False\nEnd Function'

# func_body の末尾の End Sub を置換
last_end_sub = func_body.rfind('\nEnd Sub')
if last_end_sub >= 0:
    func_body = func_body[:last_end_sub] + '\n    RegisterPushAccounts = False\r\nEnd Function'
    print('3. End Sub → End Function fixed')
else:
    print('WARNING: End Sub not found')

print('\n=== 修正後の関数本体 ===')
for i, l in enumerate(func_body.split('\n')):
    print(str(i) + ': ' + repr(l))

# ファイル全体を再結合
text = text[:func_start] + func_body + text[func_start + (func_end - func_start):]

# ステップ4・5が誤って他の関数に適用されていないか確認・修正
# RegisterToken と DeleteToken の End Function への誤適用を確認
import re
wrong = list(re.finditer(r'RegisterPushAccounts = (?:True|False).*?(?=\nPublic|\nPrivate|\Z)', text, re.DOTALL))
print('\n=== RegisterPushAccounts 代入箇所 ===')
for m in wrong:
    start_line = text[:m.start()].count('\n') + 1
    print('L' + str(start_line) + ': ' + repr(m.group()[:80]))

with open(fpath, 'wb') as f:
    f.write(text.encode('cp932'))

print('\nOK: patch4 applied')
