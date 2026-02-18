# -*- coding: utf-8 -*-
"""RegisterPushAccounts を Sub→Function(Boolean) に変更し、
   CreateNewRepository でキャンセル時に処理終了するパッチ（行単位処理）"""

fpath = r'c:/Users/1784/claude/VBA/Excel/GitAdminVBA/src/ModuleGit.bas'

with open(fpath, 'rb') as f:
    raw = f.read()
text = raw.decode('cp932')

# ============================================================
# 1. CreateNewRepository 内の "Call RegisterPushAccounts()" を
#    "If Not RegisterPushAccounts() Then Exit Sub" に変更
#    ※ 前のコメント行ごと書き換える（コメント行は行単位で検索）
# ============================================================
# "    Call RegisterPushAccounts()" の行を探す（CreateNewRepository 内 = 最初の出現）
target_call = '    Call RegisterPushAccounts()'
replacement_call = '    If Not RegisterPushAccounts() Then Exit Sub'

# 最初の出現箇所（CreateNewRepository 内）のみ変更
idx = text.find(target_call)
assert idx >= 0, 'ERROR: Call RegisterPushAccounts() not found'
text = text[:idx] + replacement_call + text[idx + len(target_call):]
print('1. Call → If Not done')

# ============================================================
# 2. RegisterPushAccounts の定義を Sub → Function(Boolean) に変更
# ============================================================
# "Public Sub RegisterPushAccounts()" を "Public Function RegisterPushAccounts() As Boolean" に
text = text.replace(
    'Public Sub RegisterPushAccounts()',
    'Public Function RegisterPushAccounts() As Boolean',
    1
)
print('2. Sub → Function done')

# "    If StrPtr(ans) = 0 Then Exit Sub  " を含む行を
# "    If StrPtr(ans) = 0 Then\r\n        RegisterPushAccounts = False  'キャンセル\r\n        Exit Function\r\n    End If" に変更
# まず行を探す
strptr_line_lf  = "    If StrPtr(ans) = 0 Then Exit Sub  '"
strptr_block_crlf = (
    "    If StrPtr(ans) = 0 Then\r\n"
    "        RegisterPushAccounts = False  'キャンセル\r\n"
    "        Exit Function\r\n"
    "    End If"
)
strptr_block_lf = (
    "    If StrPtr(ans) = 0 Then\n"
    "        RegisterPushAccounts = False  'キャンセル\n"
    "        Exit Function\n"
    "    End If"
)

# StrPtr 行の終端（日本語コメント含む）を検索して行末まで丸ごと置換
import re
# "    If StrPtr(ans) = 0 Then Exit Sub" で始まる行を見つける
m = re.search(r"    If StrPtr\(ans\) = 0 Then Exit Sub[^\n]*", text)
assert m, 'ERROR: StrPtr line not found'
# 改行コードに合わせて挿入
after = text[m.end():]
if after.startswith('\r\n'):
    block = strptr_block_crlf
elif after.startswith('\n'):
    block = strptr_block_lf
else:
    block = strptr_block_crlf
text = text[:m.start()] + block + text[m.end():]
print('3. StrPtr block done')

# "    Exit Sub" (末尾付近) を "    RegisterPushAccounts = True\r\n    Exit Function" に変更
# RegisterPushAccounts 関数内の "Exit Sub" を探す
# ※ "Catch:" の前の "Exit Sub" が対象
# Catch: の直前にある "    Exit Sub" を置換
# パターン: "    Exit Sub\r\nCatch:" または "    Exit Sub\nCatch:"
exit_sub_crlf = "    Exit Sub\r\nCatch:"
exit_sub_lf   = "    Exit Sub\nCatch:"
exit_func_crlf = "    RegisterPushAccounts = True\r\n    Exit Function\r\nCatch:"
exit_func_lf   = "    RegisterPushAccounts = True\n    Exit Function\nCatch:"

if exit_sub_crlf in text:
    text = text.replace(exit_sub_crlf, exit_func_crlf, 1)
    print('4a. Exit Sub → Exit Function (CRLF) done')
elif exit_sub_lf in text:
    text = text.replace(exit_sub_lf, exit_func_lf, 1)
    print('4b. Exit Sub → Exit Function (LF) done')
else:
    print('ERROR: Exit Sub before Catch not found')

# "End Sub" (RegisterPushAccounts の末尾) を
# "    RegisterPushAccounts = False\r\nEnd Function" に変更
# ※ Catch: ブロック内の "MsgBox Err.Description, vbExclamation" の後の "End Sub"
end_sub_pattern = "    MsgBox Err.Description, vbExclamation\r\nEnd Sub"
end_func_replacement = "    MsgBox Err.Description, vbExclamation\r\n    RegisterPushAccounts = False\r\nEnd Function"
end_sub_pattern_lf = "    MsgBox Err.Description, vbExclamation\nEnd Sub"
end_func_replacement_lf = "    MsgBox Err.Description, vbExclamation\n    RegisterPushAccounts = False\nEnd Function"

if end_sub_pattern in text:
    text = text.replace(end_sub_pattern, end_func_replacement, 1)
    print('5a. End Sub → End Function (CRLF) done')
elif end_sub_pattern_lf in text:
    text = text.replace(end_sub_pattern_lf, end_func_replacement_lf, 1)
    print('5b. End Sub → End Function (LF) done')
else:
    print('ERROR: End Sub pattern not found')

# Dim ans の宣言を1行スタイルから2行スタイルに変更（StrPtr が確実に動作するよう）
text = text.replace(
    '    Dim ans As String: ans = InputBox(',
    '    Dim ans As String\r\n    ans = InputBox(',
    1
)
print('6. Dim ans split done')

with open(fpath, 'wb') as f:
    f.write(text.encode('cp932'))

print('OK: patch3 applied')
