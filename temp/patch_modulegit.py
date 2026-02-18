# -*- coding: utf-8 -*-
"""ModuleGit.bas に GetPushAccounts / RegisterPushAccounts を追加するパッチスクリプト"""

fpath = r'c:/Users/1784/claude/VBA/Excel/GitAdminVBA/src/ModuleGit.bas'

with open(fpath, 'rb') as f:
    raw = f.read()
text = raw.decode('cp932')

# ============================================================
# 1. Case Push: GetAccountList() -> GetPushAccounts(xBook.Name)
# ============================================================
old = '        Dim pushAccounts() As String: pushAccounts = GetAccountList()'
new = '        Dim pushAccounts() As String: pushAccounts = GetPushAccounts(xBook.Name)'
assert old in text, 'ERROR: push target not found'
text = text.replace(old, new, 1)

# ============================================================
# 2. Case Init: GetAccountList() -> GetPushAccounts(xBook.Name)
# ============================================================
old = '        Dim initAccounts() As String: initAccounts = GetAccountList()'
new = '        Dim initAccounts() As String: initAccounts = GetPushAccounts(xBook.Name)'
assert old in text, 'ERROR: init target not found'
text = text.replace(old, new, 1)

# ============================================================
# 3. GetAccountList 末尾の後に GetPushAccounts / RegisterPushAccounts を挿入
#    挿入点: GetAccountList のエラー時の戻り値行 "    GetAccountList = Array()\r\nEnd Function\r\n"
# ============================================================
MARKER = '    GetAccountList = Array()\r\nEnd Function\r\n'
assert MARKER in text, 'ERROR: GetAccountList end marker not found'

# CP932 文字列として追加コード（日本語コメント含む）を定義
ADD_CODE = r"""
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

"""

text = text.replace(MARKER, MARKER + ADD_CODE, 1)

with open(fpath, 'wb') as f:
    f.write(text.encode('cp932'))

print('OK: ModuleGit.bas patched successfully')
