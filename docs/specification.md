# GitAdminVBA 仕様書

## 目次

1. [システム概要](#1-システム概要)
2. [レジストリ設計](#2-レジストリ設計)
3. [ブック名の正規化（GetShortBookName）](#3-ブック名の正規化getshortbookname)
4. [アドイン管理](#4-アドイン管理)
5. [VBA コーディング仕様](#5-vba-コーディング仕様)
6. [ソースコードのエンコーディング](#6-ソースコードのエンコーディング)
7. [主要処理フロー](#7-主要処理フロー)

---

## 1. システム概要

Excel アドイン (`bin/Git管理.xlam`) を通じて、VBA プロジェクトを GitHub で管理するツール。

| 項目 | 内容 |
|------|------|
| アドインファイル | `bin/Git管理.xlam` |
| アドインフォルダ | `C:\Users\<user>\AppData\Roaming\Microsoft\AddIns\Git管理.xlam` |
| ソースコード | `src/` フォルダ（テキスト、git 管理対象） |
| リポジトリ管理先 | `%USERPROFILE%\Source\Repos\VBA\<repoName>\` |

---

## 2. レジストリ設計

VBA の `GetSetting` / `SaveSetting` は `HKCU\Software\VB and VBA Program Settings\` 配下に保存される。

### 2.1 GitHub アカウント設定（PAT）

```
HKCU\Software\VB and VBA Program Settings\GitHub\<accountName>\Classic = <PAT>
```

| キー | 内容 |
|------|------|
| AppName | `GitHub` |
| Section | アカウント名（例: `nekohei`, `sekinekh`） |
| Key | `Classic` |
| Value | GitHub Personal Access Token |

**操作関数:**
- `RegisterToken()` — アカウント名と PAT を InputBox で入力して保存
- `GetTokenFromRegistry(accountName)` — PAT を取得
- `DeleteToken()` — 指定アカウントの PAT を削除
- `GetAccountList()` — 登録済みアカウント一覧を取得（`reg query` コマンドを使用）

> **注意:** `GetAccountList()` は従来 WMI (`winmgmts:StdRegProv.EnumKey`) を使っていたが、
> 環境によってエラーになるため `WScript.Shell + cmd /c reg query` に変更した（2026-02-18）。

### 2.2 ブック個別設定

```
HKCU\Software\VB and VBA Program Settings\Excel\<shortBookName>\RepositoryName = <repoName>
HKCU\Software\VB and VBA Program Settings\Excel\<shortBookName>\PushAccounts   = <accounts>
```

| キー | 内容 |
|------|------|
| AppName | `Excel` |
| Section | **shortBookName**（後述） |
| `RepositoryName` | GitHub リポジトリ名 |
| `PushAccounts` | プッシュ先アカウント（カンマ区切り、空=全アカウント） |

**操作関数:**
- `SetAndThenGetReposName(bookName)` — リポジトリ名を取得（未登録なら InputBox で入力・保存）
- `RegisterPushAccounts()` — プッシュ先アカウントを InputBox で設定・保存（`Boolean` を返す）
- `GetPushAccounts(bookName)` — プッシュ先アカウント配列を返す（未設定時は全アカウント）

### 2.3 PushAccounts の仕様

- 複数アカウントはカンマ区切りで保存（例: `nekohei,sekinekh`）
- **未設定（空文字）の場合:** `GetPushAccounts()` が `GetAccountList()` を返す（全アカウントが対象）
- Push 時に PAT が空のアカウントは自動的にスキップされる

---

## 3. ブック名の正規化（GetShortBookName）

```vba
Private Function GetShortBookName(ByVal bookName As String) As String
```

運用ブックはファイル名の後に `_YYYYMMDD`（アンダースコア＋日付8桁）が付加される。
レジストリキーにはベース名（日付なし）を使うため、`GetShortBookName` で正規化する。

| 入力例 | 出力例 | ロジック |
|--------|--------|---------|
| `MyBook_20240218.xlsm` | `MyBook.xlsm` | アンダースコア以降〜最後のドット直前を除去 |
| `MyBook.xlsm` | `MyBook.xlsm` | アンダースコアなし → そのまま |
| `MyBook` | `MyBook` | ドットなし → そのまま |

---

## 4. アドイン管理

### 4.1 VBA エクスポート・インポート

| スクリプト | 処理 |
|-----------|------|
| `temp/export-vba.ps1` | `bin/Git管理.xlam` → `src/` へエクスポート |
| `temp/import-vba.ps1` | `src/` → `bin/Git管理.xlam` へインポート |
| `temp/deploy-test.ps1` | `bin/Git管理.xlam` → アドインフォルダへ上書きコピー（バックアップあり） |
| `temp/restore-addin.ps1` | バックアップ → アドインフォルダへ復元 |

**インポート前提条件:**
- **Excel をすべて閉じた状態で実行すること**
  - アドイン (`Git管理.xlam`) が Excel に読み込まれたまま上書きしようとするとファイルロックエラーになる
  - `import-vba.ps1` は起動中の Excel プロセスを確認し、検出時は警告して終了する

**`.dcm` ファイルのスキップ:**
- `ThisWorkbook` / `Sheet1` などのドキュメントモジュール (`.dcm`) はインポート時にスキップする
- `VBComponents.AddFromString` に `.dcm` の内容を渡すと Excel がハングする問題があるため

### 4.2 アドインのバックアップ

- テスト用デプロイ前に `bin/history/Git管理_addin_backup.xlam` へ自動バックアップ
- `restore-addin.ps1` でバックアップから復元できる

### 4.3 両アカウントへの push

`push-all.ps1` により `sekinekh/GitAdminVBA` と `nekohei/GitAdminVBA` の両リポジトリへ一括 push できる。
PAT は `nekohei` のレジストリ値から取得している。

---

## 5. VBA コーディング仕様

### 5.1 InputBox のキャンセル判定（StrPtr）

VBA の `InputBox` でキャンセルボタンを押した場合、戻り値は空文字 `""` だが、
`StrPtr` を使うとキャンセル（`0`）と空文字入力を区別できる。

```vba
' NG: 1行スタイルでは StrPtr が正常動作しないことがある
Dim ans As String: ans = InputBox("入力してください")

' OK: 宣言と代入を分離する
Dim ans As String
ans = InputBox("入力してください")
If StrPtr(ans) = 0 Then
    ' キャンセル時の処理
    Exit Function
End If
```

> **重要:** `Dim ans As String: ans = InputBox(...)` の1行スタイルでは
> `StrPtr(ans)` が正しく `0` を返さない場合があるため、**宣言と代入を必ず分ける**。

### 5.2 キャンセルの呼び元への伝播（Sub → Function）

呼び元の処理をキャンセル時に中断させたい場合は、`Sub` ではなく `Function(Boolean)` にして戻り値で制御する。

```vba
' キャンセルを呼び元に伝える場合は Function にする
Public Function SomeDialog() As Boolean
    Dim ans As String
    ans = InputBox("...")
    If StrPtr(ans) = 0 Then
        SomeDialog = False   ' キャンセル
        Exit Function
    End If
    ' ... 処理 ...
    SomeDialog = True        ' 正常完了
    Exit Function
Catch:
    MsgBox Err.Description, vbExclamation
    SomeDialog = False
End Function

' 呼び元では戻り値を確認して中断
Public Sub MainProcess()
    If Not SomeDialog() Then Exit Sub   ' キャンセルで終了
    ' ... 後続処理 ...
End Sub
```

### 5.3 エラーハンドラの構造

```vba
Public Function SomeFunction() As ReturnType
    On Error GoTo Catch
    ' ... 処理 ...
    SomeFunction = 正常値
    Exit Function
Catch:
    MsgBox Err.Description, vbExclamation
    SomeFunction = エラー値
End Function
```

- `GoTo Finally` パターン（`Application.DisplayAlerts` の復元など後処理が必要な場合）も使用している

### 5.4 WMI の代替（reg query）

WMI (`winmgmts:StdRegProv.EnumKey`) は環境によって動作しないことがある。
レジストリキーの列挙には `WScript.Shell` + `cmd /c reg query` を使う。

```vba
Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
Dim oExec As Object
Set oExec = wsh.Exec("cmd /c reg query ""HKCU\Software\VB and VBA Program Settings\GitHub""")
Dim output As String
Do While Not oExec.StdOut.AtEndOfStream
    output = output & oExec.StdOut.ReadLine() & vbLf
Loop
```

出力は `HKEY_CURRENT_USER\...\GitHub\<accountName>` の形式で1行ずつ返るため、
`baseKey` プレフィックスで始まる行のサフィックス部分をアカウント名として取得する。

---

## 6. ソースコードのエンコーディング

| ファイル種別 | エンコーディング | 理由 |
|------------|----------------|------|
| `src/*.bas` / `src/*.cls` / `src/*.dcm` | CP932（Shift-JIS） | Excel の VBA エクスポートが日本語 Windows で CP932 で出力するため |
| `temp/*.ps1` | UTF-8 with BOM | PowerShell 5.x と 7+ の両方に対応するため |
| `temp/*.py` | UTF-8（BOMなし） | Python はデフォルト UTF-8 |

### Python で src/ ファイルを編集する場合

```python
# 読み込み
with open(fpath, 'rb') as f:
    raw = f.read()
text = raw.decode('cp932')

# 編集処理
text = text.replace(old, new)

# 書き戻し
with open(fpath, 'wb') as f:
    f.write(text.encode('cp932'))
```

**注意事項:**
- 日本語コメントが含まれるため、Python スクリプト自体は UTF-8 で書き、ファイルの読み書きのみ `cp932` を使う
- 改行コードが CRLF/LF 混在することがある（git の autocrlf の影響）
- 文字列パターンマッチングは改行コードを意識して行うか、行単位で処理する
- `assert` で事前確認し、置換対象が見つからない場合は早期終了させる

---

## 7. 主要処理フロー

### 7.1 新規リポジトリ作成（CreateNewRepository）

```
1. SetAndThenGetReposName()  → リポジトリ名を入力・レジストリ保存
       ↓ キャンセルで終了
2. RegisterPushAccounts()    → プッシュ先アカウントを入力・レジストリ保存
       ↓ キャンセルで終了
3. フォルダ作成（.vscode, bin, bin\old, src）
4. GenerateGitFiles()        → .gitignore, settings.json 等を生成
5. xBook.Save → bin\ にコピー
6. Decombine()               → VBA を src\ にエクスポート
7. GetPushAccounts() でアカウント取得
8. CreateRemoteRepos()       → 各アカウントに GitHub API でリモートリポジトリ作成
9. GitCmd(Init)              → git init / add / commit / push
```

**PAT の前提:** ステップ 7〜9 で使用する PAT は事前に `RegisterToken()` で登録済みであること。
PAT が未登録のアカウントは push がスキップされる（エラーにはならない）。

### 7.2 Push（GitCmd - Case Push）

```
1. GetSetting("Excel", shortBookName, "RepositoryName") → リポジトリ名取得
       ↓ 未登録の場合はメッセージ表示して終了
2. GetPushAccounts(bookName) → プッシュ先アカウント配列取得
3. 各アカウントについて:
   a. GetTokenFromRegistry(account) → PAT 取得（空なら skip）
   b. https://<PAT>@github.com/<account>/<repo>.git の URL を構築
   c. git push <url> <branch> を実行
```

### 7.3 レジストリ設定の優先順位

```
PushAccounts が設定済み → 設定されたアカウントのみ push
PushAccounts が未設定   → GetAccountList() の全アカウントに push
```

---

## 変更履歴

| 日付 | 変更内容 |
|------|---------|
| 2026-02-18 | 初版作成。GetAccountList の WMI→reg query 改修、ブックごとのプッシュ先設定機能追加に伴い仕様を整理 |
