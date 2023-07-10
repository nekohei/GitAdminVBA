Attribute VB_Name = "ModuleMenu"
Option Explicit

Public Enum GitCommand
    Stage
    Commit
    Push
End Enum

Private Const TempFileName = "TempGitOutput"

Public Sub ExportCodeModules(ByVal xBook As Workbook, ByVal srcPath As String)
    On Error GoTo Catch
    Dim vbPjt As VBIDE.VBProject: Set vbPjt = xBook.VBProject
    
    Dim vbCmp As VBIDE.VBComponent
    For Each vbCmp In vbPjt.VBComponents
        Select Case vbCmp.Type
            Case vbext_ct_StdModule
                vbCmp.Export srcPath & "¥" & vbCmp.Name & ".bas"
            Case vbext_ct_MSForm
                vbCmp.Export srcPath & "¥" & vbCmp.Name & ".frm"
            Case vbext_ct_ClassModule
                vbCmp.Export srcPath & "¥" & vbCmp.Name & ".cls"
            Case vbext_ct_Document
                vbCmp.Export srcPath & "¥" & vbCmp.Name & ".dcm"
        End Select
    Next
    MsgBox "Export が完了しました。", vbInformation
    GoTo Finally
Catch:
    OutputError "ExportCodeModules"
Finally:
    ' 何もしない
End Sub
 
Public Function RunCmd(cmd As String, Optional showInt As Integer = 0, Optional toWait As Boolean = True) As String

    Dim tmpPath As String: tmpPath = Environ$("temp") & "¥" & TempFileName
    Call GenerateUTF8("_", tmpPath)
    
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim rootDir As String: rootDir = GetRootDir
    If rootDir = "" Then
        MsgBox "リポジトリ名が登録されていません。", vbInformation
        Exit Function
    End If
    Call ChDir(rootDir)
    Call wsh.Run("cmd /c " & cmd, showInt, toWait)

    Dim sm As Object: Set sm = CreateObject("ADODB.Stream")
    sm.Type = 2
    sm.Charset = "utf-8"
    sm.Open
    sm.LoadFromFile tmpPath
    Dim tmp As String: tmp = sm.ReadText
    sm.Close: Set sm = Nothing
    
    RunCmd = tmp

End Function

Public Sub GitCmd(cmd As GitCommand, Optional arg As String = Empty)
    On Error GoTo Catch
    Dim rootDir As String: rootDir = GetRootDir
    If rootDir = "" Then
        MsgBox "リポジトリ名が登録されていません。", vbInformation
        Exit Sub
    End If
    Call ChDir(rootDir)
    Dim tmpPath As String: tmpPath = Environ$("temp") & "¥TempGitOutput"
    Call GenerateUTF8("_", tmpPath)
    
    Dim rt As String
    Select Case cmd
    Case Stage
        If MsgBox(ActiveWorkbook.Name & " の変更をステージします。" & vbLf & vbLf & _
                  ActiveWorkbook.Name & " の保存とエクスポートを伴います。", vbInformation + vbOKCancel) = vbOK Then
            Application.DisplayAlerts = False
            ActiveWorkbook.Save
            Call Decombine
            rt = RunCmd("git add . | git status")
        Else
            GoTo Finally
        End If
    Case Commit
        If arg = Empty Then
            MsgBox "コメントは必須です。", vbInformation + vbOKOnly
            GoTo Finally
        End If
        rt = RunCmd("git commit -m """ & arg & """ > " & tmpPath)
    Case Push
        Dim mBranch As String
        If arg = Empty Then mBranch = "main"
        rt = RunCmd("git push origin " & mBranch)
    End Select
    Debug.Print rt
    GoTo Finally
Catch:
    OutputError "GitCmd"
Finally:
    Application.DisplayAlerts = True
End Sub

Private Sub Decombine()
    Dim srcPath As String: srcPath = GetRootDir & "¥src¥" & ActiveWorkbook.Name
    Call 指定フォルダが無ければ作る(srcPath)
    Call ExportCodeModules(ActiveWorkbook, srcPath)
End Sub

Public Sub CreateGitFolder()
    Dim xName As String: xName = ActiveWorkbook.Name
    xName = Left(xName, Len(xName) - 5)
    xName = StrConv(xName, vbNarrow)
    xName = Replace(xName, " ", "_")
    Dim savePath As String: savePath = ParentDir & "¥" & xName
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(savePath) Then fso.CreateFolder savePath
    Set fso = Nothing
    ChDir savePath
End Sub

Public Sub SetRepositoryName()
    Dim reposName As String: reposName = ActiveWorkbook.BuiltinDocumentProperties(5).Value
    If reposName = "" Then
        reposName = InputBox("リポジトリー名を英字で入力してください。")
        If reposName = "" Then Exit Sub
        If CheckReposName(reposName) = "" Then
            MsgBox "無効なリポジトリ名です。", vbInformation
            Exit Sub
        End If
        ActiveWorkbook.BuiltinDocumentProperties(5).Value = reposName
    End If
End Sub

Private Function CheckReposName(ByVal stg As String) As String
    
    Dim i As Integer
    For i = 1 To Len(stg)
        Select Case Asc(Mid(stg, i, 1))
        Case 0 To 127
            If InStr("/¥@‾ ", Mid(stg, i, 1)) > 0 Or _
                (i = Len(stg) And Mid(stg, i, 1) = ".") Then _
                    GoTo Invalid
        Case Else
            GoTo Invalid
        End Select
    Next
    
    CheckReposName = stg
    Exit Function
Invalid:
    CheckReposName = ""
End Function

Private Function GetRootDir() As String
    Dim reposName As String
    reposName = ActiveWorkbook.BuiltinDocumentProperties(5).Value
    If reposName = "" Then
        GetRootDir = ""
    Else
        GetRootDir = ParentDir & "¥" & reposName
    End If
End Function
