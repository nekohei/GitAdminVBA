Attribute VB_Name = "ModuleRibbon"
Option Explicit

Public Sub rbnリポジトリの作成(ByVal ctrl As IRibbonControl)
    Call CreateNewRepository
End Sub

Public Sub rbnトークンの登録(ByVal ctrl As IRibbonControl)
    Call RegisterToken
End Sub

Public Sub rbnトークンの削除(ByVal ctrl As IRibbonControl)
    Call DeleteToken
End Sub

Public Sub rbn変更をステージ(ByVal ctrl As IRibbonControl)
    Call GitCmd(Stage)
End Sub

Public Sub rbn変更をコミット(ByVal ctrl As IRibbonControl)
    Call GitCmd(Commit)
End Sub

Public Sub rbn変更をプッシュ(ByVal ctrl As IRibbonControl)
    Call GitCmd(Push)
End Sub


