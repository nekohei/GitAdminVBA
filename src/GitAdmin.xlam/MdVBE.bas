Attribute VB_Name = "MdVBE"
Option Explicit

Public Sub CloseAllVBEWindows()
    Dim vbWindow As VBIDE.Window
    On Error Resume Next
    For Each vbWindow In Application.VBE.Windows
        ' コードウィンドウのみを閉じる
        If vbWindow.Type = vbext_wt_CodeWindow Or vbWindow.Type = vbext_wt_Designer Then
            vbWindow.Close
        End If
    Next
    On Error GoTo 0
End Sub

