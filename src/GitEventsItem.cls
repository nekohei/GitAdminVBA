VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GitEventsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mEventItem As VBIDE.CommandBarEvents
Attribute mEventItem.VB_VarHelpID = -1
Private mParent As GitEventsItems

Public Sub Initialize(ByVal ctl As Office.CommandBarControl, ByRef prt As GitEventsItems)
    Rem メニュー項目とイベントの紐づけ
    Set mEventItem = Application.VBE.Events.CommandBarEvents(ctl)
    Set mParent = prt
End Sub

Private Sub mEventItem_Click(ByVal ctl As Object, handled As Boolean, cancelDefault As Boolean)
    Rem イベントで得た引数をイベント配列に伝播する
    Call mParent.RaiseClick(ctl, handled, cancelDefault)
End Sub

Private Sub Class_Terminate()
    Set mEventItem = Nothing
    Set mParent = Nothing
End Sub

