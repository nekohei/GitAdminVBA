VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GitEventsItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem イベント配列（1始まりの配列）
Private mEventsItems() As GitEventsItem
Private mMaxIndex As Long
Public Event Click(ByVal ctl As Object, handled As Boolean, cancelDefault As Boolean)

Public Sub AddItem(ctl As CommandBarControl)
    If ctl Is Nothing Then Exit Sub
    Rem イベント配列の要素を増やす
    mMaxIndex = mMaxIndex + 1
    ReDim Preserve mEventsItems(1 To mMaxIndex)
    Rem イベントの生成と初期化
    Set mEventsItems(mMaxIndex) = New GitEventsItem
    Call mEventsItems(mMaxIndex).Initialize(ctl, Me)
End Sub

Rem イベントから伝播された引数をクラス利用者に伝播する
Public Sub RaiseClick(ByVal ctl As Object, handled As Boolean, cancelDefault As Boolean)
    RaiseEvent Click(ctl, handled, cancelDefault)
End Sub

Private Sub Class_Terminate()
    Rem イベント配列の解放
    Dim i As Long
    For i = 1 To mMaxIndex
        Set mEventsItems(i) = Nothing
    Next
End Sub
