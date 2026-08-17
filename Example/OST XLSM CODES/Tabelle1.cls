VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    ' B列：时间戳自动格式化
    If Not Intersect(Target, Me.Range("B2:B2000")) Is Nothing Then
        FormatTimestampInput Target
    End If

    ' No-Match 临时ID管理
    If Not Intersect(Target, Me.Range("A2:F2000")) Is Nothing Then
        HandleNoMatchWorkflow Target
    End If

    ' 原有 OST Timeline → Collection 同步逻辑
    HandleTimelineChange Target

    ' 更新 Match / No Match / Total / First Appears / EP Stats
    If Not Intersect(Target, Me.Range("A2:F2000")) Is Nothing Then
        UpdateOSTStatistics
    End If

End Sub

Private Sub Worksheet_Activate()
    Me.Range("C:F").EntireColumn.AutoFit
End Sub
