VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'========================
'3) Tabelle2£¨OST Collection£©´úÂë
'========================
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    HandleCollectionChange Target
End Sub

Private Sub Worksheet_Activate()
    Me.Range("A:E").EntireColumn.AutoFit
    Me.Range("H:I").EntireColumn.AutoFit

    Me.Columns(CL_DISPLAY_COL).Hidden = True
    Me.Columns(CL_ID_COL).Hidden = True
    
    SyncEpisodeStatsButton
End Sub
