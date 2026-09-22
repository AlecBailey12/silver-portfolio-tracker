VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    ' Restore normal dashboard color
    Sheet1.Range("PriceValueCells").Interior.Color = RGB(221, 232, 247)

    ' Launch the watcher.
    ' The watcher will immediately refresh because this workbook is active.
    AppleScriptTask _
        "Silver Spot Price Auto Refresh.scpt", _
        "startWatcher", _
        ""

End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)

    ' Quit the external watcher
    AppleScriptTask _
        "Silver Spot Price Auto Refresh.scpt", _
        "stopWatcher", _
        ""

End Sub

