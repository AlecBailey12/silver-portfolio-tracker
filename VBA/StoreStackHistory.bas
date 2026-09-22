Attribute VB_Name = "StoreStackHistory"
Sub StoreStackValue()

    Dim sourceWs As Worksheet
    Dim helperWs As Worksheet
    Dim lastRow As Long
    Dim todayRow As Variant
    Dim lastStoredDate As Date
    Dim lastStoredValue As Double
    Dim lastStoredSpotPrice As Double
    Dim fillDate As Date

    Set sourceWs = ThisWorkbook.Worksheets("Silver Stack Sheet")
    Set helperWs = ThisWorkbook.Worksheets("Helper Sheet")

    ' Find the last existing history row
    lastRow = helperWs.Cells(helperWs.Rows.Count, "B").End(xlUp).Row

    If lastRow >= 3 Then

        lastStoredDate = helperWs.Cells(lastRow, "B").Value
        lastStoredValue = helperWs.Cells(lastRow, "C").Value
        lastStoredSpotPrice = helperWs.Cells(lastRow, "D").Value

        ' Fill any dates missed since the workbook was last opened.
        ' Carry forward both the last stack value and spot price.
        fillDate = lastStoredDate + 1

        Do While fillDate < Date

            lastRow = lastRow + 1

            helperWs.Cells(lastRow, "B").Value = fillDate
            helperWs.Cells(lastRow, "C").Value = lastStoredValue
            helperWs.Cells(lastRow, "D").Value = lastStoredSpotPrice

            fillDate = fillDate + 1

        Loop

    End If

    ' Check whether today already has a row
    todayRow = Application.Match(CLng(Date), helperWs.Range("B:B"), 0)

    If IsError(todayRow) Then

        lastRow = helperWs.Cells(helperWs.Rows.Count, "B").End(xlUp).Row + 1

        If lastRow < 3 Then lastRow = 3

        helperWs.Cells(lastRow, "B").Value = Date
        helperWs.Cells(lastRow, "C").Value = sourceWs.Range("StackValue").Value
        helperWs.Cells(lastRow, "D").Value = sourceWs.Range("SilverSpotPrice").Value

    Else

        ' Keep today's values current as the silver price changes
        helperWs.Cells(todayRow, "C").Value = sourceWs.Range("StackValue").Value
        helperWs.Cells(todayRow, "D").Value = sourceWs.Range("SilverSpotPrice").Value

    End If

End Sub
