Attribute VB_Name = "SpotPriceRefresh"
Sub RefreshSilverPrice()

    Dim SilverPrice As String
    Dim OldPrice As Double
    Dim NewPrice As Double
    Dim OldCents As Long
    Dim NewCents As Long

    On Error GoTo RefreshError

    OldPrice = Sheet1.Range("SilverSpotPrice").Value

    SilverPrice = AppleScriptTask( _
        "Silver Spot Price Auto Refresh.scpt", _
        "getSilverPrice", _
        "" _
    )

    NewPrice = CDbl(SilverPrice)

    ' Convert both prices to displayed cents
    OldCents = CLng(WorksheetFunction.Round(OldPrice * 100, 0))
    NewCents = CLng(WorksheetFunction.Round(NewPrice * 100, 0))

    ' If the displayed price has not changed, do nothing
    If NewCents = OldCents Then Exit Sub

    ' Only update the workbook when the displayed price changes
    Sheet1.Range("SilverSpotPrice").Value = NewPrice

    Application.Calculate
    
    ' Update today's Helper Sheet values
    StoreStackValue

    If NewCents > OldCents Then
        FlashPriceCells RGB(198, 239, 206)
    Else
        FlashPriceCells RGB(255, 199, 206)
    End If

    Exit Sub

RefreshError:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation

End Sub


Sub FlashPriceCells(FlashColor As Long)

    Dim OriginalColor As Long
    Dim StartTime As Double
    Dim StepTime As Double
    Dim i As Integer

    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    Dim r As Long, g As Long, b As Long

    OriginalColor = Sheet1.Range("SilverSpotPrice").Interior.Color

    Sheet1.Range("PriceValueCells").Interior.Color = FlashColor

    ' Hold flash for 1.5 seconds
    StartTime = Timer
    Do While Timer < StartTime + 1.5
        DoEvents
    Loop

    r1 = FlashColor Mod 256
    g1 = (FlashColor \ 256) Mod 256
    b1 = (FlashColor \ 65536) Mod 256

    r2 = OriginalColor Mod 256
    g2 = (OriginalColor \ 256) Mod 256
    b2 = (OriginalColor \ 65536) Mod 256

    ' Fade back over roughly 1 second
    For i = 1 To 20

        r = r1 + (r2 - r1) * i / 20
        g = g1 + (g2 - g1) * i / 20
        b = b1 + (b2 - b1) * i / 20
        Sheet1.Range("PriceValueCells").Interior.Color = RGB(r, g, b)

        StepTime = Timer
        Do While Timer < StepTime + 0.05
            DoEvents
        Loop

    Next i

    Sheet1.Range("PriceValueCells").Interior.Color = OriginalColor

End Sub


Public Sub WatcherRefresh()

    If Application.ActiveWorkbook Is ThisWorkbook Then
        RefreshSilverPrice
    End If

End Sub
