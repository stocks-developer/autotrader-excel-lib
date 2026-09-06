Attribute VB_Name = "AutoTraderClient"
' *****************************************************************************
'
' The bulk-order driver for the direct (HTTP) module.
'
' This is the direct replacement for the AutoTraderClient module inside the
' ready-made bulk-order workbooks. The macro names are deliberately unchanged
' -- PlaceOrders, PlaceOrdersManual, StartTimer, StopTimer -- so the buttons
' already on those sheets keep working after you swap the module.
'
' TO MIGRATE A BULK-ORDER WORKBOOK
'   1. Press Alt+F11 to open the VBA editor.
'   2. Right-click the old AutoTraderClient module and choose Remove. Say No
'      when it offers to export.
'   3. File -> Import File, and import these three:
'         AutoTraderConfig.bas
'         AutoTraderWebDirect.bas
'         AutoTraderClientDirect.bas
'   4. Put your API key in AutoTraderConfig.
'   5. Save the workbook.
'
' WHAT CHANGES ON THE SHEET
'
' Placing now waits for each broker's answer instead of dropping a line in a
' file, so a run over many accounts takes longer and reports what actually
' happened. The summary at the end counts three outcomes, and the middle one
' matters: an order the broker never confirmed may still be live.
'
' Version: 1.0
'
' *****************************************************************************

Option Explicit

Private TIMER_RUNNING As Boolean

Sub PlaceOrders()

    Dim wOrders As Worksheet
    Dim rAccounts As Range
    Dim rOrders As Range
    Dim temp As String
    Dim OrderRow As Long, AccountRow As Long
    Dim placed As Long, failed As Long, unconfirmed As Long
    Dim orderId As String

    If Not IsAutoTraderReady() Then
        Exit Sub
    End If

    Set wOrders = Sheets("orders")
    Set rAccounts = Sheets("accounts").Range("A:A")
    Set rOrders = wOrders.Range("A:M")

    For OrderRow = 2 To rOrders.rows.count
        temp = rOrders.Cells(RowIndex:=OrderRow, columnIndex:="A").value

        If (Trim(temp) = "") Then
            If (OrderRow = 2) Then
                MsgBox "Please enter orders in <orders> sheet.", vbCritical, "No orders found"
                Exit Sub
            End If

            Exit For
        End If

        For AccountRow = 2 To rAccounts.rows.count
            temp = rAccounts.Cells(RowIndex:=AccountRow, columnIndex:="A").value

            If (Trim(temp) = "") Then
                If (AccountRow = 2) Then
                    MsgBox "Please add accounts in <accounts> sheet.", vbCritical, "Accounts missing"
                    Exit Sub
                End If

                Exit For
            End If

            Dim Variety As String: Variety = Trim(wOrders.Cells(OrderRow, 1))
            Dim pseudoAccount As String: pseudoAccount = Trim(temp)
            Dim Exchange As String: Exchange = Trim(wOrders.Cells(OrderRow, 2))
            Dim Symbol As String: Symbol = Trim(wOrders.Cells(OrderRow, 3))
            Dim TradeType As String: TradeType = Trim(wOrders.Cells(OrderRow, 4))
            Dim OrderType As String: OrderType = Trim(wOrders.Cells(OrderRow, 6))
            Dim ProductType As String: ProductType = Trim(wOrders.Cells(OrderRow, 5))
            Dim Quantity As Long: Quantity = wOrders.Cells(OrderRow, 7)
            Dim Price As Double: Price = wOrders.Cells(OrderRow, 8)
            Dim TriggerPrice As Double: TriggerPrice = wOrders.Cells(OrderRow, 9)
            Dim Target As Double: Target = wOrders.Cells(OrderRow, 10)
            Dim Stoploss As Double: Stoploss = wOrders.Cells(OrderRow, 11)
            Dim TrailingStoploss As Double: TrailingStoploss = wOrders.Cells(OrderRow, 12)
            Dim DisclosedQuantity As Long: DisclosedQuantity = 0
            Dim Validity As String: Validity = "DAY"
            Dim Amo As Boolean: Amo = wOrders.Cells(OrderRow, 13)
            Dim StrategyId As Integer: StrategyId = -1
            Dim Comments As String: Comments = ""

            orderId = PlaceOrderAdvanced(Variety, pseudoAccount, Exchange, _
                Symbol, TradeType, OrderType, ProductType, Quantity, Price, _
                TriggerPrice, Target, Stoploss, TrailingStoploss, DisclosedQuantity, _
                Validity, Amo, StrategyId, Comments)

            If orderId = AT_UNCONFIRMED Then
                unconfirmed = unconfirmed + 1
            ElseIf orderId = "" Then
                failed = failed + 1
            Else
                placed = placed + 1
            End If

        Next

    Next

    ShowSummary placed, failed, unconfirmed

End Sub

' Says what happened, and says it differently when something is unresolved.
'
' A plain "done" would hide the one outcome that needs a person: an order the
' broker never confirmed is not a failure, and re-running the sheet could
' double it.
Private Sub ShowSummary(placed As Long, failed As Long, unconfirmed As Long)

    Dim Message As String

    Message = "Placed: " & placed & vbNewLine & _
              "Not placed: " & failed & vbNewLine & _
              "Not confirmed: " & unconfirmed

    If unconfirmed > 0 Then
        Message = Message & vbNewLine & vbNewLine & _
            unconfirmed & " order(s) reached the broker but were never confirmed. " & _
            "They may be live." & vbNewLine & vbNewLine & _
            "Check your order book before running this again."

        MsgBox Message, vbExclamation, "Check your order book"
    ElseIf failed > 0 Then
        MsgBox Message, vbExclamation, "Some orders were not placed"
    Else
        MsgBox Message, vbInformation, "Success"
    End If

End Sub

Sub PlaceOrdersManual()

    If MsgBox("Are you sure you want to place these orders?", vbYesNo) = vbNo Then
        Exit Sub
    End If

    PlaceOrders

End Sub

Sub StopTimer()
    TIMER_RUNNING = False
    Worksheets("timer").Buttons("timer-button").Caption = "Start Timer"
End Sub

Sub PlaceOrdersOnTime()

    Dim RemainingTime As Double, Deadline As Double

    Deadline = Worksheets("timer").Range("B1")
    RemainingTime = Deadline - (Now - Date)

    If TIMER_RUNNING = False Then
        Exit Sub
    End If

    If RemainingTime > (-1 / 86400) Then
        Worksheets("timer").Range("B2").value = Format(RemainingTime, "h:mm:ss")
        Application.OnTime Now + 1 / 86400, "PlaceOrdersOnTime"
    Else
        PlaceOrders
        StopTimer
    End If

End Sub

Sub StartTimer()

    Dim RemainingTime As Double, Deadline As Double

    Deadline = Worksheets("timer").Range("B1")
    RemainingTime = Deadline - (Now - Date)

    If TIMER_RUNNING Then
        StopTimer
        MsgBox "Timer has been stopped."
        Exit Sub
    ElseIf RemainingTime <= (-1 / 86400) Then
        MsgBox "Time has already expired, please correct time."
        Exit Sub
    End If

    TIMER_RUNNING = True
    Worksheets("timer").Buttons("timer-button").Caption = "Stop Timer"
    PlaceOrdersOnTime

End Sub
