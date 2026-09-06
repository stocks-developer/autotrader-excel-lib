Attribute VB_Name = "AutoTraderClient"
' *****************************************************************************
'
' {title}
'
' The bulk-order driver for this workbook. It reads the sheet and sends each
' order straight to the platform over HTTP, waiting for the answer.
'
' The macro names are unchanged -- PlaceOrders, PlaceOrdersManual, StartTimer,
' StopTimer -- so the buttons already on these sheets keep working.
'
' This workbook needs three modules to run, and this is one of them. The other
' two are AutoTraderConfig (where your API key goes) and AutoTraderWebDirect.
'
' WHAT CHANGED FOR YOU
'
' Placing now waits for each broker's answer instead of dropping a line in a
' file, so a run over many accounts takes longer and tells you what actually
' happened. The summary at the end counts three outcomes, and the middle one
' matters: an order the broker never confirmed may still be live, so check your
' order book before running the sheet again.
'
' Version: 1.0
'
' *****************************************************************************
Option Explicit

Dim TIMER_RUNNING As Boolean

Sub PlaceOrders()

    ' Outcome tally for the summary at the end. An order the broker never
    ' confirmed is neither a success nor a failure and is counted apart.
    Dim placed As Long, failed As Long, unconfirmed As Long
    Dim orderId As String
    
    If Not IsAutoTraderReady() Then
        Exit Sub
    End If
    
    Dim wOrders As Worksheet: Set wOrders = Sheets("orders")
    Dim rOrders As Range: Set rOrders = wOrders.Range("A:M")
    Dim OrderRow As Long
    Dim Temp As String
        
    For OrderRow = 2 To rOrders.Rows.Count
        Temp = rOrders.Cells(RowIndex:=OrderRow, columnIndex:="A").Value
        
        If (Trim(Temp) = "") Then
            If (OrderRow = 2) Then
                MsgBox "Please enter orders in <orders> sheet.", vbCritical, "No orders found"
                Exit Sub
            End If
            
            Exit For
        End If
            
        Dim Variety As String: Variety = Trim(wOrders.Cells(OrderRow, 1))
        Dim pseudoAccount As String: pseudoAccount = Trim(wOrders.Cells(OrderRow, 14))
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

    ShowSummary placed, failed, unconfirmed

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
        Worksheets("timer").Range("B2").Value = Format(RemainingTime, "h:mm:ss")
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
