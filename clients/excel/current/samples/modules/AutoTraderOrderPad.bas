Attribute VB_Name = "AutoTraderOrderPad"
' *****************************************************************************
'
' AutoTrader Web -- Order Pad
'
' Place orders into one account or many, and see what happened to each one.
'
' The four older bulk-order tools could only send. Nothing came back, so
' "did it work?" had to be answered somewhere else. This one writes the
' broker's order id, its status and its reason next to every order it sends.
'
' WHAT EACH SHEET IS FOR
'
'   Orders     what you want to place. One row per order.
'              Leave Account blank and the order goes to every account
'              ticked on the Accounts sheet.
'   Accounts   who to place it for, and how much each one gets.
'   Results    what happened. Written by the tool, newest run at the bottom.
'   Configuration  your API key, allocation mode, and a time to place at.
'   Help       how to use it.
'
' THE BUTTONS
'
'   Check accounts    reads each ticked account and shows its free margin.
'                     Places nothing. Run this first -- it catches a wrong
'                     account name before you send fifty orders, not after.
'   Place orders      asks you to confirm, then sends.
'   Refresh results   re-reads the status of everything already in Results.
'                     Your broker's order book takes a few seconds to catch
'                     up, so a status read straight after placing is often
'                     still the old one. Press this again a few seconds later.
'   Arm timer         places automatically at the time on Configuration.
'
' Macro names are prefixed OrderPad* so they cannot collide with the
' AutoTraderClientDirect module, which also defines PlaceOrders.
'
' *****************************************************************************

Option Explicit

' --- sheets -----------------------------------------------------------------
Private Const SH_ORDERS As String = "Orders"
Private Const SH_ACCOUNTS As String = "Accounts"
Private Const SH_RESULTS As String = "Results"
Private Const SH_SETTINGS As String = "Configuration"

' --- Orders sheet columns ---------------------------------------------------
Private Const O_ACCOUNT As Long = 1
Private Const O_EXCHANGE As Long = 2
Private Const O_SYMBOL As Long = 3
Private Const O_SIDE As Long = 4
Private Const O_QTY As Long = 5
Private Const O_ORDERTYPE As Long = 6
Private Const O_PRICE As Long = 7
Private Const O_TRIGGER As Long = 8
Private Const O_PRODUCT As Long = 9
Private Const O_VARIETY As Long = 10
Private Const O_VALIDITY As Long = 11
Private Const O_TARGET As Long = 12
Private Const O_STOPLOSS As Long = 13
Private Const O_TRAIL As Long = 14

' --- Accounts sheet columns -------------------------------------------------
Private Const A_ACCOUNT As Long = 1
Private Const A_INCLUDE As Long = 2
Private Const A_QTY As Long = 3
Private Const A_WEIGHT As Long = 4

' --- Results sheet columns --------------------------------------------------
Private Const R_TIME As Long = 1
Private Const R_ROW As Long = 2
Private Const R_ACCOUNT As Long = 3
Private Const R_SYMBOL As Long = 4
Private Const R_SIDE As Long = 5
Private Const R_QTY As Long = 6
Private Const R_ORDERID As Long = 7
Private Const R_STATUS As Long = 8
Private Const R_REASON As Long = 9

Private Const FIRST_DATA_ROW As Long = 3

' Allocation modes, as written on the Configuration sheet.
Private Const MODE_SAME As String = "SAME"
Private Const MODE_PER_ACCOUNT As String = "PER ACCOUNT"
Private Const MODE_SPLIT As String = "SPLIT"

Private timerArmed As Boolean
Private timerAt As Date
Private tickAt As Date
Private tickScheduled As Boolean

' The reason the last order failed, so the summary can name it once
' rather than sending the reader to the sheet to find out.
Private lastFailureReason As String


' =============================================================================
' Check accounts -- reads only, places nothing.
' =============================================================================
Public Sub OrderPadCheckAccounts()

    Dim accounts As Variant
    Dim quantities As Variant
    Dim weights As Variant
    Dim i As Long
    Dim report As String
    Dim available As Double
    Dim reachable As Long

    If Not Ready(True, False) Then Exit Sub
    If Not LoadAccounts(accounts, quantities, weights) Then Exit Sub

    report = "Account check" & vbCrLf & vbCrLf

    For i = LBound(accounts) To UBound(accounts)
        available = GetMarginAvailableAll(CStr(accounts(i)))
        If available > 0 Then
            reachable = reachable + 1
            report = report & accounts(i) & "   free margin " & _
                Format$(available, "#,##0.00") & vbCrLf
        Else
            report = report & accounts(i) & "   nothing came back" & vbCrLf
        End If
    Next i

    report = report & vbCrLf & reachable & " of " & _
        (UBound(accounts) - LBound(accounts) + 1) & " account(s) answered."

    If reachable < (UBound(accounts) - LBound(accounts) + 1) Then
        report = report & vbCrLf & vbCrLf & _
            "An account that does not answer is usually a name that does not " & _
            "match your pseudo account, or a key that was not accepted. " & _
            "Nothing was placed."
    End If

    MsgBox report, vbInformation, "Check accounts"

End Sub


' =============================================================================
' Place every order on the Orders sheet.
' =============================================================================
Public Sub OrderPadPlaceOrders()

    Dim orders As Worksheet
    Dim accounts As Variant
    Dim quantities As Variant
    Dim weights As Variant
    Dim lastRow As Long
    Dim r As Long
    Dim orderCount As Long
    Dim sendCount As Long
    Dim fanCount As Long
    Dim namedCount As Long
    Dim accountCount As Long
    Dim headline As String
    Dim mode As String

    If Not Ready(True, True) Then Exit Sub

    Set orders = ThisWorkbook.Worksheets(SH_ORDERS)
    lastRow = LastUsedRow(orders, O_SYMBOL)

    If lastRow < FIRST_DATA_ROW Then
        MsgBox "There are no orders on the " & SH_ORDERS & " sheet.", _
            vbExclamation, "Nothing to place"
        Exit Sub
    End If

    If Not LoadAccounts(accounts, quantities, weights) Then Exit Sub

    mode = AllocationMode()

    ' Count first, so the confirmation can say exactly what is about to happen.
    ' Rows are counted in two groups, because they behave differently: a row
    ' with the Account column blank is COPIED to every ticked account, and a row
    ' naming an account goes only there.
    accountCount = UBound(accounts) - LBound(accounts) + 1

    For r = FIRST_DATA_ROW To lastRow
        If Len(Trim$(CStr(orders.Cells(r, O_SYMBOL).Value))) > 0 Then
            orderCount = orderCount + 1
            If Len(Trim$(CStr(orders.Cells(r, O_ACCOUNT).Value))) > 0 Then
                namedCount = namedCount + 1
                sendCount = sendCount + 1
            Else
                fanCount = fanCount + 1
                sendCount = sendCount + accountCount
            End If
        End If
    Next r

    ' Say where the total came from. "2 order rows will be sent as 4 real
    ' orders" leaves the reader to work out for themselves that it was two
    ' accounts, and that is the one number they most need to check.
    If namedCount = 0 Then
        headline = Plural(fanCount, "order") & " will be copied to " & _
            Plural(accountCount, "account") & " [Total orders = " & sendCount & "]"
    ElseIf fanCount = 0 Then
        headline = Plural(namedCount, "order") & " will be placed, each into the " & _
            "account named on its own row [Total orders = " & sendCount & "]"
    Else
        headline = Plural(orderCount, "order") & ", " & fanCount & " of them copied to " & _
            Plural(accountCount, "account") & " [Total orders = " & sendCount & "]"
    End If

    If MsgBox(headline & vbCrLf & vbCrLf & _
        "These are REAL orders." & vbCrLf & _
        "Allocation: " & mode & vbCrLf & vbCrLf & _
        "Place them now?", vbYesNo + vbExclamation, "Place real orders") <> vbYes Then
        Exit Sub
    End If

    SendAll orders, lastRow, accounts, quantities, weights, mode

End Sub


' =============================================================================
' Re-read the status of everything already in Results.
' =============================================================================
Public Sub OrderPadRefresh()

    If Not Ready(False, False) Then Exit Sub

    Dim results As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim orderId As String
    Dim account As String
    Dim handle As String
    Dim updated As Long

    Set results = ThisWorkbook.Worksheets(SH_RESULTS)
    lastRow = LastUsedRow(results, R_ORDERID)

    If lastRow < FIRST_DATA_ROW Then
        MsgBox "There is nothing in " & SH_RESULTS & " yet.", _
            vbInformation, "Nothing to refresh"
        Exit Sub
    End If

    For r = FIRST_DATA_ROW To lastRow
        orderId = Trim$(CStr(results.Cells(r, R_ORDERID).Value))
        account = Trim$(CStr(results.Cells(r, R_ACCOUNT).Value))

        If Len(orderId) > 0 And Len(account) > 0 Then
            handle = AtFindOrder(account, orderId)
            If AtFound(handle) Then
                results.Cells(r, R_STATUS).Value = AtText(handle, "STATUS")
                results.Cells(r, R_REASON).Value = AtText(handle, "STATUSMESSAGE")
                updated = updated + 1
            End If
        End If
    Next r

    MsgBox updated & " row(s) updated from your broker's order book." & vbCrLf & vbCrLf & _
        "An order that has just been placed can take a few seconds to appear. " & _
        "If a status looks stale, press this again shortly.", _
        vbInformation, "Refresh results"

End Sub


' =============================================================================
' Timer
' =============================================================================
Public Sub OrderPadArmTimer()

    If Not Ready(True, True) Then Exit Sub

    Dim at As Date
    Dim problem As String

    If Not PlaceAtTime(at, problem) Then
        MsgBox problem, vbExclamation, "Time not understood"
        Exit Sub
    End If

    timerAt = Date + at

    If timerAt <= Now Then
        MsgBox "That time has already passed today.", vbExclamation, "Time has passed"
        Exit Sub
    End If

    Application.OnTime timerAt, "AutoTraderOrderPad.OrderPadTimerFired"
    timerArmed = True

    CancelTick
    OrderPadTick

    MsgBox "Armed. Orders will be placed at " & Format$(timerAt, "hh:nn:ss") & _
        "." & vbCrLf & vbCrLf & "Leave this workbook open.", _
        vbInformation, "Timer armed"

End Sub


Public Sub OrderPadCancelTimer()

    If Not timerArmed Then
        MsgBox "The timer is not armed.", vbInformation, "Nothing to cancel"
        Exit Sub
    End If

    On Error Resume Next
    Application.OnTime timerAt, "AutoTraderOrderPad.OrderPadTimerFired", , False
    CancelTick
    SetCancelCaption "Cancel timer"
    On Error GoTo 0

    timerArmed = False
    MsgBox "Timer cancelled. Nothing will be placed automatically.", _
        vbInformation, "Timer cancelled"

End Sub


' Called by Application.OnTime. Places WITHOUT asking -- arming the timer was
' the confirmation, and a dialog nobody is there to answer would stop it.
Public Sub OrderPadTimerFired()

    Dim orders As Worksheet
    Dim accounts As Variant
    Dim quantities As Variant
    Dim weights As Variant
    Dim lastRow As Long

    timerArmed = False
    CancelTick
    SetCancelCaption "Cancel timer"

    Set orders = ThisWorkbook.Worksheets(SH_ORDERS)
    lastRow = LastUsedRow(orders, O_SYMBOL)

    If lastRow < FIRST_DATA_ROW Then Exit Sub
    If Not Ready(True, True) Then Exit Sub
    If Not LoadAccounts(accounts, quantities, weights) Then Exit Sub

    SendAll orders, lastRow, accounts, quantities, weights, AllocationMode()

End Sub


' =============================================================================
' Clear the Results sheet.
' =============================================================================
Public Sub OrderPadClearResults()

    Dim results As Worksheet
    Dim lastRow As Long

    Set results = ThisWorkbook.Worksheets(SH_RESULTS)
    lastRow = LastUsedRow(results, R_TIME)

    If lastRow < FIRST_DATA_ROW Then Exit Sub

    If MsgBox("Clear " & (lastRow - FIRST_DATA_ROW + 1) & " row(s) from " & _
        SH_RESULTS & "?" & vbCrLf & vbCrLf & _
        "This only clears the sheet. It does not cancel anything at your broker.", _
        vbYesNo + vbQuestion, "Clear results") <> vbYes Then Exit Sub

    results.Range(results.Cells(FIRST_DATA_ROW, 1), _
        results.Cells(lastRow, R_REASON)).ClearContents

End Sub


' =============================================================================
' The work
' =============================================================================
Private Sub SendAll(orders As Worksheet, lastRow As Long, accounts As Variant, _
    quantities As Variant, weights As Variant, mode As String)

    Dim r As Long
    Dim i As Long
    Dim account As String
    Dim qty As Long
    Dim placed As Long
    Dim failed As Long
    Dim unconfirmed As Long
    Dim orderId As String

    Application.ScreenUpdating = False

    ' Every failed order would otherwise raise its own dialog. Fifty rows of
    ' those is not a report, and a run that needs fifty clicks is a run nobody
    ' finishes. Off for the loop; the reason goes into Results instead, where
    ' it can be kept. On Error makes sure it is turned back on even if a row
    ' blows up halfway through -- leaving it off would silence the dialogs for
    ' everything the user does afterwards.
    On Error GoTo Restore
    lastFailureReason = ""
    AtSetQuiet True

    For r = FIRST_DATA_ROW To lastRow

        If Len(Trim$(CStr(orders.Cells(r, O_SYMBOL).Value))) = 0 Then GoTo NextOrder

        account = Trim$(CStr(orders.Cells(r, O_ACCOUNT).Value))

        If Len(account) > 0 Then
            ' A named account on the row overrides the fan-out entirely.
            orderId = SendOne(orders, r, account, OrderQuantity(orders, r))
            Tally orderId, placed, failed, unconfirmed
        Else
            For i = LBound(accounts) To UBound(accounts)
                qty = QuantityFor(orders, r, mode, i, accounts, quantities, weights)
                If qty > 0 Then
                    orderId = SendOne(orders, r, CStr(accounts(i)), qty)
                    Tally orderId, placed, failed, unconfirmed
                End If
            Next i
        End If

NextOrder:
    Next r

Restore:
    AtSetQuiet False
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "Stopped after " & (placed + failed + unconfirmed) & " order(s)." & _
            vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            "Whatever was sent is on the " & SH_RESULTS & " sheet.", _
            vbCritical, "Stopped early"
        Exit Sub
    End If

    OrderPadShowSummary placed, failed, unconfirmed

End Sub


Private Function SendOne(orders As Worksheet, r As Long, account As String, _
    qty As Long) As String

    Dim orderId As String
    Dim variety As String
    Dim validity As String
    Dim handle As String
    Dim status As String
    Dim reason As String

    variety = TextOr(orders.Cells(r, O_VARIETY).Value, VARIETY_REGULAR)
    validity = TextOr(orders.Cells(r, O_VALIDITY).Value, VALIDITY_DEFAULT)

    orderId = PlaceOrderAdvanced(variety, account, _
        Trim$(CStr(orders.Cells(r, O_EXCHANGE).Value)), _
        Trim$(CStr(orders.Cells(r, O_SYMBOL).Value)), _
        Trim$(CStr(orders.Cells(r, O_SIDE).Value)), _
        Trim$(CStr(orders.Cells(r, O_ORDERTYPE).Value)), _
        Trim$(CStr(orders.Cells(r, O_PRODUCT).Value)), _
        qty, _
        NumberOr(orders.Cells(r, O_PRICE).Value), _
        NumberOr(orders.Cells(r, O_TRIGGER).Value), _
        NumberOr(orders.Cells(r, O_TARGET).Value), _
        NumberOr(orders.Cells(r, O_STOPLOSS).Value), _
        NumberOr(orders.Cells(r, O_TRAIL).Value), _
        0, validity, False, -1, "Order Pad")

    ' The library knows exactly why a request failed, and with dialogs off it
    ' would otherwise be lost. Lead with its own words -- "the key was not
    ' accepted" is a different problem from "nothing reached your broker", and
    ' a generic sentence in this column hides which one happened.
    If orderId = AT_UNCONFIRMED Then
        status = "UNCONFIRMED"
        reason = TrimmedOr(AtLastMessage(), "Your broker did not confirm it.") & _
            " It MAY be live. Check your order book before sending this again."
    ElseIf Len(orderId) = 0 Then
        status = "NOT PLACED"
        reason = TrimmedOr(AtLastMessage(), "Nothing reached your broker.")
        lastFailureReason = reason
    Else
        ' Asked this soon the book is often still the old one, so whatever comes
        ' back is a first look, not the last word. Refresh results re-reads it.
        handle = AtFindOrder(account, orderId)
        If AtFound(handle) Then
            status = AtText(handle, "STATUS")
            reason = AtText(handle, "STATUSMESSAGE")
        Else
            status = "PLACED"
            reason = "Not in the order book yet. Press Refresh results shortly."
        End If
    End If

    AppendResult r, account, _
        Trim$(CStr(orders.Cells(r, O_SYMBOL).Value)), _
        Trim$(CStr(orders.Cells(r, O_SIDE).Value)), _
        qty, orderId, status, reason

    SendOne = orderId

End Function


Private Sub AppendResult(orderRow As Long, account As String, symbol As String, _
    side As String, qty As Long, orderId As String, status As String, reason As String)

    Dim results As Worksheet
    Dim r As Long

    Set results = ThisWorkbook.Worksheets(SH_RESULTS)
    r = LastUsedRow(results, R_TIME) + 1
    If r < FIRST_DATA_ROW Then r = FIRST_DATA_ROW

    ' Set the format BEFORE writing, every time. The sheet ships with these
    ' formats already on the columns, but a user can clear formatting, insert a
    ' column or paste over the sheet, and an order log that has quietly lost its
    ' seconds -- or turned an order id into 2.6091E+13 -- is worse than useless
    ' when it is the thing you are trying to reconcile against a broker.
    results.Cells(r, R_TIME).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    results.Cells(r, R_TIME).Value = Now
    results.Cells(r, R_ROW).Value = orderRow
    results.Cells(r, R_ACCOUNT).Value = account
    results.Cells(r, R_SYMBOL).Value = symbol
    results.Cells(r, R_SIDE).Value = side
    results.Cells(r, R_QTY).Value = qty
    ' An order id is an identifier made of digits, not a number. As a number
    ' Excel shows 26091000455886 as 2.6091E+13, and past fifteen digits it would
    ' round it into something that is not the id at all.
    results.Cells(r, R_ORDERID).NumberFormat = "@"
    results.Cells(r, R_ORDERID).Value = orderId
    results.Cells(r, R_STATUS).Value = status
    results.Cells(r, R_REASON).Value = reason

End Sub


' =============================================================================
' Allocation
' =============================================================================
Private Function QuantityFor(orders As Worksheet, r As Long, mode As String, _
    i As Long, accounts As Variant, quantities As Variant, weights As Variant) As Long

    Dim total As Long
    Dim weightSum As Double
    Dim share As Double

    total = OrderQuantity(orders, r)

    Select Case mode

        Case MODE_PER_ACCOUNT
            QuantityFor = CLng(quantities(i))

        Case MODE_SPLIT
            weightSum = SumOf(weights)
            If weightSum <= 0 Then
                QuantityFor = 0
            Else
                share = total * (CDbl(weights(i)) / weightSum)
                QuantityFor = CLng(Int(share))
                ' Whatever rounding dropped goes to the first account, so the
                ' pieces always add back up to the quantity you asked for.
                If i = LBound(accounts) Then
                    QuantityFor = QuantityFor + (total - WholeSplitTotal(total, weights))
                End If
            End If

        Case Else
            QuantityFor = total

    End Select

End Function


Private Function WholeSplitTotal(total As Long, weights As Variant) As Long

    Dim i As Long
    Dim weightSum As Double
    Dim sum As Long

    weightSum = SumOf(weights)
    If weightSum <= 0 Then Exit Function

    For i = LBound(weights) To UBound(weights)
        sum = sum + CLng(Int(total * (CDbl(weights(i)) / weightSum)))
    Next i

    WholeSplitTotal = sum

End Function


Private Function SumOf(values As Variant) As Double

    Dim i As Long
    Dim sum As Double

    For i = LBound(values) To UBound(values)
        sum = sum + CDbl(values(i))
    Next i

    SumOf = sum

End Function


' =============================================================================
' Sheet helpers
' =============================================================================
Private Function LoadAccounts(ByRef accounts As Variant, ByRef quantities As Variant, _
    ByRef weights As Variant) As Boolean

    Dim sheet As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim n As Long
    Dim accList() As String
    Dim qtyList() As Double
    Dim wList() As Double

    Set sheet = ThisWorkbook.Worksheets(SH_ACCOUNTS)
    lastRow = LastUsedRow(sheet, A_ACCOUNT)

    ReDim accList(0 To 0)
    ReDim qtyList(0 To 0)
    ReDim wList(0 To 0)

    For r = FIRST_DATA_ROW To lastRow
        If Len(Trim$(CStr(sheet.Cells(r, A_ACCOUNT).Value))) > 0 Then
            If Included(sheet.Cells(r, A_INCLUDE).Value) Then
                ReDim Preserve accList(0 To n)
                ReDim Preserve qtyList(0 To n)
                ReDim Preserve wList(0 To n)
                accList(n) = Trim$(CStr(sheet.Cells(r, A_ACCOUNT).Value))
                qtyList(n) = NumberOr(sheet.Cells(r, A_QTY).Value)
                wList(n) = NumberOr(sheet.Cells(r, A_WEIGHT).Value)
                n = n + 1
            End If
        End If
    Next r

    If n = 0 Then
        MsgBox "No accounts are ticked on the " & SH_ACCOUNTS & " sheet." & vbCrLf & _
            vbCrLf & "Put Y in the Include column beside the ones you want.", _
            vbExclamation, "No accounts"
        LoadAccounts = False
        Exit Function
    End If

    accounts = accList
    quantities = qtyList
    weights = wList
    LoadAccounts = True

End Function


Private Function Included(value As Variant) As Boolean

    Dim text As String

    text = UCase$(Trim$(CStr(value)))
    Included = (text = "Y" Or text = "YES" Or text = "TRUE" Or text = "1")

End Function


Private Function AllocationMode() As String

    Dim mode As String

    mode = UCase$(Trim$(Setting("Allocation mode")))

    If mode = MODE_PER_ACCOUNT Or mode = MODE_SPLIT Then
        AllocationMode = mode
    Else
        AllocationMode = MODE_SAME
    End If

End Function


Private Function Setting(settingName As String) As String
    Setting = Trim$(CStr(SettingRaw(settingName)))
End Function


' The value EXACTLY as Excel holds it, not turned into a string first. A cell
' that holds a time holds a NUMBER, and stringifying it throws away the only
' reliable way to read it.
Private Function SettingRaw(settingName As String) As Variant

    Dim sheet As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set sheet = ThisWorkbook.Worksheets(SH_SETTINGS)
    lastRow = LastUsedRow(sheet, 1)

    For r = FIRST_DATA_ROW To lastRow
        If UCase$(Trim$(CStr(sheet.Cells(r, 1).Value))) = UCase$(settingName) Then
            SettingRaw = sheet.Cells(r, 2).Value
            Exit Function
        End If
    Next r

End Function


' =============================================================================
' Read "Place at" whatever form it is in.
'
' A TIME TYPED INTO A CELL IS NOT TEXT. Type 9:20:00 AM and Excel recognises
' it, stores the NUMBER 0.3888... and formats it as a time. So the cell hands
' back a Date or a fraction of a day, and the old code turned that into a string
' and fed it to TimeValue -- which then depended on the machine's locale for
' whether it parsed. It failed on a perfectly good time.
'
' Take the hour, minute and second off the VALUE, and only fall back to parsing
' text when the cell really does contain text.
' =============================================================================
Private Function PlaceAtTime(ByRef at As Date, ByRef problem As String) As Boolean

    Dim raw As Variant
    Dim text As String
    Dim d As Date

    raw = SettingRaw("Place at")
    text = Trim$(CStr(raw))

    If Len(text) = 0 Then
        problem = "Put a time in ""Place at"" on the " & SH_SETTINGS & " sheet first."
        Exit Function
    End If

    On Error Resume Next

    If IsDate(raw) Then
        d = CDate(raw)
    ElseIf IsNumeric(raw) Then
        d = CDate(CDbl(raw))
    Else
        d = CDate(TimeValue(text))
    End If

    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        problem = "That time could not be read. Type it as 09:20:00, or as 9:20 AM."
        Exit Function
    End If

    On Error GoTo 0

    at = TimeSerial(Hour(d), Minute(d), Second(d))
    PlaceAtTime = True

End Function


Private Function OrderQuantity(orders As Worksheet, r As Long) As Long
    OrderQuantity = CLng(NumberOr(orders.Cells(r, O_QTY).Value))
End Function


Private Function LastUsedRow(sheet As Worksheet, column As Long) As Long
    LastUsedRow = sheet.Cells(sheet.Rows.Count, column).End(xlUp).Row
End Function


Private Function NumberOr(value As Variant) As Double

    Dim text As String

    text = Trim$(CStr(value))
    If Len(text) = 0 Or Not IsNumeric(text) Then
        NumberOr = 0
    Else
        NumberOr = CDbl(text)
    End If

End Function


Private Function TrimmedOr(value As String, fallback As String) As String

    If Len(Trim$(value)) = 0 Then
        TrimmedOr = fallback
    Else
        TrimmedOr = Trim$(value)
    End If

End Function


Private Function TextOr(value As Variant, fallback As String) As String

    Dim text As String

    text = Trim$(CStr(value))
    If Len(text) = 0 Then
        TextOr = fallback
    Else
        TextOr = UCase$(text)
    End If

End Function


Private Sub Tally(orderId As String, ByRef placed As Long, ByRef failed As Long, _
    ByRef unconfirmed As Long)

    If orderId = AT_UNCONFIRMED Then
        unconfirmed = unconfirmed + 1
    ElseIf Len(orderId) = 0 Then
        failed = failed + 1
    Else
        placed = placed + 1
    End If

End Sub


' Says what happened, in one line, in words.
'
' It used to read "0 placed, 2 not placed, 0 unconfirmed" -- three numbers, two
' of them usually zero, leaving the reader to work out their own answer. A
' count is only mentioned when there is something to say about it, and when
' every order failed for the same reason that reason is on the dialog, because
' the commonest case by far is one thing wrong with the whole run.
Private Sub OrderPadShowSummary(placed As Long, failed As Long, unconfirmed As Long)

    Dim total As Long
    Dim message As String
    Dim title As String
    Dim icon As Long

    total = placed + failed + unconfirmed

    If unconfirmed > 0 Then
        message = Plural(unconfirmed, "order") & " reached your broker and " & _
            IIf(unconfirmed = 1, "was", "were") & " never confirmed." & vbCrLf & vbCrLf & _
            "They MAY be live. Check your order book before sending them again."
        title = "Check your order book"
        icon = vbExclamation

    ElseIf failed = 0 Then
        message = IIf(total = 1, "Order placed.", "All " & total & " orders placed.")
        title = "Done"
        icon = vbInformation

    ElseIf placed = 0 Then
        message = IIf(total = 1, "The order was not placed.", _
            "None of the " & total & " orders were placed.")
        title = "Nothing was placed"
        icon = vbExclamation

    Else
        message = placed & " of " & total & " orders placed. " & _
            Plural(failed, "order") & " did not go through."
        title = "Some orders did not go through"
        icon = vbExclamation
    End If

    If failed > 0 And Len(lastFailureReason) > 0 Then
        message = message & vbCrLf & vbCrLf & lastFailureReason
    End If

    message = message & vbCrLf & vbCrLf & "Full detail is on the " & SH_RESULTS & " sheet."

    MsgBox message, icon, title

End Sub


Private Function Plural(count As Long, noun As String) As String
    Plural = count & " " & noun & IIf(count = 1, "", "s")
End Function


' =============================================================================
' Checks that run before anything is sent
'
' The point is to fail on the SHEET, in plain words, rather than at the broker
' one row at a time. A missing API key, an unticked Accounts sheet or a blank
' Side are all things we can see without asking anybody's broker, and finding
' out about them fifty rows into a run is the worst possible time.
'
' Everything is collected and shown ONCE. Reporting the first problem, being
' fixed, then being told about the second is how a new user gives up.
' =============================================================================
Private Function Ready(needAccounts As Boolean, needOrders As Boolean) As Boolean

    Dim problems As String
    Dim part As String

    part = ApiKeyProblem()
    If Len(part) > 0 Then problems = problems & part & vbCrLf & vbCrLf

    If needAccounts Then
        part = AccountProblem()
        If Len(part) > 0 Then problems = problems & part & vbCrLf & vbCrLf
    End If

    If needOrders Then
        part = OrderProblems()
        If Len(part) > 0 Then problems = problems & part & vbCrLf & vbCrLf
    End If

    If Len(problems) = 0 Then
        Ready = True
    Else
        MsgBox RTrim$(problems) & vbCrLf & _
            "Nothing has been sent. Fix these and run it again.", _
            vbExclamation, "Not ready yet"
    End If

End Function


Private Function ApiKeyProblem() As String

    Dim key As String

    key = Trim$(AtApiKey())

    If Len(key) = 0 Then
        ApiKeyProblem = "There is no API key. Put yours in the Value column on " & _
            "the " & SH_SETTINGS & " sheet."
    ElseIf InStr(1, key, "<API_KEY>", vbTextCompare) > 0 Then
        ApiKeyProblem = "The API key is still the placeholder it was shipped " & _
            "with. Put your own key on the " & SH_SETTINGS & " sheet. You will " & _
            "find it in AutoTrader Web under Settings > Security."
    End If

End Function


Private Function AccountProblem() As String

    Dim sheet As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim ticked As Long
    Dim mode As String
    Dim name As String
    Dim blanks As String

    Set sheet = ThisWorkbook.Worksheets(SH_ACCOUNTS)
    lastRow = LastUsedRow(sheet, A_ACCOUNT)

    If lastRow < FIRST_DATA_ROW Then
        AccountProblem = "There are no accounts on the " & SH_ACCOUNTS & " sheet."
        Exit Function
    End If

    mode = AllocationMode()

    For r = FIRST_DATA_ROW To lastRow
        name = Trim$(CStr(sheet.Cells(r, A_ACCOUNT).Value))
        If Len(name) > 0 Then
            If Included(sheet.Cells(r, A_INCLUDE).Value) Then
                ticked = ticked + 1
                If mode = MODE_PER_ACCOUNT Then
                    If NumberOr(sheet.Cells(r, A_QTY).Value) <= 0 Then
                        blanks = blanks & vbCrLf & "  " & name & " has no Quantity."
                    End If
                ElseIf mode = MODE_SPLIT Then
                    If NumberOr(sheet.Cells(r, A_WEIGHT).Value) <= 0 Then
                        blanks = blanks & vbCrLf & "  " & name & " has no Weight."
                    End If
                End If
            End If
        End If
    Next r

    If ticked = 0 Then
        AccountProblem = "No account is ticked on the " & SH_ACCOUNTS & " sheet. " & _
            "Put Y in the Include column beside each one you want to trade."
    ElseIf Len(blanks) > 0 Then
        AccountProblem = "Allocation mode is " & mode & ", so every ticked " & _
            "account needs its own number:" & blanks
    End If

End Function


Private Function OrderProblems() As String

    Dim sheet As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim found As Long
    Dim list As String
    Dim problem As String

    Set sheet = ThisWorkbook.Worksheets(SH_ORDERS)
    lastRow = LastUsedRow(sheet, O_SYMBOL)

    If lastRow < FIRST_DATA_ROW Then
        OrderProblems = "There are no orders on the " & SH_ORDERS & " sheet."
        Exit Function
    End If

    For r = FIRST_DATA_ROW To lastRow
        problem = OrderRowProblem(sheet, r)
        If Len(problem) > 0 Then
            found = found + 1
            ' Eight is enough to see the pattern. A hundred is a wall of text.
            If found <= 8 Then
                list = list & vbCrLf & "  Row " & r & ": " & problem
            End If
        End If
    Next r

    If found > 8 Then
        list = list & vbCrLf & "  ... and " & (found - 8) & " more like these."
    End If

    If Len(list) > 0 Then
        OrderProblems = "These order rows are not ready:" & list
    End If

End Function


Private Function OrderRowProblem(sheet As Worksheet, r As Long) As String

    Dim missing As String
    Dim orderType As String
    Dim variety As String

    If Len(Trim$(CStr(sheet.Cells(r, O_EXCHANGE).Value))) = 0 Then
        missing = AddItem(missing, "Exchange")
    End If
    If Len(Trim$(CStr(sheet.Cells(r, O_SYMBOL).Value))) = 0 Then
        missing = AddItem(missing, "Symbol")
    End If
    If Len(Trim$(CStr(sheet.Cells(r, O_SIDE).Value))) = 0 Then
        missing = AddItem(missing, "Side")
    End If
    If Len(Trim$(CStr(sheet.Cells(r, O_ORDERTYPE).Value))) = 0 Then
        missing = AddItem(missing, "Order Type")
    End If
    If Len(Trim$(CStr(sheet.Cells(r, O_PRODUCT).Value))) = 0 Then
        missing = AddItem(missing, "Product")
    End If
    If OrderQuantity(sheet, r) <= 0 Then
        missing = AddItem(missing, "Quantity")
    End If

    If Len(missing) > 0 Then
        OrderRowProblem = missing & " missing"
        Exit Function
    End If

    orderType = UCase$(Trim$(CStr(sheet.Cells(r, O_ORDERTYPE).Value)))
    variety = UCase$(Trim$(CStr(sheet.Cells(r, O_VARIETY).Value)))

    If orderType = "LIMIT" Then
        If NumberOr(sheet.Cells(r, O_PRICE).Value) <= 0 Then
            OrderRowProblem = "a LIMIT order needs a Price"
        End If
    ElseIf orderType = "STOP_LOSS" Or orderType = "SL_MARKET" Then
        If NumberOr(sheet.Cells(r, O_TRIGGER).Value) <= 0 Then
            OrderRowProblem = "a " & orderType & " order needs a Trigger Price"
        End If
    End If

    If Len(OrderRowProblem) = 0 And variety = "BO" Then
        If NumberOr(sheet.Cells(r, O_TARGET).Value) <= 0 And _
            NumberOr(sheet.Cells(r, O_STOPLOSS).Value) <= 0 Then
            OrderRowProblem = "a bracket order needs a Target or a Stoploss"
        End If
    End If

End Function


Private Function AddItem(list As String, item As String) As String

    If Len(list) = 0 Then
        AddItem = item
    Else
        AddItem = list & ", " & item
    End If

End Function
' =============================================================================
' The armed timer counts itself down on the Cancel timer button.
'
' An armed timer that shows nothing is unnerving. There is no way to tell it
' took, and the only way to find out is to wait and see whether real orders
' appear. The button you would press to stop it is exactly where you look when
' you are wondering, so the countdown lives there:
'
'     Cancel timer (04:37)
'
' OrderPadTick is Public because Application.OnTime can only call a Public Sub.
' It is not meant to be run by hand and is not on the Help sheet, in the same
' way OrderPadTimerFired is not.
' =============================================================================
Public Sub OrderPadTick()

    Dim secondsLeft As Double

    tickScheduled = False

    If Not timerArmed Then
        SetCancelCaption "Cancel timer"
        Exit Sub
    End If

    secondsLeft = (timerAt - Now) * 86400#

    If secondsLeft <= 0 Then
        SetCancelCaption "Cancel timer"
        Exit Sub
    End If

    SetCancelCaption "Cancel timer (" & CountdownText(secondsLeft) & ")"
    ScheduleTick

End Sub


' Everything below is cosmetic, so none of it is allowed to stop a timer. A
' missing button, a renamed sheet or a workbook that never had buttons in it
' must all be survivable: the orders matter, the caption does not.
Private Sub ScheduleTick()

    tickAt = Now + TimeSerial(0, 0, 1)

    On Error Resume Next
    Application.OnTime tickAt, "AutoTraderOrderPad.OrderPadTick"
    If Err.Number = 0 Then tickScheduled = True
    Err.Clear
    On Error GoTo 0

End Sub


Private Sub CancelTick()

    If Not tickScheduled Then Exit Sub

    On Error Resume Next
    Application.OnTime tickAt, "AutoTraderOrderPad.OrderPadTick", , False
    Err.Clear
    On Error GoTo 0

    tickScheduled = False

End Sub


Private Sub SetCancelCaption(caption As String)

    On Error Resume Next
    ThisWorkbook.Worksheets(SH_ORDERS).Buttons("btnOrderPadCancelTimer").caption = caption
    Err.Clear
    On Error GoTo 0

End Sub


Private Function CountdownText(seconds As Double) As String

    Dim s As Long

    s = CLng(Int(seconds))

    If s >= 3600 Then
        CountdownText = CStr(s \ 3600) & "h " & Format$((s Mod 3600) \ 60, "00") & "m"
    Else
        CountdownText = Format$(s \ 60, "00") & ":" & Format$(s Mod 60, "00")
    End If

End Function


' =============================================================================
' Called from Workbook_BeforeClose.
'
' A pending Application.OnTime can make Excel REOPEN a workbook by itself to
' run the macro it promised to run. With a timer ticking every second that is
' close to guaranteed, and a trading workbook reopening on its own is alarming
' whether or not it places anything. Clear both schedules on the way out, and
' say nothing while doing it -- a dialog during close is its own problem.
' =============================================================================
Public Sub OrderPadStopTimersQuietly()

    On Error Resume Next

    If timerArmed Then
        Application.OnTime timerAt, "AutoTraderOrderPad.OrderPadTimerFired", , False
    End If
    timerArmed = False

    CancelTick
    SetCancelCaption "Cancel timer"

    Err.Clear
    On Error GoTo 0

End Sub
