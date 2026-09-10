Attribute VB_Name = "AutoTraderWebDirect"
' *****************************************************************************
'
' AutoTrader Web -- direct (HTTP) module for Excel.
' DO NOT MODIFY THIS FILE
' Version: 1.0
'
' The same functions as the AutoTraderWeb add-in, talking to AutoTrader Web
' over the internet instead of through files. PlaceOrder, GetOrderStatus,
' GetPositionMtm and the rest keep their names, their arguments and their
' meaning, so an existing sheet keeps working.
'
' SETUP: put your API key in the AutoTraderConfig module. That is all -- there
' is no folder to configure and nothing else to keep running.
'
' FOUR DIFFERENCES worth knowing before you switch:
'
' 1. PlaceOrder now waits for the broker's answer and returns the BROKER's
'    order id, not an internally generated one. That is the id you pass to
'    GetOrderStatus, ModifyOrder and CancelOrder, and it is the same id you see
'    in your broker's order book.
'
' 2. PlaceOrder can return AT_UNCONFIRMED. That means the order reached the
'    broker and the broker never confirmed it, so it may be live. Do not send
'    it again -- look at your order book.
'
' 3. The portfolio summary functions (GetPortfolioMtm, GetPortfolioPnl, the
'    position and order counts) are not part of this version. They were
'    calculated on your own computer by a program this module exists to do
'    without. Use the position and order functions instead.
'
' 4. Holdings functions are NEW. Excel never had them.
'
' A NOTE ON WHERE YOU CALL THESE FROM
'
' Reading a value asks our server for it, so a getter used in hundreds of cells
' that recalculate together would be slow. The module keeps a short-lived copy
' of each portfolio (see AT_TTL_ below), which makes a sheet full of getters
' cost one request rather than hundreds. Placing an order is different: it
' waits for your broker, so call it from a button or a macro, never from a
' worksheet formula.
'
' *****************************************************************************

Option Explicit

Const CONTACT_SUPPORT As String = "Please take a screenshot of this message and mail to help@stocksdeveloper.in"

' Command verbs. The three ...BY_ID ones exist for a client with no publisher
' id to give, which is exactly what this module is.
Const PLACE_ORDER_CMD As String = "PLACE_ORDER"
Const CANCEL_ALL_ORDERS_CMD As String = "CANCEL_ALL_ORDERS"
Const SQUARE_OFF_POSITION_CMD As String = "SQUARE_OFF_POSITION"
Const SQUARE_OFF_PORTFOLIO_CMD As String = "SQUARE_OFF_PORTFOLIO"
Const MODIFY_ORDER_BY_ID_CMD As String = "MODIFY_ORDER_BY_ID"
Const CANCEL_ORDER_BY_ID_CMD As String = "CANCEL_ORDER_BY_ID"
Const CANCEL_CHILD_ORDER_BY_ID_CMD As String = "CANCEL_CHILD_ORDER_BY_ID"

Public Const EPOCH As Date = #1/1/1970#
Public Const BLANK As String = ""

Public Const VARIETY_REGULAR As String = "REGULAR"
Public Const VARIETY_BO As String = "BO"
Public Const VARIETY_CO As String = "CO"

Public Const VALIDITY_DAY As String = "DAY"
Public Const VALIDITY_IOC As String = "IOC"
Public Const VALIDITY_DEFAULT As String = VALIDITY_DAY

Public Const PRODUCT_INTRADAY As String = "INTRADAY"
Public Const PRODUCT_DELIVERY As String = "DELIVERY"
Public Const PRODUCT_NORMAL As String = "NORMAL"
Public Const PRODUCT_MTF As String = "MTF"

Public Const MARGIN_EQUITY As String = "EQUITY"
Public Const MARGIN_COMMODITY As String = "COMMODITY"
Public Const MARGIN_ALL As String = "ALL"

' Dataset names. These are the paths under /csv as well.
Public Const AT_DS_ORDERS As String = "orders"
Public Const AT_DS_POSITIONS As String = "positions"
Public Const AT_DS_MARGINS As String = "margins"
Public Const AT_DS_HOLDINGS As String = "holdings"

' How a reply is classified. Every response is exactly one of these.
Const AT_HTTP_CSV As String = "CSV"
Const AT_HTTP_EMPTY As String = "EMPTY"
Const AT_HTTP_OK As String = "OK"
Const AT_HTTP_ERROR As String = "ERROR"
Const AT_HTTP_AUTH As String = "AUTH"
Const AT_HTTP_UNREACHABLE As String = "UNREACHABLE"

' What the server says to do about a failure.
Const AT_ACTION_RETRY As String = "retry"
Const AT_ACTION_USER As String = "user"
Const AT_ACTION_CHECK As String = "check"

' First field of every structured reply.
Const AT_SD_ERROR As String = "SD-ERROR"
Const AT_SD_OK As String = "SD-OK"

' First column of every portfolio CSV, and the positive test for "this is
' data". No error reply can begin with it: an error starts with SD-, a rejected
' key with a brace, a gateway page with an angle bracket.
Const AT_CSV_HEADER As String = "PSEUDOACCOUNT"

' Returned by PlaceOrder when the order reached the broker and the broker never
' confirmed it. Deliberately NOT blank: a sheet that reads blank as failure
' would place the order a second time, and there is real money on the other
' side of that.
Public Const AT_UNCONFIRMED As String = "UNCONFIRMED"

Const HEX_DIGITS As String = "0123456789ABCDEF"

' ---------------------------------------------------------------------------
' State
' ---------------------------------------------------------------------------

' False when the request never left the machine at all. Needed because an
' unreachable server and an account with no orders both hand back an empty
' string, and treating a dead connection as "you have no orders" is how a sheet
' ends up showing a flat position that is not flat.
Private AtHttpReached As Boolean

Private AtHttpMessage As String

Private AtHttpAction As String

Private AtHttpResponse As String

Private AtHttpStatusCode As Long

' One entry per account and dataset: the rows, when they were fetched, and how
' that fetch went.
Private AtCache As Object

Private Sub AtEnsureCache()
    If AtCache Is Nothing Then
        Set AtCache = CreateObject("Scripting.Dictionary")
    End If
End Sub

' Clears every cached portfolio. Use it if you want the next read to go to the
' server no matter how recent the last one was.
Public Sub AutoTraderClearCache()
    Set AtCache = Nothing
End Sub

' ---------------------------------------------------------------------------
' Plumbing
' ---------------------------------------------------------------------------

Private Sub AtTrace(Message As String)
    If AT_HTTP_DEBUG Then
        Debug.Print "AutoTrader: " & Message
    End If
End Sub

' Says plainly when the key has not been filled in. Without this the first
' symptom is an order that does not appear, and the reason for it is a
' rejection nobody sees.
Public Function IsAutoTraderReady() As Boolean

    If AT_API_KEY = "<API_KEY>" Or Len(Trim(AT_API_KEY)) = 0 Then
        MsgBox "No API key. Open the AutoTraderConfig module and set AT_API_KEY " & _
            "to the key from your AutoTrader Web account settings.", vbCritical, "AutoTrader"
        IsAutoTraderReady = False
        Exit Function
    End If

    IsAutoTraderReady = True

End Function

' Percent-encodes a value for a form body.
'
' Not optional. Symbols legitimately contain characters that would otherwise
' end the field or start a new one -- M&M is a real NSE symbol, and an
' unencoded ampersand there would silently truncate the order.
Public Function AtUrlEncode(Text As String) As String

    Dim i As Long
    Dim code As Long
    Dim c As String
    Dim result As String

    For i = 1 To Len(Text)
        c = Mid$(Text, i, 1)
        code = AscW(c)

        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or _
           (code >= 97 And code <= 122) Or c = "-" Or c = "_" Or c = "." Or c = "~" Then
            result = result & c
        ElseIf code < 128 And code >= 0 Then
            result = result & "%" & Mid$(HEX_DIGITS, (code \ 16) + 1, 1) & _
                Mid$(HEX_DIGITS, (code Mod 16) + 1, 1)
        Else
            ' Anything outside plain ASCII is encoded byte by byte as UTF-8, so
            ' a comment in another script survives instead of being mangled.
            result = result & AtUrlEncodeUtf8(c)
        End If
    Next i

    AtUrlEncode = result

End Function

Private Function AtUrlEncodeUtf8(c As String) As String

    Dim code As Long
    Dim bytes(0 To 3) As Long
    Dim count As Integer
    Dim i As Integer
    Dim result As String

    code = AscW(c)
    If code < 0 Then code = code + 65536

    If code < 2048 Then
        bytes(0) = 192 + (code \ 64)
        bytes(1) = 128 + (code Mod 64)
        count = 2
    Else
        bytes(0) = 224 + (code \ 4096)
        bytes(1) = 128 + ((code \ 64) Mod 64)
        bytes(2) = 128 + (code Mod 64)
        count = 3
    End If

    For i = 0 To count - 1
        result = result & "%" & Mid$(HEX_DIGITS, (bytes(i) \ 16) + 1, 1) & _
            Mid$(HEX_DIGITS, (bytes(i) Mod 16) + 1, 1)
    Next i

    AtUrlEncodeUtf8 = result

End Function

' Returns the n-th pipe separated field of a structured reply, 1 based.
'
' Split on the pipe rather than the comma because the message field of an error
' routinely contains commas ("RMS: margin shortfall, order rejected"). The
' server guarantees the reverse -- it replaces any pipe inside a message.
Public Function AtPipeField(Line As String, index As Integer) As String

    Dim parts() As String
    Dim value As String

    parts = Split(Line, "|")

    If index >= 1 And index <= (UBound(parts) - LBound(parts) + 1) Then
        value = parts(LBound(parts) + index - 1)
    Else
        AtPipeField = ""
        Exit Function
    End If

    ' The last field is trimmed because it carries whatever line ending the
    ' reader left on the reply, and that field is the broker's order id.
    '
    ' An untrimmed id is still returned by PlaceOrder and still looks right in a
    ' cell, but every GetOrderStatus/GetOrderPrice call made with it silently
    ' returns blank, because no row matches an id with a line ending stuck to
    ' it.
    '
    ' MSXML hands back the whole body rather than lines, so this is a guard
    ' rather than a fix for an observed fault. VBA's own Trim removes spaces
    ' only, so the line endings have to go first.
    value = Replace(value, vbCr, "")
    value = Replace(value, vbLf, "")

    AtPipeField = Trim$(value)

End Function

' Returns one field of a CSV line, 1 based, honouring quoted fields.
'
' Quotes are not decoration here. The server quotes any field containing a
' comma, and two fields routinely do: an order's status message and its
' comments. Splitting on every comma would shift every column after such a
' field, so GetOrderAveragePrice would quietly return a piece of the broker's
' error text.
Public Function AtCsvField(Line As String, column As Integer) As String

    Dim total As Long
    Dim i As Long
    Dim c As String
    Dim field As Integer
    Dim quoted As Boolean
    Dim value As String

    total = Len(Line)
    field = 1
    quoted = False
    value = ""

    i = 1
    Do While i <= total
        c = Mid$(Line, i, 1)

        If quoted Then
            If c = """" Then
                If i < total And Mid$(Line, i + 1, 1) = """" Then
                    value = value & """"
                    i = i + 1
                Else
                    quoted = False
                End If
            Else
                value = value & c
            End If
        ElseIf c = """" And Len(value) = 0 Then
            quoted = True
        ElseIf c = "," Then
            If field = column Then
                AtCsvField = value
                Exit Function
            End If
            field = field + 1
            value = ""
        Else
            value = value & c
        End If

        i = i + 1
    Loop

    If field = column Then
        AtCsvField = value
    Else
        AtCsvField = ""
    End If

End Function

' Blank-safe converters. The lifted getters use CDbl and CLng directly, which
' is what they always did; these are for the new holdings functions, where a
' symbol you do not hold should read as zero rather than raise #VALUE.
Public Function AtToDouble(Text As String) As Double
    If Len(Trim(Text)) = 0 Or Not IsNumeric(Text) Then
        AtToDouble = 0
    Else
        AtToDouble = CDbl(Text)
    End If
End Function

Public Function AtToLong(Text As String) As Long
    If Len(Trim(Text)) = 0 Or Not IsNumeric(Text) Then
        AtToLong = 0
    Else
        AtToLong = CLng(Text)
    End If
End Function

' Sends a POST and returns the whole response body.
Private Function AtHttpPost(Path As String, PostBody As String, TimeoutMs As Long) As String

    Dim http As Object

    AtHttpReached = False
    AtHttpStatusCode = -1
    AtHttpPost = ""

    On Error GoTo Failed

    ' ServerXMLHTTP rather than XMLHTTP: it uses WinHTTP, which does not cache
    ' responses and lets us set a real timeout. A cached portfolio read would
    ' be worse than a slow one.
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts TimeoutMs, TimeoutMs, TimeoutMs, TimeoutMs
    http.Open "POST", AT_BASE_URL & Path, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.send PostBody

    AtHttpReached = True
    AtHttpStatusCode = http.Status
    AtHttpPost = http.responseText

    Exit Function

Failed:
    AtTrace "request failed: " & Err.Number & " " & Err.Description
    AtHttpReached = False

End Function

' Decides what kind of reply this is, by positive identification.
'
' A blank body is a dataset with no rows. It is NOT a valid answer to a
' command: every command is answered SD-OK or SD-ERROR, so a blank reply to one
' means the answer was lost on the way back, and the order may well have been
' placed. That is the one case that must never be retried automatically.
Private Function AtClassify(Response As String, IsCommand As Boolean) As String

    If Not AtHttpReached Then
        AtClassify = AT_HTTP_UNREACHABLE
    ElseIf InStr(1, Response, AT_SD_ERROR) = 1 Then
        AtClassify = AT_HTTP_ERROR
    ElseIf InStr(1, Response, AT_SD_OK) = 1 Then
        AtClassify = AT_HTTP_OK
    ElseIf InStr(1, Response, "{") = 1 Then
        AtClassify = AT_HTTP_AUTH
    ElseIf IsCommand Then
        AtClassify = AT_HTTP_UNREACHABLE
    ElseIf Len(Response) = 0 Then
        AtClassify = AT_HTTP_EMPTY
    ElseIf InStr(1, Response, AT_CSV_HEADER) = 1 Then
        AtClassify = AT_HTTP_CSV
    Else
        AtClassify = AT_HTTP_UNREACHABLE
    End If

End Function

' Turns a classification into the message a user should see, and remembers the
' action so a caller can tell "try later" from "fix something".
'
' The one judgement here is what an unreachable server means, and it differs by
' request. A read that did not arrive can simply be asked for again. A command
' that did not arrive may in fact have arrived, with only the answer lost.
Private Sub AtRecordStatus(Status As String, Response As String, IsCommand As Boolean)

    If Status = AT_HTTP_ERROR Then
        AtHttpAction = AtPipeField(Response, 3)
        AtHttpMessage = AtPipeField(Response, 4)
        If Len(AtHttpMessage) = 0 Then AtHttpMessage = Response

    ElseIf Status = AT_HTTP_AUTH Then
        AtHttpAction = AT_ACTION_USER
        AtHttpMessage = "API key was not accepted. Check AT_API_KEY in the AutoTraderConfig module."

    ElseIf Status = AT_HTTP_UNREACHABLE Then
        If IsCommand Then
            AtHttpAction = AT_ACTION_CHECK
            AtHttpMessage = "No usable reply from " & AT_BASE_URL & _
                ". The request may still have gone through -- check your order book."
        Else
            AtHttpAction = AT_ACTION_RETRY
            AtHttpMessage = "Could not reach " & AT_BASE_URL & ". Check the internet connection."
        End If

    Else
        AtHttpAction = ""
        AtHttpMessage = ""
    End If

End Sub

' ---------------------------------------------------------------------------
' The cache
' ---------------------------------------------------------------------------
'
' A sheet may call twenty different Get... functions on one recalculation. Each
' is answered from the copy held here, so a recalculation costs one request per
' dataset instead of twenty -- and the server allows roughly one portfolio
' request per second per account, so a sheet doing otherwise would spend its
' whole allowance re-asking for data it already had.

Private Function AtTtlFor(dataset As String) As Long
    Select Case dataset
        Case AT_DS_ORDERS
            AtTtlFor = AT_TTL_ORDERS
        Case AT_DS_POSITIONS
            AtTtlFor = AT_TTL_POSITIONS
        Case AT_DS_MARGINS
            AtTtlFor = AT_TTL_MARGINS
        Case Else
            AtTtlFor = AT_TTL_HOLDINGS
    End Select
End Function

' Asks the server for one dataset and replaces what is cached for it.
'
' On any failure the previous rows are LEFT ALONE and only the status changes.
' A sheet mid-position should not suddenly show an empty order book because one
' request timed out.
Private Sub AtFetchDataset(pseudoAccount As String, dataset As String)

    Dim PostBody As String
    Dim Response As String
    Dim Status As String
    Dim body As String
    Dim cut As Long
    Dim rows As Variant

    AtEnsureCache

    PostBody = "api-key=" & AtUrlEncode(AT_API_KEY) & _
        "&pseudoAccount=" & AtUrlEncode(pseudoAccount)

    AtTrace "reading " & dataset & " for " & pseudoAccount

    Response = AtHttpPost("/csv/" & dataset, PostBody, AT_HTTP_TIMEOUT_READ)
    Status = AtClassify(Response, False)
    AtRecordStatus Status, Response, False

    If Status = AT_HTTP_CSV Then
        ' Separate the header row from the records, so row 1 is the first real
        ' record and the column numbers match the add-in exactly.
        '
        ' The header itself is KEPT rather than thrown away. It names every
        ' column, which is what lets a lookup ask for "QUANTITY" instead of
        ' counting to seven. Discarding it is what let ReadHoldingColumn match
        ' the broker symbol column while every caller passed an independent
        ' symbol -- a silent wrong answer that reading the code would not show.
        body = Replace(Response, vbCr, "")
        cut = InStr(1, body, vbLf)

        If cut = 0 Then
            rows = Split("", vbLf)
            ' Header only: still worth keeping, so a name can be resolved
            ' against an empty dataset without another request.
            AtCache(dataset & "|" & pseudoAccount & "|H") = body
            AtCache(dataset & "|" & pseudoAccount & "|N") = 0
        Else
            AtCache(dataset & "|" & pseudoAccount & "|H") = Left$(body, cut - 1)
            body = Mid$(body, cut + 1)
            Do While Len(body) > 0 And Right$(body, 1) = vbLf
                body = Left$(body, Len(body) - 1)
            Loop

            If Len(body) = 0 Then
                AtCache(dataset & "|" & pseudoAccount & "|N") = 0
            Else
                rows = Split(body, vbLf)
                AtCache(dataset & "|" & pseudoAccount & "|R") = rows
                AtCache(dataset & "|" & pseudoAccount & "|N") = UBound(rows) - LBound(rows) + 1
            End If
        End If

        AtTrace dataset & ": " & AtCache(dataset & "|" & pseudoAccount & "|N") & " rows"

    ElseIf Status = AT_HTTP_EMPTY Then
        ' A real answer: this account has nothing in this dataset right now.
        AtCache(dataset & "|" & pseudoAccount & "|N") = 0
        AtTrace dataset & ": empty"

    Else
        Debug.Print "AutoTrader: SD-ERR-XL-READ: " & dataset & " for " & pseudoAccount & _
            " failed [" & Status & "] " & AtHttpMessage
    End If

    AtCache(dataset & "|" & pseudoAccount & "|ST") = Status

End Sub

' Makes sure the cached copy of a dataset is recent enough to use.
'
' The timestamp is written whatever the outcome, including failures. Without
' that, a sheet that cannot reach the server would retry on every single
' getter -- hundreds of hanging requests per recalculation, at the moment the
' network is already in trouble.
Private Sub AtEnsureFresh(pseudoAccount As String, dataset As String)

    Dim stampKey As String
    Dim fetchedAt As Date

    AtEnsureCache
    stampKey = dataset & "|" & pseudoAccount & "|TS"

    If AtCache.Exists(stampKey) Then
        fetchedAt = AtCache(stampKey)
        If DateDiff("s", fetchedAt, Now) < AtTtlFor(dataset) Then
            Exit Sub
        End If
    End If

    AtFetchDataset pseudoAccount, dataset

    AtCache(stampKey) = Now

End Sub

Private Function AtRowCountInternal(pseudoAccount As String, dataset As String) As Long
    Dim key As String
    key = dataset & "|" & pseudoAccount & "|N"
    If AtCache.Exists(key) Then
        AtRowCountInternal = AtCache(key)
    Else
        AtRowCountInternal = 0
    End If
End Function

' Reads one column of the row whose key column equals the given value.
'
' Same contract as the add-in's FileReadCsvColumnByRowId: columns are numbered
' from 1, and a row that is not there gives a blank rather than an error, so a
' sheet asking about an order that has not appeared yet behaves as it always
' has.
Public Function AtReadColumn(pseudoAccount As String, dataset As String, _
    keyValue As String, keyColumn As Integer, column As Integer) As String

    Dim rows As Variant
    Dim i As Long
    Dim total As Long

    AtEnsureFresh pseudoAccount, dataset
    AtReadColumn = ""

    total = AtRowCountInternal(pseudoAccount, dataset)
    If total = 0 Then Exit Function

    rows = AtCache(dataset & "|" & pseudoAccount & "|R")

    For i = LBound(rows) To UBound(rows)
        If AtCsvField(CStr(rows(i)), keyColumn) = keyValue Then
            AtReadColumn = AtCsvField(CStr(rows(i)), column)
            Exit Function
        End If
    Next i

End Function

' The same, for rows identified by four columns at once.
'
' A position has no id of its own; it is identified by category, type, exchange
' and symbol together.
Public Function AtReadColumn4(pseudoAccount As String, dataset As String, _
    value1 As String, column1 As Integer, value2 As String, column2 As Integer, _
    value3 As String, column3 As Integer, value4 As String, column4 As Integer, _
    column As Integer) As String

    Dim rows As Variant
    Dim i As Long
    Dim total As Long
    Dim Line As String

    AtEnsureFresh pseudoAccount, dataset
    AtReadColumn4 = ""

    total = AtRowCountInternal(pseudoAccount, dataset)
    If total = 0 Then Exit Function

    rows = AtCache(dataset & "|" & pseudoAccount & "|R")

    For i = LBound(rows) To UBound(rows)
        Line = CStr(rows(i))

        If AtCsvField(Line, column1) = value1 And _
           AtCsvField(Line, column2) = value2 And _
           AtCsvField(Line, column3) = value3 And _
           AtCsvField(Line, column4) = value4 Then

            AtReadColumn4 = AtCsvField(Line, column)
            Exit Function
        End If
    Next i

End Function

' Number of rows currently held for a dataset. Refreshes first.
Public Function AtRowCount(pseudoAccount As String, dataset As String) As Long
    AtEnsureFresh pseudoAccount, dataset
    AtRowCount = AtRowCountInternal(pseudoAccount, dataset)
End Function

' Reads a column from a row by its position, 1 for the first record.
Public Function AtReadRowColumn(pseudoAccount As String, dataset As String, _
    row As Long, column As Integer) As String

    Dim rows As Variant

    AtEnsureFresh pseudoAccount, dataset
    AtReadRowColumn = ""

    If row < 1 Or row > AtRowCountInternal(pseudoAccount, dataset) Then Exit Function

    rows = AtCache(dataset & "|" & pseudoAccount & "|R")
    AtReadRowColumn = AtCsvField(CStr(rows(LBound(rows) + row - 1)), column)

End Function

' ---------------------------------------------------------------------------
' Lookup by name
' ---------------------------------------------------------------------------

' Position of a named column in a dataset, 1 based. 0 when there is no such
' column.
'
' This is the whole point of keeping the header. Every lookup that identifies a
' row goes through a NAME, so a change to the server's column order cannot
' quietly make a getter return the wrong field -- the worst it can do is return
' 0 here, which is visible.
'
' The answer is cached per account and dataset, because a sheet may ask for
' twenty fields on a recalculation and re-scanning the header each time would be
' twenty scans for an answer that never changes within a session.
'
' Names are compared without regard to case, so "quantity" and "QUANTITY" are
' the same column.
Public Function AtColumnOf(pseudoAccount As String, dataset As String, _
    fieldName As String) As Integer

    Dim cacheKey As String
    Dim header As String
    Dim wanted As String
    Dim colName As String
    Dim i As Integer
    Dim found As Integer

    AtEnsureFresh pseudoAccount, dataset

    wanted = UCase$(Trim$(fieldName))
    cacheKey = dataset & "|" & pseudoAccount & "|C" & wanted

    If AtCache.Exists(cacheKey) Then
        ' -1 records "looked and it is not there", so a missing column is not
        ' re-scanned on every call.
        If AtCache(cacheKey) < 0 Then
            AtColumnOf = 0
        Else
            AtColumnOf = AtCache(cacheKey)
        End If
        Exit Function
    End If

    header = ""
    If AtCache.Exists(dataset & "|" & pseudoAccount & "|H") Then
        header = CStr(AtCache(dataset & "|" & pseudoAccount & "|H"))
    End If

    found = 0

    For i = 1 To AT_HTTP_MAX_COLUMNS
        colName = UCase$(Trim$(AtCsvField(header, i)))

        If Len(colName) = 0 Then
            ' Past the end of the header.
            Exit For
        End If

        If colName = wanted Then
            found = i
            Exit For
        End If
    Next i

    ' A miss is only worth remembering when there WAS a header to miss in. An
    ' account with no orders yet gets an empty reply, and an empty reply carries
    ' no header at all -- the server has no rows to write one from. Remembering
    ' "not there" at that moment would keep the column unresolvable for the rest
    ' of the session, so the first order placed afterwards could never be read
    ' back.
    If found = 0 And Len(header) = 0 Then
        AtColumnOf = 0
        Exit Function
    End If

    If found = 0 Then
        AtCache(cacheKey) = -1
    Else
        AtCache(cacheKey) = found
    End If

    AtColumnOf = found

End Function

' One field of one row, addressed by column NAME. Blank when the dataset has no
' such column or the row is out of range.
Public Function AtFieldByName(pseudoAccount As String, dataset As String, _
    row As Long, fieldName As String) As String

    Dim column As Integer
    column = AtColumnOf(pseudoAccount, dataset, fieldName)

    If column < 1 Then
        AtFieldByName = ""
    Else
        AtFieldByName = AtReadRowColumn(pseudoAccount, dataset, row, column)
    End If

End Function

' The first row whose named column equals a value, as a row number. 0 when
' nothing matches.
Public Function AtFindRowByName(pseudoAccount As String, dataset As String, _
    fieldName As String, value As String) As Long

    Dim rows As Variant
    Dim column As Integer
    Dim i As Long

    AtFindRowByName = 0

    column = AtColumnOf(pseudoAccount, dataset, fieldName)
    If column < 1 Then Exit Function
    If AtRowCountInternal(pseudoAccount, dataset) = 0 Then Exit Function

    rows = AtCache(dataset & "|" & pseudoAccount & "|R")

    For i = LBound(rows) To UBound(rows)
        If AtCsvField(CStr(rows(i)), column) = value Then
            AtFindRowByName = i - LBound(rows) + 1
            Exit Function
        End If
    Next i

End Function

' The first row matching TWO named columns at once. 0 when nothing matches.
Public Function AtFindRowByName2(pseudoAccount As String, dataset As String, _
    field1 As String, value1 As String, _
    field2 As String, value2 As String) As Long

    Dim rows As Variant
    Dim col1 As Integer, col2 As Integer
    Dim i As Long
    Dim Line As String

    AtFindRowByName2 = 0

    col1 = AtColumnOf(pseudoAccount, dataset, field1)
    col2 = AtColumnOf(pseudoAccount, dataset, field2)
    If col1 < 1 Or col2 < 1 Then Exit Function
    If AtRowCountInternal(pseudoAccount, dataset) = 0 Then Exit Function

    rows = AtCache(dataset & "|" & pseudoAccount & "|R")

    For i = LBound(rows) To UBound(rows)
        Line = CStr(rows(i))

        If AtCsvField(Line, col1) = value1 And _
           AtCsvField(Line, col2) = value2 Then

            AtFindRowByName2 = i - LBound(rows) + 1
            Exit Function
        End If
    Next i

End Function

' The first row matching FOUR named columns at once.
'
' A position has no id of its own -- it is identified by category, type,
' exchange and symbol together.
Public Function AtFindRowByName4(pseudoAccount As String, dataset As String, _
    field1 As String, value1 As String, _
    field2 As String, value2 As String, _
    field3 As String, value3 As String, _
    field4 As String, value4 As String) As Long

    Dim rows As Variant
    Dim col1 As Integer, col2 As Integer, col3 As Integer, col4 As Integer
    Dim i As Long
    Dim Line As String

    AtFindRowByName4 = 0

    col1 = AtColumnOf(pseudoAccount, dataset, field1)
    col2 = AtColumnOf(pseudoAccount, dataset, field2)
    col3 = AtColumnOf(pseudoAccount, dataset, field3)
    col4 = AtColumnOf(pseudoAccount, dataset, field4)
    If col1 < 1 Or col2 < 1 Or col3 < 1 Or col4 < 1 Then Exit Function
    If AtRowCountInternal(pseudoAccount, dataset) = 0 Then Exit Function

    rows = AtCache(dataset & "|" & pseudoAccount & "|R")

    For i = LBound(rows) To UBound(rows)
        Line = CStr(rows(i))

        If AtCsvField(Line, col1) = value1 And _
           AtCsvField(Line, col2) = value2 And _
           AtCsvField(Line, col3) = value3 And _
           AtCsvField(Line, col4) = value4 Then

            AtFindRowByName4 = i - LBound(rows) + 1
            Exit Function
        End If
    Next i

End Function

' ---------------------------------------------------------------------------
' Commands
' ---------------------------------------------------------------------------

' Sends one command and returns how it went, as one of the AT_HTTP_ values.
' The reply itself is left in AtHttpResponse for a caller that needs the order
' id out of it.
Private Function AtSendCommand(csv As String) As String

    Dim PostBody As String

    PostBody = "api-key=" & AtUrlEncode(AT_API_KEY) & "&command=" & AtUrlEncode(csv)

    AtTrace "command: " & csv

    AtHttpResponse = AtHttpPost("/csv/command", PostBody, AT_HTTP_TIMEOUT_COMMAND)

    AtSendCommand = AtClassify(AtHttpResponse, True)
    AtRecordStatus AtSendCommand, AtHttpResponse, True

    AtTrace "reply: " & AtHttpResponse

End Function

' Runs a command that has nothing to return but success or failure -- modify,
' cancel, square off. Returns True only if the server confirmed it.
Private Function AtRunCommand(csv As String, Description As String) As Boolean

    Dim Status As String
    Status = AtSendCommand(csv)

    If Status = AT_HTTP_OK Then
        AtTrace Description & " done."
        AtRunCommand = True
        Exit Function
    End If

    AtRunCommand = False

    Debug.Print "AutoTrader: SD-ERR-XL-CMD: " & Description & " failed [" & Status & "] " & AtHttpMessage

    If AtHttpAction = AT_ACTION_CHECK Then
        ' The request reached the broker and was never confirmed. Reporting
        ' False is right -- we did not see it succeed -- but a bare False would
        ' let a sheet quietly assume nothing happened, and something may have.
        MsgBox Description & " may still have gone through." & vbNewLine & vbNewLine & _
            AtHttpMessage & vbNewLine & vbNewLine & _
            "Check your order book before repeating it.", vbExclamation, "AutoTrader"
    ElseIf AtHttpAction = AT_ACTION_USER Then
        MsgBox Description & " failed." & vbNewLine & vbNewLine & AtHttpMessage, _
            vbCritical, "AutoTrader"
    End If

End Function

' ---------------------------------------------------------------------------
' Placing orders
' ---------------------------------------------------------------------------

' An advanced function to place orders.
'
' Returns the BROKER's order id on success. Returns AT_UNCONFIRMED when the
' order reached the broker and the broker never confirmed it -- that order may
' be live, so check the order book rather than sending it again. Returns blank
' only when the order was definitely not placed.
Public Function PlaceOrderAdvanced(Variety As String, _
    pseudoAccount As String, _
    Exchange As String, _
    Symbol As String, _
    TradeType As String, _
    OrderType As String, _
    ProductType As String, _
    Quantity As Long, _
    Price As Double, _
    TriggerPrice As Double, _
    Target As Double, _
    Stoploss As Double, _
    TrailingStoploss As Double, _
    DisclosedQuantity As Long, _
    Validity As String, _
    Amo As Boolean, _
    StrategyId As Integer, _
    Comments As String) As String

    Dim cols(0 To 20) As String
    Dim csv As String
    Dim Status As String

    PlaceOrderAdvanced = ""

    If Not IsAutoTraderReady() Then Exit Function

    cols(0) = PLACE_ORDER_CMD
    cols(1) = pseudoAccount
    ' Column three is the publisher id and is deliberately left empty. The
    ' add-in generates one so it can refer to an order it has not yet heard
    ' back about; here the reply carries the broker's own id, so there is
    ' nothing to invent and nothing to reconcile later.
    cols(2) = BLANK
    cols(3) = Variety
    cols(4) = Exchange
    cols(5) = Symbol
    cols(6) = TradeType
    cols(7) = OrderType
    cols(8) = ProductType
    cols(9) = CStr(Quantity)
    cols(10) = CStr(Price)
    cols(11) = CStr(TriggerPrice)
    cols(12) = CStr(Target)
    cols(13) = CStr(Stoploss)
    cols(14) = CStr(TrailingStoploss)
    cols(15) = CStr(DisclosedQuantity)
    cols(16) = Validity
    cols(17) = IIf(Amo, "true", "false")
    ' Milliseconds since epoch. The add-in sent SECONDS here, which the server
    ' read as a moment in January 1970; fixed rather than carried over.
    cols(18) = Format$(CDbl((Now - EPOCH)) * 86400# * 1000#, "0")
    cols(19) = CStr(StrategyId)
    ' A comma here would start a new column and shift everything after it.
    cols(20) = Replace(Comments, ",", ";")

    csv = Join(cols, ",")

    AtTrace "Order csv data: " & csv

    Status = AtSendCommand(csv)

    If Status = AT_HTTP_OK Then
        PlaceOrderAdvanced = AtPipeField(AtHttpResponse, 2)
        Exit Function
    End If

    If AtHttpAction = AT_ACTION_CHECK Then
        PlaceOrderAdvanced = AT_UNCONFIRMED

        MsgBox "Order was sent but the broker did not confirm it." & vbNewLine & vbNewLine & _
            AtHttpMessage & vbNewLine & vbNewLine & _
            "Do not place it again. Check your order book first.", vbExclamation, "AutoTrader"
        Exit Function
    End If

    Debug.Print "AutoTrader: SD-ERR-XL-PLACE: order placement failed [" & Status & "] " & AtHttpMessage

    MsgBox "Order was not placed." & vbNewLine & vbNewLine & AtHttpMessage, _
        vbCritical, "AutoTrader"

End Function

' A function to place regular orders.
Public Function PlaceOrder( _
    pseudoAccount As String, _
    Exchange As String, _
    Symbol As String, _
    TradeType As String, _
    OrderType As String, _
    ProductType As String, _
    Quantity As Long, _
    Price As Double, _
    TriggerPrice As Double) As String

    PlaceOrder = PlaceOrderAdvanced(VARIETY_REGULAR, pseudoAccount, Exchange, _
        Symbol, TradeType, OrderType, ProductType, Quantity, Price, TriggerPrice, _
        0, 0, 0, 0, VALIDITY_DEFAULT, False, -1, BLANK)

End Function

' A function to place bracket orders.
Public Function PlaceBracketOrder( _
    pseudoAccount As String, _
    Exchange As String, _
    Symbol As String, _
    TradeType As String, _
    OrderType As String, _
    Quantity As Long, _
    Price As Double, _
    TriggerPrice As Double, _
    Target As Double, _
    Stoploss As Double, _
    TrailingStoploss As Double) As String

    PlaceBracketOrder = PlaceOrderAdvanced(VARIETY_BO, pseudoAccount, Exchange, _
        Symbol, TradeType, OrderType, PRODUCT_INTRADAY, Quantity, Price, TriggerPrice, _
        Target, Stoploss, TrailingStoploss, 0, VALIDITY_DEFAULT, False, -1, BLANK)

End Function

' A function to place cover orders.
Public Function PlaceCoverOrder( _
    pseudoAccount As String, _
    Exchange As String, _
    Symbol As String, _
    TradeType As String, _
    OrderType As String, _
    Quantity As Long, _
    Price As Double, _
    TriggerPrice As Double) As String

    PlaceCoverOrder = PlaceOrderAdvanced(VARIETY_CO, pseudoAccount, Exchange, _
        Symbol, TradeType, OrderType, PRODUCT_INTRADAY, Quantity, Price, TriggerPrice, _
        0, 0, 0, 0, VALIDITY_DEFAULT, False, -1, BLANK)

End Function

' ---------------------------------------------------------------------------
' Modifying, cancelling and squaring off
' ---------------------------------------------------------------------------
'
' Pass the order id a Place... function returned -- the BROKER's order id.

Public Function ModifyOrder(pseudoAccount As String, _
    orderId As String, _
    OrderType As String, _
    Quantity As Long, _
    Price As Double, _
    TriggerPrice As Double) As Boolean

    Dim csv As String

    ModifyOrder = False
    If Not IsAutoTraderReady() Then Exit Function

    csv = MODIFY_ORDER_BY_ID_CMD & "," & pseudoAccount & "," & orderId & "," & _
        OrderType & "," & CStr(Quantity) & "," & CStr(Price) & "," & CStr(TriggerPrice)

    ModifyOrder = AtRunCommand(csv, "Order modify [" & orderId & "]")

End Function

Public Function ModifyOrderPrice(pseudoAccount As String, _
    orderId As String, Price As Double) As Boolean

    ModifyOrderPrice = ModifyOrder(pseudoAccount, orderId, BLANK, 0, Price, 0)

End Function

Public Function ModifyOrderQuantity(pseudoAccount As String, _
    orderId As String, Quantity As Long) As Boolean

    ModifyOrderQuantity = ModifyOrder(pseudoAccount, orderId, BLANK, Quantity, 0, 0)

End Function

Public Function CancelOrder(pseudoAccount As String, orderId As String) As Boolean

    CancelOrder = False
    If Not IsAutoTraderReady() Then Exit Function

    CancelOrder = AtRunCommand(CANCEL_ORDER_BY_ID_CMD & "," & pseudoAccount & "," & orderId, _
        "Order cancel [" & orderId & "]")

End Function

' Cancels child orders of a bracket or cover order. Useful for exiting one.
Public Function CancelOrderChildren(pseudoAccount As String, orderId As String) As Boolean

    CancelOrderChildren = False
    If Not IsAutoTraderReady() Then Exit Function

    CancelOrderChildren = AtRunCommand( _
        CANCEL_CHILD_ORDER_BY_ID_CMD & "," & pseudoAccount & "," & orderId, _
        "Order cancel children [" & orderId & "]")

End Function

Public Function CancelAllOrders(pseudoAccount As String) As Boolean

    CancelAllOrders = False
    If Not IsAutoTraderReady() Then Exit Function

    CancelAllOrders = AtRunCommand(CANCEL_ALL_ORDERS_CMD & "," & pseudoAccount, _
        "Cancel all open orders [" & pseudoAccount & "]")

End Function

' category - position category (DAY, NET). Pass DAY if you are not sure.
' posType  - position type (MIS, NRML, CNC, BO, CO)
Public Function SquareOffPosition(pseudoAccount As String, _
    category As String, _
    posType As String, _
    independentExchange As String, _
    independentSymbol As String) As Boolean

    Dim csv As String

    SquareOffPosition = False
    If Not IsAutoTraderReady() Then Exit Function

    csv = SQUARE_OFF_POSITION_CMD & "," & pseudoAccount & "," & category & "," & _
        posType & "," & independentExchange & "," & independentSymbol

    SquareOffPosition = AtRunCommand(csv, "Square-off position [" & pseudoAccount & _
        "|" & category & "|" & posType & "|" & independentExchange & "|" & independentSymbol & "]")

End Function

Public Function SquareOffPortfolio(pseudoAccount As String, category As String) As Boolean

    SquareOffPortfolio = False
    If Not IsAutoTraderReady() Then Exit Function

    SquareOffPortfolio = AtRunCommand( _
        SQUARE_OFF_PORTFOLIO_CMD & "," & pseudoAccount & "," & category, _
        "Square-off portfolio [" & pseudoAccount & "|" & category & "]")

End Function

' *****************************************************************************
' ************************ ORDER DETAIL FUNCTIONS - START ***********************
' *****************************************************************************

' Reads orders and returns a column value for the given order id.
'
' Matched on column 4, the BROKER's order id. The add-in matched on column 3,
' the publisher id it generated before sending the order; this module never
' makes one, because the reply to a placement carries the broker's own id.
Public Function ReadOrderColumn(pseudoAccount As String, _
    orderId As String, columnIndex As Integer) As String
    ReadOrderColumn = AtReadColumn(pseudoAccount, AT_DS_ORDERS, orderId, 4, columnIndex)
End Function

' Retrieve order's trading account.
Public Function GetOrderTradingAccount(pseudoAccount As String, _
    orderId As String) As String
    GetOrderTradingAccount = ReadOrderColumn(pseudoAccount, orderId, 2)
End Function

' Retrieve order's trading platform id.
Public Function GetOrderId(pseudoAccount As String, _
    orderId As String) As String
    GetOrderId = ReadOrderColumn(pseudoAccount, orderId, 4)
End Function

' Retrieve order's exchange id.
Public Function GetOrderExchangeId(pseudoAccount As String, _
    orderId As String) As String
    GetOrderExchangeId = ReadOrderColumn(pseudoAccount, orderId, 5)
End Function

' Retrieve order's variety (REGULAR, BO, CO).
Public Function GetOrderVariety(pseudoAccount As String, _
    orderId As String) As String
    GetOrderVariety = ReadOrderColumn(pseudoAccount, orderId, 6)
End Function

' Retrieve order's (platform independent) exchange.
Public Function GetOrderIndependentExchange(pseudoAccount As String, _
    orderId As String) As String
    GetOrderIndependentExchange = ReadOrderColumn(pseudoAccount, orderId, 7)
End Function

' Retrieve order's (platform independent) symbol.
Public Function GetOrderIndependentSymbol(pseudoAccount As String, _
    orderId As String) As String
    GetOrderIndependentSymbol = ReadOrderColumn(pseudoAccount, orderId, 8)
End Function

' Retrieve order's trade type (BUY, SELL).
Public Function GetOrderTradeType(pseudoAccount As String, _
    orderId As String) As String
    GetOrderTradeType = ReadOrderColumn(pseudoAccount, orderId, 9)
End Function

' Retrieve order's order type (LIMIT, MARKET, STOP_LOSS, SL_MARKET).
Public Function GetOrderOrderType(pseudoAccount As String, _
    orderId As String) As String
    GetOrderOrderType = ReadOrderColumn(pseudoAccount, orderId, 10)
End Function

' Retrieve order's product type (INTRADAY, DELIVERY, NORMAL, MTF).
Public Function GetOrderProductType(pseudoAccount As String, _
    orderId As String) As String
    GetOrderProductType = ReadOrderColumn(pseudoAccount, orderId, 11)
End Function

' Retrieve order's quantity.
Public Function GetOrderQuantity(pseudoAccount As String, _
    orderId As String) As Long
    GetOrderQuantity = AtToLong(ReadOrderColumn(pseudoAccount, orderId, 12))
End Function

' Retrieve order's price.
Public Function GetOrderPrice(pseudoAccount As String, _
    orderId As String) As Double
    GetOrderPrice = AtToDouble(ReadOrderColumn(pseudoAccount, orderId, 13))
End Function

' Retrieve order's trigger price.
Public Function GetOrderTriggerPrice(pseudoAccount As String, _
    orderId As String) As Double
    GetOrderTriggerPrice = AtToDouble(ReadOrderColumn(pseudoAccount, orderId, 14))
End Function

' Retrieve order's filled quantity.
Public Function GetOrderFilledQuantity(pseudoAccount As String, _
    orderId As String) As Long
    GetOrderFilledQuantity = AtToLong(ReadOrderColumn(pseudoAccount, orderId, 15))
End Function

' Retrieve order's pending quantity.
Public Function GetOrderPendingQuantity(pseudoAccount As String, _
    orderId As String) As Long
    GetOrderPendingQuantity = AtToLong(ReadOrderColumn(pseudoAccount, orderId, 16))
End Function

' Retrieve order's (platform independent) status.
' (OPEN, COMPLETE, CANCELLED, REJECTED, TRIGGER_PENDING, UNKNOWN)
Public Function GetOrderStatus(pseudoAccount As String, _
    orderId As String) As String
    GetOrderStatus = ReadOrderColumn(pseudoAccount, orderId, 17)
End Function

' Retrieve order's status message or rejection reason.
Public Function GetOrderStatusMessage(pseudoAccount As String, _
    orderId As String) As String
    GetOrderStatusMessage = ReadOrderColumn(pseudoAccount, orderId, 18)
End Function

' Retrieve order's validity (DAY, IOC).
Public Function GetOrderValidity(pseudoAccount As String, _
    orderId As String) As String
    GetOrderValidity = ReadOrderColumn(pseudoAccount, orderId, 19)
End Function

' Retrieve order's average price at which it got traded.
Public Function GetOrderAveragePrice(pseudoAccount As String, _
    orderId As String) As Double
    GetOrderAveragePrice = AtToDouble(ReadOrderColumn(pseudoAccount, orderId, 20))
End Function

' Retrieve order's parent order id. The id of parent bracket or cover order.
Public Function GetOrderParentOrderId(pseudoAccount As String, _
    orderId As String) As String
    GetOrderParentOrderId = ReadOrderColumn(pseudoAccount, orderId, 21)
End Function

' Retrieve order's disclosed quantity.
Public Function GetOrderDisclosedQuantity(pseudoAccount As String, _
    orderId As String) As Long
    GetOrderDisclosedQuantity = AtToLong(ReadOrderColumn(pseudoAccount, orderId, 22))
End Function

' Retrieve order's exchange time as a string (YYYY-MM-DD HH:MM:SS.MILLIS).
Public Function GetOrderExchangeTime(pseudoAccount As String, _
    orderId As String) As String
    GetOrderExchangeTime = ReadOrderColumn(pseudoAccount, orderId, 23)
End Function

' Retrieve order's platform time as a string (YYYY-MM-DD HH:MM:SS.MILLIS).
Public Function GetOrderPlatformTime(pseudoAccount As String, _
    orderId As String) As String
    GetOrderPlatformTime = ReadOrderColumn(pseudoAccount, orderId, 24)
End Function

' Retrieve order's AMO (after market order) flag. (true/false)
Public Function GetOrderAmo(pseudoAccount As String, _
    orderId As String) As Boolean
    Dim flag As String
    flag = ReadOrderColumn(pseudoAccount, orderId, 25)
    GetOrderAmo = (LCase(flag) = "true")
End Function

' Retrieve order's comments.
Public Function GetOrderComments(pseudoAccount As String, _
    orderId As String) As String
    GetOrderComments = ReadOrderColumn(pseudoAccount, orderId, 26)
End Function

' Retrieve order's raw (platform specific) status.
Public Function GetOrderRawStatus(pseudoAccount As String, _
    orderId As String) As String
    GetOrderRawStatus = ReadOrderColumn(pseudoAccount, orderId, 27)
End Function

' Retrieve order's (platform specific) exchange.
Public Function GetOrderExchange(pseudoAccount As String, _
    orderId As String) As String
    GetOrderExchange = ReadOrderColumn(pseudoAccount, orderId, 28)
End Function

' Retrieve order's (platform specific) symbol.
Public Function GetOrderSymbol(pseudoAccount As String, _
    orderId As String) As String
    GetOrderSymbol = ReadOrderColumn(pseudoAccount, orderId, 29)
End Function

' Retrieve order's date (DD-MM-YYYY).
Public Function GetOrderDay(pseudoAccount As String, _
    orderId As String) As String
    GetOrderDay = ReadOrderColumn(pseudoAccount, orderId, 30)
End Function

' Retrieve order's trading platform.
Public Function GetOrderPlatform(pseudoAccount As String, _
    orderId As String) As String
    GetOrderPlatform = ReadOrderColumn(pseudoAccount, orderId, 31)
End Function

' Retrieve order's client id (as received from trading platform).
Public Function GetOrderClientId(pseudoAccount As String, _
    orderId As String) As String
    GetOrderClientId = ReadOrderColumn(pseudoAccount, orderId, 32)
End Function

' Retrieve order's stock broker.
Public Function GetOrderStockBroker(pseudoAccount As String, _
    orderId As String) As String
    GetOrderStockBroker = ReadOrderColumn(pseudoAccount, orderId, 33)
End Function

' Checks whether order is open.
Public Function IsOrderOpen(pseudoAccount As String, _
    orderId As String) As Boolean
    Dim oStatus As String
    oStatus = GetOrderStatus(pseudoAccount, orderId)
    IsOrderOpen = (UCase(oStatus) = "OPEN" Or UCase(oStatus) = "TRIGGER_PENDING")
End Function

' Checks whether order is complete.
Public Function IsOrderComplete(pseudoAccount As String, _
    orderId As String) As Boolean
    Dim oStatus As String
    oStatus = GetOrderStatus(pseudoAccount, orderId)
    IsOrderComplete = UCase(oStatus) = "COMPLETE"
End Function

' Checks whether order is rejected.
Public Function IsOrderRejected(pseudoAccount As String, _
    orderId As String) As Boolean
    Dim oStatus As String
    oStatus = GetOrderStatus(pseudoAccount, orderId)
    IsOrderRejected = UCase(oStatus) = "REJECTED"
End Function

' Checks whether order is cancelled.
Public Function IsOrderCancelled(pseudoAccount As String, _
    orderId As String) As Boolean
    Dim oStatus As String
    oStatus = GetOrderStatus(pseudoAccount, orderId)
    IsOrderCancelled = UCase(oStatus) = "CANCELLED"
End Function

' *****************************************************************************
' ************************ POSITION DETAIL FUNCTIONS - START ***********************
' *****************************************************************************

' Reads positions and returns a column value for the given position id.
' Position id is a combination of category, type, independentExchange & independentSymbol.
Public Function ReadPositionColumnInternal(pseudoAccount As String, _
    category As String, categoryColumnIndex As Integer, _
    posType As String, typeColumnIndex As Integer, _
    independentExchange As String, independentExchangeColumnIndex As Integer, _
    independentSymbol As String, independentSymbolColumnIndex As Integer, _
    columnIndex As Integer) As String

    ReadPositionColumnInternal = AtReadColumn4(pseudoAccount, AT_DS_POSITIONS, _
        category, categoryColumnIndex, posType, typeColumnIndex, _
        independentExchange, independentExchangeColumnIndex, _
        independentSymbol, independentSymbolColumnIndex, columnIndex)
End Function

' Reads positions and returns a column value for the given position id.
Public Function ReadPositionColumn(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String, columnIndex As Integer) As String

    ReadPositionColumn = ReadPositionColumnInternal(pseudoAccount, _
        category, 4, posType, 3, independentExchange, 5, independentSymbol, 6, columnIndex)
End Function

' Retrieve positions's trading account.
Public Function GetPositionTradingAccount(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As String
    GetPositionTradingAccount = ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 2)
End Function

' Retrieve positions's MTM (Mtm calculated by your stock broker).
Public Function GetPositionMtm(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionMtm = AtToDouble(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 7))
End Function

' Retrieve positions's PNL (Pnl calculated by your stock broker).
Public Function GetPositionPnl(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionPnl = AtToDouble(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 8))
End Function

' Retrieve positions's AT PNL (Pnl calculated by AutoTrader Web).
Public Function GetPositionAtPnl(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionAtPnl = AtToDouble(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 31))
End Function

' Retrieve positions's buy quantity.
Public Function GetPositionBuyQuantity(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Long
    GetPositionBuyQuantity = AtToLong(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 9))
End Function

' Retrieve positions's sell quantity.
Public Function GetPositionSellQuantity(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Long
    GetPositionSellQuantity = AtToLong(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 10))
End Function

' Retrieve positions's net quantity.
Public Function GetPositionNetQuantity(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Long
    GetPositionNetQuantity = AtToLong(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 11))
End Function

' Retrieve positions's buy value.
Public Function GetPositionBuyValue(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionBuyValue = AtToDouble(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 12))
End Function

' Retrieve positions's sell value.
Public Function GetPositionSellValue(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionSellValue = AtToDouble(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 13))
End Function

' Retrieve positions's net value.
Public Function GetPositionNetValue(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionNetValue = AtToDouble(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 14))
End Function

' Retrieve positions's buy average price.
Public Function GetPositionBuyAvgPrice(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionBuyAvgPrice = AtToDouble(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 15))
End Function

' Retrieve positions's sell average price.
Public Function GetPositionSellAvgPrice(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionSellAvgPrice = AtToDouble(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 16))
End Function

' Retrieve positions's realised pnl.
Public Function GetPositionRealisedPnl(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionRealisedPnl = AtToDouble(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 17))
End Function

' Retrieve positions's unrealised pnl.
Public Function GetPositionUnrealisedPnl(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionUnrealisedPnl = AtToDouble(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 18))
End Function

' Retrieve positions's overnight quantity.
Public Function GetPositionOvernightQuantity(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Long
    GetPositionOvernightQuantity = AtToLong(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 19))
End Function

' Retrieve positions's multiplier.
Public Function GetPositionMultiplier(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Long
    GetPositionMultiplier = AtToLong(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 20))
End Function

' Retrieve positions's LTP.
Public Function GetPositionLtp(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As Double
    GetPositionLtp = AtToDouble(ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 21))
End Function

' Retrieve positions's (platform specific) exchange.
Public Function GetPositionExchange(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As String
    GetPositionExchange = ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 22)
End Function

' Retrieve positions's (platform specific) symbol.
Public Function GetPositionSymbol(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As String
    GetPositionSymbol = ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 23)
End Function

' Retrieve positions's date (DD-MM-YYYY).
Public Function GetPositionDay(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As String
    GetPositionDay = ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 24)
End Function

' Retrieve positions's trading platform.
Public Function GetPositionPlatform(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As String
    GetPositionPlatform = ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 25)
End Function

' Retrieve positions's account id as received from trading platform.
Public Function GetPositionAccountId(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As String
    GetPositionAccountId = ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 26)
End Function

' Retrieve positions's stock broker.
Public Function GetPositionStockBroker(pseudoAccount As String, _
    category As String, posType As String, independentExchange As String, _
    independentSymbol As String) As String
    GetPositionStockBroker = ReadPositionColumn(pseudoAccount, _
        category, posType, independentExchange, independentSymbol, 28)
End Function

' *****************************************************************************
' ************************ MARGIN DETAIL FUNCTIONS - START ***********************
' *****************************************************************************

' Reads margins and returns a column value for the given margin category.
Public Function ReadMarginColumn(pseudoAccount As String, _
    category As String, columnIndex As Integer) As String
    ReadMarginColumn = AtReadColumn(pseudoAccount, AT_DS_MARGINS, category, 3, columnIndex)
End Function

' Retrieve margin funds.
Public Function GetMarginFunds(pseudoAccount As String, _
    category As String) As Double
    GetMarginFunds = AtToDouble(ReadMarginColumn(pseudoAccount, category, 4))
End Function

' Retrieve margin utilized.
Public Function GetMarginUtilized(pseudoAccount As String, _
    category As String) As Double
    GetMarginUtilized = AtToDouble(ReadMarginColumn(pseudoAccount, category, 5))
End Function

' Retrieve margin available.
Public Function GetMarginAvailable(pseudoAccount As String, _
    category As String) As Double
    GetMarginAvailable = AtToDouble(ReadMarginColumn(pseudoAccount, category, 6))
End Function

' Retrieve margin funds for equity category.
Public Function GetMarginFundsEquity(pseudoAccount As String) As Double
    GetMarginFundsEquity = AtToDouble(ReadMarginColumn(pseudoAccount, MARGIN_EQUITY, 4))
End Function

' Retrieve margin utilized for equity category.
Public Function GetMarginUtilizedEquity(pseudoAccount As String) As Double
    GetMarginUtilizedEquity = AtToDouble(ReadMarginColumn(pseudoAccount, MARGIN_EQUITY, 5))
End Function

' Retrieve margin available for equity category.
Public Function GetMarginAvailableEquity(pseudoAccount As String) As Double
    GetMarginAvailableEquity = AtToDouble(ReadMarginColumn(pseudoAccount, MARGIN_EQUITY, 6))
End Function

' Retrieve margin funds for commodity category.
Public Function GetMarginFundsCommodity(pseudoAccount As String) As Double
    GetMarginFundsCommodity = AtToDouble(ReadMarginColumn(pseudoAccount, MARGIN_COMMODITY, 4))
End Function

' Retrieve margin utilized for commodity category.
Public Function GetMarginUtilizedCommodity(pseudoAccount As String) As Double
    GetMarginUtilizedCommodity = AtToDouble(ReadMarginColumn(pseudoAccount, MARGIN_COMMODITY, 5))
End Function

' Retrieve margin available for commodity category.
Public Function GetMarginAvailableCommodity(pseudoAccount As String) As Double
    GetMarginAvailableCommodity = AtToDouble(ReadMarginColumn(pseudoAccount, MARGIN_COMMODITY, 6))
End Function

' Retrieve margin funds for entire account.
Public Function GetMarginFundsAll(pseudoAccount As String) As Double
    GetMarginFundsAll = AtToDouble(ReadMarginColumn(pseudoAccount, MARGIN_ALL, 4))
End Function

' Retrieve margin utilized for entire account.
Public Function GetMarginUtilizedAll(pseudoAccount As String) As Double
    GetMarginUtilizedAll = AtToDouble(ReadMarginColumn(pseudoAccount, MARGIN_ALL, 5))
End Function

' Retrieve margin available for entire account.
Public Function GetMarginAvailableAll(pseudoAccount As String) As Double
    GetMarginAvailableAll = AtToDouble(ReadMarginColumn(pseudoAccount, MARGIN_ALL, 6))
End Function

' *****************************************************************************
' ************************ HOLDING DETAIL FUNCTIONS - START ***********************
' *****************************************************************************

' Reads holdings and returns a column value for the given symbol.
'
' Matched on the INDEPENDENT symbol, by name.
'
' This used to match column 5, taken from the same column table the AmiBroker
' library was checked against -- which is how the fault travelled. In the file
' the Desktop Client wrote, column 5 was the independent symbol. In the
' server's holdings CSV column 5 is the BROKER symbol -- "IOC-EQ" where the
' caller passes "IOC" -- so every holding getter matched nothing and returned 0
' or blank for an account that did hold the stock. Nothing in the code showed
' it, because 0 is also the honest answer for a stock you do not hold.
'
' Holdings are the one dataset with no single INDEPENDENTSYMBOL column; they
' carry one per exchange. The getters take no exchange, so try NSE and then
' BSE.
Public Function ReadHoldingColumn(pseudoAccount As String, _
    Symbol As String, columnIndex As Integer) As String

    Dim row As Long

    row = AtFindRowByName(pseudoAccount, AT_DS_HOLDINGS, "INDEPENDENTSYMBOLNSE", Symbol)

    If row = 0 Then
        row = AtFindRowByName(pseudoAccount, AT_DS_HOLDINGS, "INDEPENDENTSYMBOLBSE", Symbol)
    End If

    If row = 0 Then
        ReadHoldingColumn = ""
    Else
        ReadHoldingColumn = AtReadRowColumn(pseudoAccount, AT_DS_HOLDINGS, row, columnIndex)
    End If

End Function

' Retrieve holding exchange.
Public Function GetHoldingExchange(pseudoAccount As String, _
    Symbol As String) As String
    GetHoldingExchange = ReadHoldingColumn(pseudoAccount, Symbol, 4)
End Function

' Retrieve holding symbol.
Public Function GetHoldingSymbol(pseudoAccount As String, _
    Symbol As String) As String
    GetHoldingSymbol = ReadHoldingColumn(pseudoAccount, Symbol, 5)
End Function

' Retrieve holding ISIN.
Public Function GetHoldingIsin(pseudoAccount As String, _
    Symbol As String) As String
    GetHoldingIsin = ReadHoldingColumn(pseudoAccount, Symbol, 6)
End Function

' Retrieve holding quantity.
Public Function GetHoldingQuantity(pseudoAccount As String, _
    Symbol As String) As Long
    GetHoldingQuantity = AtToLong(ReadHoldingColumn(pseudoAccount, Symbol, 7))
End Function

' Retrieve holding T1 quantity.
Public Function GetHoldingT1Quantity(pseudoAccount As String, _
    Symbol As String) As Long
    GetHoldingT1Quantity = AtToLong(ReadHoldingColumn(pseudoAccount, Symbol, 8))
End Function

' Retrieve holding P&L.
Public Function GetHoldingPnl(pseudoAccount As String, _
    Symbol As String) As Double
    GetHoldingPnl = AtToDouble(ReadHoldingColumn(pseudoAccount, Symbol, 9))
End Function

' Retrieve holding product.
Public Function GetHoldingProduct(pseudoAccount As String, _
    Symbol As String) As String
    GetHoldingProduct = ReadHoldingColumn(pseudoAccount, Symbol, 10)
End Function

' Retrieve holding collateral type.
Public Function GetHoldingCollateralType(pseudoAccount As String, _
    Symbol As String) As String
    GetHoldingCollateralType = ReadHoldingColumn(pseudoAccount, Symbol, 11)
End Function

' Retrieve holding collateral quantity.
Public Function GetHoldingCollateralQuantity(pseudoAccount As String, _
    Symbol As String) As Long
    GetHoldingCollateralQuantity = AtToLong(ReadHoldingColumn(pseudoAccount, Symbol, 12))
End Function

' Retrieve holding haircut.
Public Function GetHoldingHaircut(pseudoAccount As String, _
    Symbol As String) As Double
    GetHoldingHaircut = AtToDouble(ReadHoldingColumn(pseudoAccount, Symbol, 13))
End Function

' Retrieve holding average price.
Public Function GetHoldingAvgPrice(pseudoAccount As String, _
    Symbol As String) As Double
    GetHoldingAvgPrice = AtToDouble(ReadHoldingColumn(pseudoAccount, Symbol, 14))
End Function

' Retrieve holding instrument token.
Public Function GetHoldingInstToken(pseudoAccount As String, _
    Symbol As String) As String
    GetHoldingInstToken = ReadHoldingColumn(pseudoAccount, Symbol, 15)
End Function

' Retrieve holding last traded price.
Public Function GetHoldingLtp(pseudoAccount As String, _
    Symbol As String) As Double
    GetHoldingLtp = AtToDouble(ReadHoldingColumn(pseudoAccount, Symbol, 21))
End Function

' Retrieve holding current value.
Public Function GetHoldingCurrentValue(pseudoAccount As String, _
    Symbol As String) As Double
    GetHoldingCurrentValue = AtToDouble(ReadHoldingColumn(pseudoAccount, Symbol, 22))
End Function

' ***************************************************************************
'
' PORTFOLIO BY NAME -- the recommended way to read a portfolio.
'
' The Get...() functions above still work and are not going away. They come
' from the add-in, so each one takes the whole identity of a row and looks that
' row up again, and each one addresses its field by a column NUMBER. Twenty
' fields means twenty lookups, and a column number is only correct until the
' server's CSV changes shape.
'
' These read by NAME instead, and find the row once:
'
'     h = AtFindHolding(AT_ACCOUNT, "NSE", "IOC")
'     If AtFound(h) Then
'         qty = AtNum(h, "QUANTITY")
'         isin = AtText(h, "ISIN")
'     End If
'
' AtFound() matters. A holding you do not have and a lookup that is broken both
' read as 0, and telling them apart is exactly what was missing when the
' holdings fault above went unnoticed.
'
' Field names are the column names in the server's CSV header, case does not
' matter: QUANTITY, PNL, ISIN, PRODUCT, LTP, AVGPRICE, STATUS, TRADETYPE,
' NETQUANTITY, BUYAVGPRICE, and so on. An unknown name returns blank rather
' than the wrong field.
'
' Every one of these is safe to call from a worksheet cell as well as from VBA.
'
' ***************************************************************************

' A row handle: which account, which dataset, which row. Held as text so it can
' sit in a cell like any other value.
Public Function AtHandle(pseudoAccount As String, dataset As String, _
    row As Long) As String
    AtHandle = pseudoAccount & "|" & dataset & "|" & CStr(row)
End Function

' True when a find...() actually found something.
Public Function AtFound(handle As String) As Boolean
    AtFound = (Len(handle) > 0)
End Function

' One field of a found row, as text. Blank for an unknown field or a handle
' that found nothing.
Public Function AtText(handle As String, fieldName As String) As String

    If Len(handle) = 0 Then
        AtText = ""
        Exit Function
    End If

    AtText = AtFieldByName(AtPipeField(handle, 1), AtPipeField(handle, 2), _
        CLng(Val(AtPipeField(handle, 3))), fieldName)

End Function

' The same, as a number.
Public Function AtNum(handle As String, fieldName As String) As Double
    AtNum = AtToDouble(AtText(handle, fieldName))
End Function

' Finds one holding. Exchange decides which independent symbol column is
' matched, so BSE holdings are addressable too.
Public Function AtFindHolding(pseudoAccount As String, exchange As String, _
    Symbol As String) As String

    Dim field As String
    Dim row As Long

    field = "INDEPENDENTSYMBOLNSE"
    If UCase$(Trim$(exchange)) = "BSE" Then field = "INDEPENDENTSYMBOLBSE"

    row = AtFindRowByName(pseudoAccount, AT_DS_HOLDINGS, field, Symbol)

    If row = 0 Then
        AtFindHolding = ""
    Else
        AtFindHolding = AtHandle(pseudoAccount, AT_DS_HOLDINGS, row)
    End If

End Function

' Finds one order by the broker's order id -- what PlaceOrder returns.
Public Function AtFindOrder(pseudoAccount As String, orderId As String) As String

    Dim row As Long
    row = AtFindRowByName(pseudoAccount, AT_DS_ORDERS, "ID", orderId)

    If row = 0 Then
        AtFindOrder = ""
    Else
        AtFindOrder = AtHandle(pseudoAccount, AT_DS_ORDERS, row)
    End If

End Function

' Finds one position. A position has no id of its own, so it is identified by
' category, type, exchange and symbol together.
Public Function AtFindPosition(pseudoAccount As String, category As String, _
    posType As String, exchange As String, Symbol As String) As String

    Dim row As Long

    row = AtFindRowByName4(pseudoAccount, AT_DS_POSITIONS, _
        "CATEGORY", category, "TYPE", posType, _
        "INDEPENDENTEXCHANGE", exchange, "INDEPENDENTSYMBOL", Symbol)

    If row = 0 Then
        AtFindPosition = ""
    Else
        AtFindPosition = AtHandle(pseudoAccount, AT_DS_POSITIONS, row)
    End If

End Function

' Finds one margin category: EQUITY, COMMODITY or ALL.
Public Function AtFindMargin(pseudoAccount As String, category As String) As String

    Dim row As Long
    row = AtFindRowByName(pseudoAccount, AT_DS_MARGINS, "CATEGORY", category)

    If row = 0 Then
        AtFindMargin = ""
    Else
        AtFindMargin = AtHandle(pseudoAccount, AT_DS_MARGINS, row)
    End If

End Function

' How many rows a portfolio holds, and the n-th of them, 1 based.
'
' There was no way to WALK a portfolio before: every getter needed a symbol you
' already knew, so a sheet could not ask "what am I holding?" or "what is still
' open?". These make that possible -- fill a column with
' =AtText(AtHoldingAt($A$1, ROW()-1), "INDEPENDENTSYMBOLNSE") and the portfolio
' lists itself.
Public Function AtHoldingCount(pseudoAccount As String) As Long
    AtHoldingCount = AtRowCount(pseudoAccount, AT_DS_HOLDINGS)
End Function

Public Function AtPositionCount(pseudoAccount As String) As Long
    AtPositionCount = AtRowCount(pseudoAccount, AT_DS_POSITIONS)
End Function

Public Function AtOrderCount(pseudoAccount As String) As Long
    AtOrderCount = AtRowCount(pseudoAccount, AT_DS_ORDERS)
End Function

Public Function AtHoldingAt(pseudoAccount As String, n As Long) As String
    If n < 1 Or n > AtRowCount(pseudoAccount, AT_DS_HOLDINGS) Then
        AtHoldingAt = ""
    Else
        AtHoldingAt = AtHandle(pseudoAccount, AT_DS_HOLDINGS, n)
    End If
End Function

Public Function AtPositionAt(pseudoAccount As String, n As Long) As String
    If n < 1 Or n > AtRowCount(pseudoAccount, AT_DS_POSITIONS) Then
        AtPositionAt = ""
    Else
        AtPositionAt = AtHandle(pseudoAccount, AT_DS_POSITIONS, n)
    End If
End Function

Public Function AtOrderAt(pseudoAccount As String, n As Long) As String
    If n < 1 Or n > AtRowCount(pseudoAccount, AT_DS_ORDERS) Then
        AtOrderAt = ""
    Else
        AtOrderAt = AtHandle(pseudoAccount, AT_DS_ORDERS, n)
    End If
End Function
