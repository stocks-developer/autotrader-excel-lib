Attribute VB_Name = "AutoTraderConfig"
' *****************************************************************************
'
' AutoTrader Web -- settings for the direct (HTTP) Excel module.
'
' THIS IS THE ONLY FILE YOU EDIT. Put your API key in it and save.
'
' Get your API key from your account settings:
'   https://webx.stocksdeveloper.in/
'
' Version: 1.0
'
' *****************************************************************************

Option Explicit

' *****************************************************************************
' YOUR API KEY.
'
' Treat it like a password. Anyone holding it can place orders in your
' accounts. Do not share a workbook that still has your key in it, and do not
' post it on a forum when asking for help.
' *****************************************************************************

Public Const AT_API_KEY As String = "<API_KEY>"

' *****************************************************************************
' Where the requests go. Leave this alone unless support asks you to change it.
' *****************************************************************************

Public Const AT_BASE_URL As String = "https://apix.stocksdeveloper.in"

' *****************************************************************************
' How long the module re-uses portfolio data before asking the server again,
' in seconds.
'
' A sheet may call twenty different Get... functions on one recalculation. Each
' is answered from the copy held in memory, so a recalculation costs one
' request per dataset instead of twenty.
'
' Lowering these does NOT get you fresher data. The server allows roughly one
' portfolio request per second per account and answers anything faster than
' that from its own copy, so a smaller number here buys extra requests and the
' same numbers. Raise them if you trade slowly and want less traffic.
'
' Holdings barely move during the day, which is why they are refreshed rarely.
' *****************************************************************************

Public Const AT_TTL_ORDERS As Long = 2

Public Const AT_TTL_POSITIONS As Long = 2

Public Const AT_TTL_MARGINS As Long = 30

Public Const AT_TTL_HOLDINGS As Long = 300

' *****************************************************************************
' How long to wait for the server, in milliseconds.
'
' A read is quick. A command is not: the server sends the order to your broker
' and waits for the broker's answer before replying, so it needs the longer
' allowance.
'
' AT_HTTP_TIMEOUT_COMMAND must stay ABOVE the server's own 25 second deadline.
' If Excel gives up first you lose the one thing worth having -- the server's
' own account of what happened to the order -- and are left unable to tell a
' rejected order from a live one.
' *****************************************************************************

Public Const AT_HTTP_TIMEOUT_READ As Long = 10000

Public Const AT_HTTP_TIMEOUT_COMMAND As Long = 30000

' *****************************************************************************
' How far to read along the header when resolving a column name.
'
' The widest dataset the server sends is orders, at 33 columns, so this is
' generous. It is only a stop so that a malformed header cannot spin a sheet:
' the scan ends at the first empty column anyway.
' *****************************************************************************

Public Const AT_HTTP_MAX_COLUMNS As Integer = 60

' *****************************************************************************
' Print every request and reply to the VBA Immediate window (Ctrl+G).
'
' Useful when setting up, noisy afterwards. It does NOT print your API key.
' *****************************************************************************

Public Const AT_HTTP_DEBUG As Boolean = False
