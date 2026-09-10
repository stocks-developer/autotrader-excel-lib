# AutoTrader Web Excel Tools: Bulk Multi-Account Order Placement for 40+ Indian Brokers

> Place and copy **bulk orders from a spreadsheet** into one account or many at once, across **40+ Indian brokers**, straight from **Excel**. Includes the ready-made Order Pad and VBA modules that expose the full trading API. Part of **[AutoTrader Web](https://stocksdeveloper.in/)** by **Stocks Developer**.

[![Brokers supported](https://img.shields.io/badge/brokers-40%2B-2ea44f)](https://stocksdeveloper.in/#supported-brokers)
[![Free trial](https://img.shields.io/badge/free%20trial-1%20month-blue)](https://webx.stocksdeveloper.in/register)
[![Uptime](https://img.shields.io/badge/uptime-99.98%25-brightgreen)](https://stocksdeveloper.in/features/)
[![Setup guide](https://img.shields.io/badge/docs-Excel%20setup-8a2be2)](https://stocksdeveloper.in/documentation/client-setup/excel-library/)

---

## What is this?

The **AutoTrader Web Excel tools** let you place orders into one or many broker accounts straight from a spreadsheet. There is a ready-made workbook, the **Order Pad**, for placing orders across accounts and seeing what happened to each one, and a set of **VBA modules** that expose the full AutoTrader Web trading API for anyone who codes strategies in Excel macros.

- **Bulk orders from a spreadsheet.** Enter orders in a sheet and send them all at once.
- **Multi-account and multi-broker.** Copy one order across many accounts (with per-account quantity), or place a different order per account, across different brokers.
- **Scheduled placement.** Fire all your orders at a set time, for example at market open.
- **Full API in VBA.** Call every API function from your own Excel macros.
- **Direct connection.** Your spreadsheet talks to AutoTrader Web over the internet. There is nothing else to install and nothing that has to keep running.

## What is AutoTrader Web?

**[AutoTrader Web](https://stocksdeveloper.in/)** by **Stocks Developer** is copy trading and multi-account software for Indian brokers. Monitor every broker account on one screen and act across all of them at once.

- **All your accounts, one screen.** Live, consolidated P&L, holdings, positions, orders and margins across every account and broker.
- **Copy trading, two ways.** PMS copy from our terminal, and master-child copy in the background, across brokers, with per-account sizing. [Copy trading software](https://stocksdeveloper.in/copy-trading-software/)
- **Bulk orders.** Place, modify, cancel and square-off across many accounts in one action.
- **GTT, bracket and cover orders**, order slicing and market price protection.
- **TradingView automation.** Turn your own chart alerts into real orders.
- **APIs and SDKs.** Excel, AmiBroker and MetaTrader, plus Java, Python, C# and HTTP REST / CSV.
- **8+ years in operation. 99.98% uptime. 40+ brokers. Under 100 ms data latency.**

## Why traders and developers choose us

- 🆓 **Free static IP included** with every account. Saves up to **₹500 per broker account per month** that other tools charge extra for.
- 💸 **One of the lowest prices in the category.** **₹295 to ₹495 per account per month**, all taxes and the static IP included. No setup fee, no hidden charges.
- ☁️ **Nothing to install for the platform.** Monitor and trade from your browser on PC or mobile, from anywhere.
- 🔗 **40+ Indian brokers on one platform.** One of the widest broker coverages available.
- 🔁 **Two ways to copy trade**, PMS and master-child, both included.
- 🔒 **Security you control.** API credentials encrypted and stored in India, broker OAuth login and two-factor authentication, portfolio data never stored, plus a Kill Switch and a full activity log.
- 🎁 **Free 1-month trial** on supported brokers.

## Supported brokers

AutoTrader Web works with **40+ Indian brokers**:

5paisa · AC Agarwal · Aetram Trades · Alice Blue · Ambalal Shares · Anand Rathi · Angel One · Arham Share · ATS · AxisDirect · Choice · DBOnline · Dhan · Eureka Share · Finvasia · Flattrade · FYERS · Groww · IIFL Securities · Jainam (Prop & Retail) · Kotak Securities · Mastertrust · Mirae Asset Sharekhan · MLB Stock Broking · Motilal Oswal · Nuvama · PL Capital (PLIndia) · Profitmart · Pune E-Stock Broking (PESB) · Raghunandan Money · Religare · Share India (Prop & Retail) · SMC India · Stocko · SW Capital · Tradejini · Tradeswift · Upstox · Wisdom Capital · Zebu · Zerodha

*Plus any broker that supports the Symphony XTS API.* See the [full, always-current broker list](https://stocksdeveloper.in/#supported-brokers) and the [broker setup guides](https://stocksdeveloper.in/documentation/supported-brokers/).

## Quick start

Excel talks to AutoTrader Web directly.

### Use the Order Pad

A ready-made workbook that places orders into one account or many, and writes back what happened to each one: the broker's order id, its status and its reason.

There is nothing to import. Every module it needs is already inside the workbook.

1. Download [`AutoTraderWeb-OrderPad.xlsm`](clients/excel/current/samples/AutoTraderWeb-OrderPad.xlsm) from the [`clients/excel/current/samples`](clients/excel/current/samples) folder.
2. Open it and clear the Excel warnings (Enable Editing, Enable Content, and Unblock the file).
3. Put your API key in the **Value** column on the **Configuration** sheet.
4. List your pseudo accounts on the **Accounts** sheet and put `Y` beside each one you want to trade.
5. Press **Check accounts**. It reads each account's free margin and places nothing, so a wrong account name is caught before you send fifty orders rather than after.
6. Type your orders on the **Orders** sheet and press **Place orders**. Every order lands on the **Results** sheet with its id, status and reason.

The buttons sit along the top of the Orders, Accounts and Results sheets and stay in place as you scroll. Hover any column heading for a note on what goes in it.

Leave the **Account** column blank on an order row and that order goes to every account you ticked. How much each one gets is set by **Allocation mode** on the Configuration sheet: `SAME`, `PER ACCOUNT` or `SPLIT`.

**Place at** on the same sheet arms a timer, and the Cancel timer button counts down while it is armed.

Your API key is in AutoTrader Web under **Settings -> Security**. It is hidden, so click **Reveal**, type your AutoTrader password, then **Copy**. Treat it like a password: anyone who has it can place orders in your accounts, so do not share a workbook that still has your key in it.

### Write your own workbook

1. Sign in at [webx.stocksdeveloper.in](https://webx.stocksdeveloper.in/) and go to **Tools -> Library**.
2. Download the Excel modules. Your API key is already inside the download, which is why it asks for your password.
3. Open your workbook, press **Alt+F11**, then **File -> Import File** and import all three `.bas` files.
4. Save the workbook as a macro-enabled workbook (`.xlsm`).

> The modules also live in [`direct/`](direct) in this repository, but the copy you download from your account is the one that already carries your API key.

The Order Pad covers all of this from one grid: many orders across many accounts, one order copied across accounts with a different quantity or weight for each, a different order per account across brokers, and placement at a set time with the **Arm timer** button.

Full step-by-step guide: **[Excel tools setup](https://stocksdeveloper.in/documentation/client-setup/excel-library/)** and [bulk orders from Excel](https://stocksdeveloper.in/documentation/excel/). Get your API key from your [account settings](https://webx.stocksdeveloper.in/register).

### Reading your portfolio

Find the row once, then read its fields by name. This works in a cell and in VBA:

```vb
holdRow = AtFindHolding(AT_ACCOUNT, "NSE", "IOC")

If AtFound(holdRow) Then
    qty = AtNum(holdRow, "QUANTITY")
    isin = AtText(holdRow, "ISIN")
End If
```

There is a finder for each kind of row:

| Finder | Identified by |
|---|---|
| `AtFindHolding(account, exchange, symbol)` | exchange and symbol |
| `AtFindPosition(account, category, type, exchange, symbol)` | all four together |
| `AtFindOrder(account, orderId)` | the broker's order id |
| `AtFindMargin(account, category)` | `EQUITY`, `COMMODITY` or `ALL` |

`AtFound()` tells you whether the row exists. This matters: a holding you do not have and a lookup that went wrong both read as `0`, and only `AtFound()` separates them.

Field names are the column names your data already uses, and case does not matter — `QUANTITY`, `PNL`, `LTP`, `AVGPRICE`, `ISIN`, `STATUS`, `TRADETYPE`, `NETQUANTITY`, `BUYAVGPRICE` and so on. Ask for a name that does not exist and you get a blank, never a different field by mistake.

You can also list a whole portfolio down a column, which the older functions cannot do. Put the account in `$A$1`, then fill down:

```
=AtText(AtHoldingAt($A$1, ROW()-1), "INDEPENDENTSYMBOLNSE")
=AtNum(AtHoldingAt($A$1, ROW()-1), "QUANTITY")
```

`AtHoldingCount()`, `AtPositionCount()` and `AtOrderCount()` tell you how many rows there are, with `AtPositionAt()` and `AtOrderAt()` alongside `AtHoldingAt()`.

The older `GetHoldingQuantity()`, `GetPositionNetQuantity()`, `GetOrderStatus()` style functions still work exactly as before and are not going away. Use these when you want to read several fields of the same row, or when you need to go through a portfolio without knowing the symbols in advance.

### Reading an order back after you place or change it

An order does not update the instant you place, modify or cancel it. Your broker's order book takes a few seconds to catch up, and the modules re-use portfolio data for a couple of seconds so that a recalculating sheet does not send the same request twenty times.

So a read taken immediately after a change can show the previous state. That is not a failure — it means *not updated yet*. Give it a few seconds before deciding an order did not work, and never send it a second time on the strength of a blank read.

## Repository layout

| Path | What it is |
|---|---|
| [`direct/`](direct) | The three modules. **Everything in this folder is sent to users verbatim.** |
| [`clients/excel/current/samples/`](clients/excel/current/samples) | The ready-made Order Pad workbook. |
| `clients/excel/current/samples/modules/` | Source of the Order Pad's own module, kept here so it can be read and diffed. |
| `clients/excel/current/addin/modules/` | Source of the retired add-in, kept for reference only. Nothing here is used or shipped. |
| [`tools/orderpad-check/`](tools/orderpad-check) | Checks the Order Pad workbook still matches the modules in this repository. |

> **Do not add files to [`direct/`](direct).** The download on **Tools -> Library** is built by taking that whole folder from this repository at the moment a user asks for it, so anything placed there arrives in their zip. The setup instructions tell them to import three `.bas` files, and a fourth one would collide with a module of the same name and stop the workbook compiling. Source that is not meant for users belongs beside the thing it builds, which is why the sample drivers sit under `samples/modules/`.

## Pricing and free trial

- **Free 1-month trial** on supported brokers, with every feature included.
- Then **₹295 to ₹495 per account per month**. All taxes and a free static IP are included. No setup fee, no hidden charges.
- [See full pricing](https://stocksdeveloper.in/pricing/) · [Start free](https://webx.stocksdeveloper.in/register)

## Documentation and links

| Resource | Link |
|---|---|
| 🌐 Website | https://stocksdeveloper.in/ |
| ✨ Features | https://stocksdeveloper.in/features/ |
| 💰 Pricing | https://stocksdeveloper.in/pricing/ |
| 🔁 Copy trading software | https://stocksdeveloper.in/copy-trading-software/ |
| 🏦 Supported brokers | https://stocksdeveloper.in/#supported-brokers |
| 🔒 Security and data handling | https://stocksdeveloper.in/security/ |
| 📘 Documentation | https://stocksdeveloper.in/documentation/getting-started/ |
| 🧩 API reference | https://stocksdeveloper.in/documentation/api/ |
| ⚙️ Excel tools setup | https://stocksdeveloper.in/documentation/client-setup/excel-library/ |
| 📗 Bulk orders from Excel | https://stocksdeveloper.in/documentation/excel/ |
| 🆓 Start free (1-month trial) | https://webx.stocksdeveloper.in/register |
| ✉️ Contact us | https://stocksdeveloper.in/contact/ |

## About Stocks Developer

Stocks Developer is a technology company building software tools for Indian markets, shaped by 8+ years of trader feedback. Our software runs on Google Cloud in its Mumbai, India region for fast, low-latency performance, with strong security and high reliability.

Stocks Developer provides software tools only. It gives no investment advice, tips, recommendations, or trading strategies, and it makes no trading decisions for you. All trading and investment decisions remain solely your responsibility. You set up and control every activity, and you can stop it at any time.
