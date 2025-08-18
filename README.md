# Finance & Crypto Trading Tracker (Google Sheets + Apps Script)

Centralized tracker for spot & perps trading, wallet balances, and rollups-built on **Google Sheets** with **Apps Script** automations. Pulls live market/wallet data, standardizes logs, and maintains summaries for quick P\&L review.

> **Live sheet:** *https://docs.google.com/spreadsheets/d/1P9Iue5bApHW_WJjJ5ZmfBOmB2sJYG7m6u2_UOeTw8hw/edit?usp=sharing* (view-only)
> **Code:** this repository (Apps Script source)<br>
> **Contact:** *ktabesbusiness@gmail.com | https://www.linkedin.com/in/kyle-table/*

---

## Contents

* [Features](#features)
* [Sheet Structure](#sheet-structure)
* [Tech & APIs](#tech--apis)
* [Quick Start](#quick-start)
  * [Configure Secrets](#configure-secrets-required)
  * [Wallets & Constants](#wallets--constants)
  * [How Keys Are Read](#how-keys-are-read)
* [Typical Workflows](#typical-workflows)
* [Automation & Triggers](#automation--triggers)
* [Roadmap](#roadmap)
* [Acknowledgments](#acknowledgments)

---

## Features

* **Wallet balance sync (Solana)** via **Helius RPC**
* **Price data** via **CoinMarketCap (CMC)** for USD conversion and summaries
* **Perps fills ingestion** (Hyperliquid API) into a normalized **Perps Data Table**
* **Spot trade logging** with calculated fields (fees, PnL %, PnL USD)
* **Automations**
  * `onEdit` helpers: entry-row shifting, dependent cell updates, validation
  * Optional time-driven refresh (daily/hourly summary rebuilds)
* **Dashboards**: **Assets Summary**, **Trading Summary**, and **Assets** with time-driven triggers to keep data up-to-date

---

## Sheet Structure

- **Assets Summary:** Portfolio view; current balances, USD valuation, daily and weekly deltas.
- **Trading Summary:** P&L rollups (win rate, average R, weekly totals).
- **Assets:** Detailed holdings by asset and wallet.
- **Assets Manual Input History:** Audit log for manual adjustments.
- **Assets Total History:** Daily or scheduled snapshots of total portfolio value, used for trend lines and WoW or MoM comparisons.
- **Spot Trading Log:** Manual spot trades (date, pair, side, quantity, price, fees, notes).
- **Perps Data Table:** Normalized perps executions (date, time, asset, side, entry and exit, quantity, PnL percent and USD, fees).
- **Memecoin Trading Analysis:** Aggregations and visuals built from the Memecoin Trading Data Table (PnL by token or day, win rate, time of day performance, etc.)
- **Memecoin Trading Data Table:** Row level trade and execution table for memecoin activity (timestamp, token or pair, size, price, market cap or volume at entry and exit, notes).
- **Change Log:** Human readable record of feature changes to the sheet and Apps Script (version and date, what changed, impact and notes).

> You may also have hidden helper tabs (for example, `_HL_Ledger`, `_AutocompleteHelper`) that support validation, lookups, or collateral tracking.

---

## Tech & APIs

* **Google Sheets** + **Apps Script** (JavaScript)
* **APIs**
  * **Helius** (Solana): wallet balances & token metadata
  * **CoinMarketCap (CMC)**: price quotes & conversions
  * **Hyperliquid**: perps/spot fills (optional/if enabled)

---

## Quick Start

1. Open the live sheet (*link above*)
2. Make your own copy: **File → Make a copy**
3. In your copy: **Extensions → Apps Script** to open the script project

### Configure Secrets (required)

Store API keys in **Script properties** so they are **not in code**:

1. Apps Script editor → **Project Settings (gear)**
2. **Script properties** → **Add property** for each key
   * `CMC_API_KEY` = `your-cmc-key`
   * `HELIUS_API_KEY` = `your-helius-key`
3. **Save**

> Do **not** put keys in code. (Recommended if you plan to share your sheet publically.)

### Wallets & Constants

Set/adjust IDs you use (tab names, wallets, etc.). Example:

```js
// Example constants used by the project
const ASSETS_SUMMARY_SHEET  = 'Assets Summary';
const TRADING_SUMMARY_SHEET = 'Trading Summary';
const ASSETS_SHEET          = 'Assets';
const HISTORY_SHEET         = 'Assets Manual Input History';
const SPOT_SHEET            = 'Spot Trading Log';
const PERPS_SHEET           = 'Perps Data Table';

const HELIUS_RPC_URL = 'https://mainnet.helius-rpc.com/?api-key=';
const LAMPORTS_PER_SOL = 1e9;

// Wallets (replace with yours)
const WALLET_1 = '...';
const WALLET_2 = '...';
const WALLET_3 = '...';
```

### How Keys Are Read

Load keys from **Script properties** at runtime (no helpers required):

```js
const SP = PropertiesService.getScriptProperties();
const CMC_API_KEY    = SP.getProperty('CMC_API_KEY');
const HELIUS_API_KEY = SP.getProperty('HELIUS_API_KEY');

// Optional: fail fast if missing
(function validateSecrets(){
  if (!CMC_API_KEY) throw new Error('Missing CMC_API_KEY in Script properties');
  if (!HELIUS_API_KEY) throw new Error('Missing HELIUS_API_KEY in Script properties');
})();
```

---

## Typical Workflows

### Refresh balances & prices

* Run a refresh function (e.g., `UpdateTradingSummary_Hyperliquid()` or your balance updater) from **Run** in the Apps Script editor
* Or attach a **time-driven trigger** (see below)

### Log a new spot trade

* Enter the row in **Spot Trading Log** (date, pair, side, qty, price, fee)
* Calculated fields update automatically

### Perps data

* Run your rebuild (e.g., `Rebuild_Perps_Data_Table()`) to pull fills via API, normalize rows, and update summary metrics

### Manual asset adjustments

* Edit holdings in **Assets**; script preserves **B3\:G3** as the entry row, shifts completed entries down, and records edits in **Assets Manual Input History**

---

## Automation & Triggers

**Global onEdit behavior:** runs context specific updates based on the edited sheet, row, and column.

- **Assets Summary:**
  - When any of the three watched columns inside `NR.currentPositions` change, run `overridePerpsTokens()`.
  - When the price column of `NR.watchlist` changes, run `syncAssetTotalValue()` and `updateTopAssetsSummary()`.

- **Spot Trading Log:**
  - Rows 16 and 43, even columns from B to Z: run `syncAssetTags()`.
  - Rows 6 and 33, any column: run `syncAssetTokens()`, `syncAssetCostBasis()`, `mergeSpotAndManualHistory()`, `syncAssetTotalValue()`, `sortAssetsSheet()`, `updateTopAssetsSummary()`, `overridePerpsTokens()`.
  - Rows 4 and 31, any column: write frequency metrics to `Z1 = Spot_Trade_Frequency()` and `Z2 = Spot_Avg_Trades_Per_Month()`.

- **Perps Data Table:**
  - Watch columns include Date, Time, Asset, Direction, Entry, Token, Leverage, Collateral, plus columns R, U, AH, AX. Any edit in these columns sets `shouldTrigger = true`.
  - Columns N, AE, AU, BC trigger only when the new value is a numeric USD value. If numeric, set `shouldTrigger = true`.
  - When `shouldTrigger` is true, run `rebuildTradeBay()`, `sortAssetsSheet()`, `updateTopAssetsSummary()`.
  - If Date or Time were edited, set `Z1 = PerpsFrequency()` and `Z2 = PERPS_AVG_TRADES_PER_WEEK()`.

- **Assets Manual Input History:**
  - Range `C5:H1000`: when the edited row within this block is fully filled, run `overridePerpsTokens()`.

- **Assets:**
  - Range `C4:L32`: any change recalculates totals and summaries with `syncAssetTotalValue()`, `updateTopAssetsSummary()`, `overridePerpsTokens()`.
  - Range `C4:L32`: always re sort by column F except when editing column F, via `sortAssetsSheet()`.
  - Range `C4:L32`: final refresh of top assets with `updateTopAssetsSummary()` and `overridePerpsTokens()`.

---

## Repo Layout

```
/src
Code.gs
Hyperliquid.gs
appsscript.json
README.md
```

* `appsscript.json` - Apps Script project manifest

---

## Roadmap 

* [ ] Fully Developed Hyperliquid API Integration and Automation (Perps and Spot)
* [ ] A New Sheet with the Price of All Top 100 Cryptocurrencies and Top TradFi Assets (S&P 500, APPL, etc.)
* [ ] Expanded dashboards (weekly win rate, time-of-day analysis)

---

## Acknowledgments

APIs: **Helius**, **CoinMarketCap**, **Hyperliquid**. Built with **Google Sheets** + **Apps Script**.
