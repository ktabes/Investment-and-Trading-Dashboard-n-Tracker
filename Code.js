/***********************************************
 *             --- CONFIG & CONSTANTS ---      *
 ***********************************************/
const ASSETS_SUMMARY_SHEET    = 'Assets Summary';
const TRADING_SUMMARY_SHEET   = 'Trading Summary';
const ASSETS_SHEET            = 'Assets';
const HISTORY_SHEET           = 'Assets Manual Input History';
const SPOT_SHEET              = 'Spot Trading Log';
const PERPS_SHEET             = 'Perps Data Table';
const HELIUS_RPC_URL          = 'https://mainnet.helius-rpc.com/?api-key='
const LAMPORTS_PER_SOL        = 1e9;

const __SP = PropertiesService.getScriptProperties();
const CMC_API_KEY    = __SP.getProperty('CMC_API_KEY');
const HELIUS_API_KEY = __SP.getProperty('HELIUS_API_KEY');

const WALLET_1  = __SP.getProperty('WALLET_1'); // Axiom Wallet
const WALLET_2  = __SP.getProperty('WALLET_2'); // Main Wallet
const WALLET_3  = __SP.getProperty('WALLET_3'); // Other Trading Wallet

const POS_START_ROW = 26, POS_END_ROW = 31, POS_FIRST_COL = 1, POS_NUM_COLS = 9;

// Column indexes on PERPS_SHEET:
const COL_DATE  = 1;   // Date
const COL_ASSET = 5;   // Asset
const COL_DIR   = 6;   // Direction
const COL_ENTRY = 7;   // Entry Price
const COL_TOK   = 8;   // Position Size
const COL_LEV   = 10;  // Leverage
const COL_COLL  = 11;  // Collateral
const COL_R     = 18;  // “R” fill-marker column
const COL_U     = 21;  // “U” percentage column
const COL_AH    = 37;  // “AK” percentage column
const COL_AX    = 53;  // “BA” percentage column

// ─── named ranges for the Assets sheet ───────────────────────────────────
const NR = {
  mainTable    : 'Assets_Main',      // refers to Assets!C4:L32
  historyTable : 'Assets_History',  // refers to Assets!N9:S32
  inputRow     : 'Assets_Input',     // refers to Assets!N5:R5
  
  // new named ranges on the Assets sheet
  topAssets       : 'Top_18_Assets',      // Assets!C3:N22
  currentPositions: 'Current_Positions',  // Assets!C28:K33
  watchlist       : 'Asset_Watchlist'     // Assets!P21:R33
};

/***********************************************
 *                  UTILITIES                  *
 ***********************************************/
function normalize(v){ return (v||'').toString().trim(); }
function pct(v){
  if (!v) return 0;
  if (typeof v==='string' && v.includes('%')) return parseFloat(v)/100;
  if (typeof v==='number' && v>1) return v/100;
  return v;
}
function notBlank(sh,r,c){ return sh.getRange(r,c).getValue() !== ''; }

/***********************************************
 *                   onOpen                    *
 ***********************************************/
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('🛠️Manual Functions🛠️')
      .addItem('Main Assets Table Refresh','Assets_Sheet_Refresh')
      .addItem('Manual Input','ManualInput')
      .addItem('Rebuild Positions','rebuildTradeBay')
      .addSeparator()
      .addItem('Refresh CMC Prices','refreshAllCMCPrices')
      .addItem('Sort "Assets" Sheet','sortAssetsSheet')
      .addItem('Update Top Assets Summary','updateTopAssetsSummary')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  const ss = e.source;
  const sh = e.range.getSheet();
  const r  = e.range.getRow();
  const c  = e.range.getColumn();

  // ── ASSETS SUMMARY ─────────────────────────────────────────────────────────
  if (sh.getName() === ASSETS_SUMMARY_SHEET) {
    const posRng = ss.getRangeByName(NR.currentPositions);
    const pr     = posRng.getRow();
    const pc     = posRng.getColumn();
    const nr     = posRng.getNumRows();
    const watchRng = ss.getRangeByName(NR.watchlist);
    const wr     = watchRng.getRow();
    const wc     = watchRng.getColumn();
    const wn     = watchRng.getNumRows();

    // override Perps when any cell in I28:K33 changes
    if (r >= pr && r < pr + nr && [pc+6, pc+7, pc+8].includes(c)) {
      overridePerpsTokens();
    }
    // recalc USD & summary when price in R21:R33 changes
    if (c === wc + 2 && r >= wr && r < wr + wn) {
      syncAssetTotalValue();
      updateTopAssetsSummary();
    }
    return;
  }

  // ── SPOT TRADING LOG ────────────────────────────────────────────────────────
  if (sh.getName() === SPOT_SHEET) {
    // 1) Tag-flips on rows 16 & 43
    if ([16, 43].includes(r) && c >= 2 && c <= 26 && c % 2 === 0) {
      syncAssetTags();
      return;
    }
    // 2) Snapshot rows 6 & 33
    if ([6, 33].includes(r) && c >= 1 && c <= sh.getLastColumn()) {
      syncAssetTokens();
      syncAssetCostBasis();
      mergeSpotAndManualHistory();
      syncAssetTotalValue();
      sortAssetsSheet();
      updateTopAssetsSummary();
      overridePerpsTokens();
      return;
    }
    // 3) Frequency metrics on row 4 or 31
    if ([4, 31].includes(r) && c >= 1 && c <= sh.getLastColumn()) {
      sh.getRange('Z1').setValue(Spot_Trade_Frequency());
      sh.getRange('Z2').setValue(Spot_Avg_Trades_Per_Month());
    }
    return;
  }

  // ── PERPS DATA TABLE ────────────────────────────────────────────────────────
// ---- PERPS block (inline, no helper) --------------------------------------
const PERPS_SHEET_NAME =
  (typeof globalThis.PERPS_SHEET === 'string' && globalThis.PERPS_SHEET)
    ? globalThis.PERPS_SHEET : 'Perps';

if (sh.getName() === PERPS_SHEET_NAME) {
  // Use global constants if you have them; otherwise fall back safely.
  // IMPORTANT: refer to globalThis to avoid TDZ with same-named locals.
  const COL_DATE = (typeof globalThis.COL_DATE !== 'undefined') ? globalThis.COL_DATE : 1;   // A
  const COL_TIME = (typeof globalThis.COL_TIME !== 'undefined') ? globalThis.COL_TIME : 2;   // B
  const COL_ASSET= (typeof globalThis.COL_ASSET!== 'undefined') ? globalThis.COL_ASSET: 3;
  const COL_DIR  = (typeof globalThis.COL_DIR  !== 'undefined') ? globalThis.COL_DIR  : 4;
  const COL_ENTRY= (typeof globalThis.COL_ENTRY!== 'undefined') ? globalThis.COL_ENTRY: 5;
  const COL_TOK  = (typeof globalThis.COL_TOK  !== 'undefined') ? globalThis.COL_TOK  : 6;
  const COL_LEV  = (typeof globalThis.COL_LEV  !== 'undefined') ? globalThis.COL_LEV  : 7;
  const COL_COLL = (typeof globalThis.COL_COLL !== 'undefined') ? globalThis.COL_COLL : 8;
  const COL_R    = (typeof globalThis.COL_R    !== 'undefined') ? globalThis.COL_R    : 18;  // adjust to your real index
  const COL_U    = (typeof globalThis.COL_U    !== 'undefined') ? globalThis.COL_U    : 21;  // adjust
  const COL_AH   = (typeof globalThis.COL_AH   !== 'undefined') ? globalThis.COL_AH   : 34;  // AH
  const COL_AX   = (typeof globalThis.COL_AX   !== 'undefined') ? globalThis.COL_AX   : 50;  // AX

  // Main watched columns from your snippet (always trigger on any edit)
  const watchCols = [
    COL_DATE, COL_TIME, COL_ASSET, COL_DIR,
    COL_ENTRY, COL_TOK, COL_LEV, COL_COLL,
    COL_R, COL_U, COL_AH, COL_AX
  ].filter(n => typeof n === 'number' && n > 0);

  // Extra columns that should only trigger when a **numeric USD** value is entered:
  // N, AE, AU, BC  -> 14, 31, 47, 55
  const USD_WATCH_COLS = [14, 31, 47, 55];

  // Edited range geometry
  const r1 = r;
  const r2 = r + e.range.getNumRows() - 1;
  const c1Local = c1;
  const c2Local = c2;

  // Helpers (kept inline per your preference)
  const intersectsAny = (cols) => cols.some(col => col >= c1Local && col <= c2Local);
  const isNumericUSD = (val) => {
    if (val === null || val === '') return false;
    if (typeof val === 'number' && !isNaN(val)) return true;
    if (typeof val === 'string') {
      const cleaned = val.replace(/\$/g, '').replace(/,/g, '').trim();
      if (cleaned === '') return false;
      return !isNaN(Number(cleaned));
    }
    return false;
  };

  // Should we trigger the pipeline?
  let shouldTrigger = false;

  // A) Any edit in your original watch columns triggers
  if (intersectsAny(watchCols)) {
    shouldTrigger = true;
  }

  // B) Edits in N/AE/AU/BC only trigger if at least one edited cell in those columns is numeric
  if (!shouldTrigger && intersectsAny(USD_WATCH_COLS)) {
    const vals = e.range.getValues(); // 2D
    outer:
    for (let i = 0; i < vals.length; i++) {
      for (let j = 0; j < vals[i].length; j++) {
        const absCol = c1Local + j;
        if (USD_WATCH_COLS.indexOf(absCol) !== -1 && isNumericUSD(vals[i][j])) {
          shouldTrigger = true;
          break outer;
        }
      }
    }
  }

  if (shouldTrigger) {
    rebuildTradeBay();
    sortAssetsSheet();
    updateTopAssetsSummary();
  }

  // Independently: if Date or Time columns were touched, update Z1 / Z2
  if (intersectsAny([COL_DATE, COL_TIME])) {
    sh.getRange('Z1').setValue(PerpsFrequency());
    sh.getRange('Z2').setValue(PERPS_AVG_TRADES_PER_WEEK());
  }

  return; // PERPS handled; stop here
}

  // ── ASSETS sheet ────────────────────────────────────────────────────────────

    // ── Full‐history sheet (C5:H1000) ────────────────────────────────
if (sh.getName() === HISTORY_SHEET) {
  const fullRng = ss.getRange('C5:H1000');
  const fr = fullRng.getRow(), lr = fullRng.getLastRow();
  const fc = fullRng.getColumn(), lc = fullRng.getLastColumn();

  // only when editing anywhere inside that block
  if (r >= fr && r <= lr && c >= fc && c <= lc) {
    // grab all six columns (asset,chain,tok,cost,tag,date)
    const vals = sh.getRange(r, fc, 1, lc - fc + 1)
                   .getValues()[0]
                   .map(String);
    // only fire once the entire row is filled
    if (vals.every(v => v)) {
      overridePerpsTokens();
    }
  }
  return;
}

    // D) Any change in main table C4:L32 → recalc & summary
    if (r >= 4 && r <= 32 && c >= 3 && c <= 12) {
      syncAssetTotalValue();
      updateTopAssetsSummary();
      overridePerpsTokens();
    }

    // E) Always re-sort C4:L32 by F except when editing F itself
    if (!(r >= 4 && r <= 32 && c === 6)) {
      sortAssetsSheet();
    }

    // F) Final top-assets summary refresh
    if (r >= 4 && r <= 32 && c >= 3 && c <= 12) {
      updateTopAssetsSummary();
      overridePerpsTokens(); 
    }
  }

/**
 * Full Refresh of "Assets" Main Table (C4:L32)
 */
function Assets_Sheet_Refresh() {
  const ss       = SpreadsheetApp.getActive();
  const mainRng  = ss.getRangeByName('Assets_Main');

  // 0) clear out the main table
  mainRng.clearContent();

  // 1) pull in spot data
  syncAssetNames();
  syncAssetBlockchains();
  syncAssetCostBasis();
  overridePerpsTokens();
  syncAssetTokens();
  syncAssetFormulas();
  syncAssetTags();

  //2) Merge Spot and Manual History into "Assets" Main Table (C4:L32)
  mergeSpotAndManualHistory();

  //3) Override "Hyperliquid (Perps)" Tokens in Main Table and Recent History.
  overrideHistoryTokens();
  overridePerpsTokens();

  //4) Syncs All Assets USD Values  
  syncAssetTotalValue();

  // 5) resort, summarize & final refreshes
  updateTopAssetsSummary();
  sortAssetsSheet();
  rebuildTradeBay();
  updateSolValueInTradingSummary();
  refreshAllCMCPrices();
}

/**
  * Starts the Manual Input Workflow (N5:R5) on the "Assets" Sheet
  */
function ManualInput() {
  const ss       = SpreadsheetApp.getActive();
  const inputRng = ss.getRangeByName('Assets_Input');
  const mainRng  = ss.getRangeByName('Assets_Main');

  // 1) archive the manual row
  recordManualHistory();

  // 2) clear out the main table
  mainRng.clearContent();

  // 3) pull in spot data
  syncAssetNames();
  syncAssetBlockchains();
  syncAssetCostBasis();
  overridePerpsTokens();
  syncAssetTokens();
  syncAssetFormulas();
  syncAssetTags();

  //4) Merge Spot and Manual History into "Assets" Main Table (C4:L32)
  mergeSpotAndManualHistory();

  //5) Override "Hyperliquid (Perps)" Tokens in Main Table and Recent History.
  overrideHistoryTokens();
  overridePerpsTokens();

  //6) Syncs All Assets USD Values  
  syncAssetTotalValue();

  // 7) resort, summarize & final refreshes
  updateTopAssetsSummary();
  rebuildTradeBay();
  sortAssetsSheet();
  updateSolValueInTradingSummary();
  refreshAllCMCPrices();

  //8) Refresh Manual Input
  inputRng.clearContent();

} 

/** 
 * Syncs Assets Names on the "Assets"'s Main Table (C4:C32) with "Spot Trading Log" 
 */
function syncAssetNames() {
  const ss       = SpreadsheetApp.getActive();
  const spot     = ss.getSheetByName(SPOT_SHEET);
  const assets   = ss.getSheetByName(ASSETS_SHEET);
  const mainRng  = ss.getRangeByName(NR.mainTable);  // Assets!C4:L32
  const startRow = mainRng.getRow();                 // 4
  const startCol = mainRng.getColumn();              // 3 (column C)
  const numRows  = mainRng.getNumRows();             // 29 rows

  // 1) Pull spot names from row 1 & 28 of the Spot sheet
  const r1 = spot.getRange(1, 1, 1, 26).getValues()[0];
  const r2 = spot.getRange(28,1, 1, 26).getValues()[0];
  const spotNames = r1.concat(r2)
                      .map(normalize)
                      .filter(v => v && v.toLowerCase() !== 'asset');

  // 2) Pull existing manual names from C4:C32
  const existing = assets
    .getRange(startRow, startCol, numRows, 1)
    .getValues().flat()
    .map(normalize)
    .filter(v => v);

  // 3) Union them
  const all = Array.from(new Set(spotNames.concat(existing)));

  // 4) Clear and rewrite into C4:C32
  assets.getRange(startRow, startCol, numRows, 1).clearContent();
  if (all.length) {
    assets
      .getRange(startRow, startCol, all.length, 1)
      .setValues(all.map(v => [v]));
  }
}

/**
 * Syncs Asset Blockchain Tags on the "Assets"'s Main Table (D3:D32) with "Spot Trading Log" (Only fills in B when it’s blank, otherwise leaves your manual Blockchain tags untouched.)
 */
function syncAssetBlockchains() {
  const ss      = SpreadsheetApp.getActive();
  const spot    = ss.getSheetByName(SPOT_SHEET);
  const assets  = ss.getSheetByName(ASSETS_SHEET);
  const mainRng = ss.getRangeByName(NR.mainTable); // C4:L32
  const numRows = mainRng.getNumRows();
  if (numRows < 1) return;

  // build lookup from Spot headers → blockchain
  const n1 = spot.getRange(1, 1, 1, 26).getValues()[0];
  const b1 = spot.getRange(3, 1, 1, 26).getValues()[0];
  const n2 = spot.getRange(28,1, 1, 26).getValues()[0];
  const b2 = spot.getRange(30,1, 1, 26).getValues()[0];
  const bcMap = {};
  for (let i = 0; i < 26; i++) {
    const k1 = normalize(n1[i]).toLowerCase();
    if (k1) bcMap[k1] = normalize(b1[i]);
    const k2 = normalize(n2[i]).toLowerCase();
    if (k2) bcMap[k2] = normalize(b2[i]);
  }

  // read your C:D block
  const block = mainRng.getValues(); 
  // block[r][0] = col C (asset), block[r][1] = col D (manual override)

  // compute what to put into D
  const outD = block.map(([asset, manualChain]) => {
    // if no asset or user already typed a chain in D → leave it alone
    if (!asset || String(manualChain).trim() !== '') {
      return [ manualChain ];
    }
    const key = String(asset).trim().toLowerCase();
    return [ bcMap[key] || '' ];
  });

  // write back into **column D** of C4:L32 (that’s offset by 1 column)
  mainRng.offset(0, 1, numRows, 1).setValues(outD);
}

/**
 * Syncs Asset Token Amounts on the "Assets"'s Main Table (E3:E32) with "Spot Trading Log"
 */
function syncAssetTokens() {
  const ss       = SpreadsheetApp.getActive();
  const spot     = ss.getSheetByName(SPOT_SHEET);
  const assets   = ss.getSheetByName(ASSETS_SHEET);
  const mainRng  = ss.getRangeByName(NR.mainTable);  // e.g. C4:L32
  const startRow = mainRng.getRow();
  const startCol = mainRng.getColumn();              // column C = 3
  const numRows  = mainRng.getNumRows();

  if (numRows < 1) return;

  // 1) Build a map of (asset|chain) → deltaTokens across two snapshots
  const lastCol   = spot.getLastColumn();
  const tokenMap  = {};
  const blocks    = [
    { a: 1, c: 3, s: 6, p: 20 },
    { a: 28, c: 30, s: 33, p: 47 }
  ];

  blocks.forEach(({ a, c, s, p }) => {
    const assetsHdr = spot.getRange(a, 1, 1, lastCol).getValues()[0];
    const chainsHdr = spot.getRange(c, 1, 1, lastCol).getValues()[0];
    const snapT     = spot.getRange(s, 1, 1, lastCol).getValues()[0];
    const prevT     = spot.getRange(p, 1, 1, lastCol).getValues()[0];

    assetsHdr.forEach((rawAsset, i) => {
      const rawChain = chainsHdr[i];
      const asset    = normalize(rawAsset);
      const chain    = normalize(rawChain);
      if (!asset || !chain) return;

      const delta    = (Number(snapT[i]) || 0) - (Number(prevT[i]) || 0);
      const key      = `${asset.toLowerCase()}|${chain.toLowerCase()}`;
      tokenMap[key]  = (tokenMap[key] || 0) + delta;
    });
  });

  // 2) Read your main‐table’s Asset+Chain columns (C:D)
  const mainBlock = assets
    .getRange(startRow, startCol, numRows, 2)
    .getValues();  // [[asset, chain], ...]

  // 3) Compute the output array for column F (startCol+2)
  const out = mainBlock.map(([asset, chain]) => {
    if (!asset || !chain) return [''];
    const key = `${String(asset).trim().toLowerCase()}|${String(chain).trim().toLowerCase()}`;
    return [ tokenMap[key] || 0 ];
  });

  // 4) Write it back
  assets
    .getRange(startRow, startCol + 2, numRows, 1)
    .setValues(out);
}

/**
 * Syncs Asset Cost Basis on the "Assets"'s Main Table (G3:G32) with "Spot Trading Log" (Only writes into G when it’s blank; preserves any manual‐cost rows.)
 */
function syncAssetCostBasis() {
  const ss       = SpreadsheetApp.getActive();
  const spot     = ss.getSheetByName(SPOT_SHEET);
  const assets   = ss.getSheetByName(ASSETS_SHEET);
  const mainRng  = ss.getRangeByName(NR.mainTable);    // e.g. C4:L32
  const startRow = mainRng.getRow();
  const startCol = mainRng.getColumn();                // C = 3
  const numRows  = mainRng.getNumRows();

  if (numRows < 1) return;

  // ─── 1) Fetch all relevant Spot rows in one batch ─────────────────────────────
  //    rows 1 (hdr1), 5 (price1), 6 (tokens1), 28 (hdr2), 32 (price2), 33 (tokens2)
  const spotData = spot.getRange(1, 2, 33, 25).getValues();
  //    spotData indices: [0]=hdr1, [4]=price1, [5]=tokens1, [27]=hdr2, [31]=price2, [32]=tokens2

  // ─── 2) Build costMap: { asset.toLowerCase(): totalCost } ────────────────────
  const costMap = {};
  for (let i = 0; i < 25; i++) {
    const a1     = normalize(spotData[0][i]);
    const a2     = normalize(spotData[27][i]);
    const p1     = Number(spotData[4][i])  || 0;
    const t1     = Number(spotData[5][i])  || 0;
    const p2     = Number(spotData[31][i]) || 0;
    const t2     = Number(spotData[32][i]) || 0;
    const batchCost = p1 * t1 + p2 * t2;

    if (a1) {
      const key = a1.toLowerCase();
      costMap[key] = (costMap[key] || 0) + batchCost;
    }
    if (a2) {
      const key = a2.toLowerCase();
      costMap[key] = (costMap[key] || 0) + batchCost;
    }
  }

  // ─── 3) Read your main-table’s columns C–G ([asset,chain,tokens,USD,cost]) ─────
  const mainBlock = assets
    .getRange(startRow, startCol, numRows, 5)
    .getValues();

  // ─── 4) Compute new costBasis (col G = startCol+4) ───────────────────────────
  const outCost = mainBlock.map(([asset, chain, , , oldCost]) => {
    const name = String(asset || '').trim();
    if (!name || String(oldCost).trim() !== '') {
      // blank row or manual override → keep empty or existing
      return [ oldCost !== '' ? oldCost : '' ];
    }
    const key = name.toLowerCase();
    return [ costMap[key] || 0 ];
  });

  // ─── 5) Write back into column G ─────────────────────────────────────────────
  assets
    .getRange(startRow, startCol + 4, numRows, 1)
    .setValues(outCost);
}

/**
 * Syncs Asset Location(s) Tags on the "Assets"'s Main Table (J3:J32) with "Spot Trading Log"
 */
function syncAssetTags() {
  const ss           = SpreadsheetApp.getActive();
  const spotSheet    = ss.getSheetByName(SPOT_SHEET);
  const assetsSheet  = ss.getSheetByName(ASSETS_SHEET);
  const historySheet = ss.getSheetByName(HISTORY_SHEET);
  const mainRange    = ss.getRangeByName(NR.mainTable);

  const startRow = mainRange.getRow();
  const startCol = mainRange.getColumn();   // C
  const numRows  = mainRange.getNumRows();
  if (numRows < 1) return;
  if (!historySheet) throw new Error(`Cannot find sheet: ${HISTORY_SHEET}`);

  // 1) Build spotTags map: { "asset|chain" → Set of tags }
  const spotTags = {};
  [
    { namesRow:  1, tagsRow: 16 },
    { namesRow: 28, tagsRow: 43 }
  ].forEach(({ namesRow, tagsRow }) => {
    const names = spotSheet.getRange(namesRow,  1, 1, 26).getValues()[0];
    const tags  = spotSheet.getRange(tagsRow,   1, 1, 26).getValues()[0];
    names.forEach((rawAsset, i) => {
      const asset = normalize(rawAsset).toLowerCase();
      const tag   = normalize(tags[i]);
      if (asset && tag) {
        spotTags[asset] = spotTags[asset] || new Set();
        spotTags[asset].add(tag);
      }
    });
  });

  // 2) Build manualTags map from C5:H1000 on historySheet
  const manualTags = {};
  historySheet
    .getRange('C5:H1000')
    .getValues()
    .forEach(row => {
      const asset = String(row[0]).trim().toLowerCase();
      const chain = String(row[1]).trim().toLowerCase();
      const tag   = String(row[4]).trim();
      if (asset && chain && tag) {
        const key = `${asset}|${chain}`;
        manualTags[key] = manualTags[key] || new Set();
        manualTags[key].add(tag);
      }
    });

  // 3) Read your main-table’s Asset+Chain columns (C:D)
  const mainBlock = assetsSheet
    .getRange(startRow, startCol, numRows, 2)
    .getValues();  // [[asset, chain], ...]

  // 4) Merge tags for each row
  const out = mainBlock.map(([rawAsset, rawChain]) => {
    const asset = String(rawAsset).trim();
    const chain = String(rawChain).trim();
    if (!asset || !chain) return [''];
    const aKey = asset.toLowerCase();
    const key  = `${aKey}|${chain.toLowerCase()}`;

    const merged = new Set();
    (manualTags[key] || []).forEach(t => merged.add(t));
    (spotTags[aKey]   || []).forEach(t => merged.add(t));

    return [ Array.from(merged).join(', ') ];
  });

  // 5) Write merged tags into column J (C + 7)
  assetsSheet
    .getRange(startRow, startCol + 7, numRows, 1)
    .setValues(out);
}

/**
 * Syncs Asset Total Current Value on the "Assets"'s Main Table (F3:32) with "Spot Trading Log"
 */
function syncAssetTotalValue(){
  const ss       = SpreadsheetApp.getActive();
  const assets   = ss.getSheetByName(ASSETS_SHEET);
  const summary  = ss.getSheetByName(ASSETS_SUMMARY_SHEET);
  const mainRng  = ss.getRangeByName(NR.mainTable);
  const startRow = mainRng.getRow();
  const numRows  = mainRng.getNumRows();

  // 1) Read asset names from C and token amounts from E over your mainTable block
  const names = assets.getRange(startRow, 3, numRows, 1).getValues().flat(); // C4:C32
  const toks  = assets.getRange(startRow, 5, numRows, 1)                    // E4:E32
                       .getValues().flat()
                       .map(x => parseFloat(x) || 0);

  // 2) Build lookup from Summary P21:P33 → R21:R33
  const sumNames  = summary.getRange('P21:P33').getValues().flat()
                           .map(n => String(n).trim().toLowerCase());
  const sumPrices = summary.getRange('R21:R33').getValues().flat()
                           .map(p => parseFloat(p) || 0);

  // 3) Compute USD value = tokens × price
  const out = names.map((n, i) => {
    const key = String(n).trim().toLowerCase();
    if (!key || toks[i] <= 0) return [''];
    const idx = sumNames.indexOf(key);
    return [ idx >= 0 ? toks[i] * sumPrices[idx] : '' ];
  });

  // 4) Write USD values into F (column 6) of your mainTable block
  assets.getRange(startRow, 6, numRows, 1).setValues(out); // F4:F32

  // 5) Re‑sort after updating
  sortAssetsSheet();
}

/**
 * Build a single row for the “bay” table.
 * Now targets C:D:E:F:G:H:I:J:K (columns 3–11).
 */
function buildBayRow(sh, row, dst) {
  const g = c => sh.getRange(row, c).getValue();
  const u  = pct(g(COL_U)),
        ah = pct(g(COL_AH)),
        ax = pct(g(COL_AX));
  const of = Math.max(0, 1 - u - ah - ax);

  // lookup by Asset at column D (dst in C→K block)
  const priceF = `INDEX($R$21:$R$33, MATCH($D${dst}, $P$21:$P$33, 0))`;
  // token cell is now in column H
  const tokRef = `$H${dst}`;

  return [
    // C: Date
    g(COL_DATE),
    // D: Asset
    g(COL_ASSET),
    // E: Dir
    g(COL_DIR),
    // F: collateral*of
    g(COL_COLL) * of,
    // G: leverage
    g(COL_LEV),
    // H: tokens*of
    g(COL_TOK) * of,

    // I: USD Value = tokens * price
    `=IF($D${dst}="","",${tokRef}*IFERROR(${priceF},0))`,

    // J: PnL USD
    `=IF($D${dst}="","",${tokRef}*(IFERROR(${priceF},0)-${g(COL_ENTRY)}))`,

    // K: % PnL
    `=IF($D${dst}="","",${g(COL_LEV)}*(IFERROR(${priceF},0)-${g(COL_ENTRY)})/${g(COL_ENTRY)})`
  ];
}

function rebuildTradeBay() {
  const ss     = SpreadsheetApp.getActive();
  const src    = ss.getSheetByName(PERPS_SHEET);
  const sum    = ss.getSheetByName(ASSETS_SUMMARY_SHEET);
  const posRng = ss.getRangeByName(NR.currentPositions);

  const pr  = posRng.getRow();
  const pc  = posRng.getColumn();
  const nr  = posRng.getNumRows();
  const nc  = posRng.getNumColumns();
  const tol = 1e-4;

  // 1) Collect open trades
  const open = [];
  const last = src.getLastRow();
  for (let r = 3; r <= last; r++) {
    // a) skip zero-size
    if ((Number(src.getRange(r, COL_TOK).getValue()) || 0) === 0) continue;

    // b) inline tradeClosed logic
    const u  = pct(src.getRange(r, COL_U ).getValue());
    const ah = pct(src.getRange(r, COL_AH).getValue());
    const ax = pct(src.getRange(r, COL_AX).getValue());
    if (
      (Math.abs(u - 1)           < tol && notBlank(src, r, COL_R )) ||
      (Math.abs(u + ah - 1)      < tol && notBlank(src, r, COL_AH)) ||
      (Math.abs(u + ah + ax - 1) < tol && notBlank(src, r, COL_AX))
    ) continue;

    // c) skip blank asset
    const asset = src.getRange(r, COL_ASSET).getValue();
    if (!asset) continue;

    // d) collect it
    open.push({ row: r, date: src.getRange(r, COL_DATE).getValue() });
  }

  // 2) Sort newest first, clear old bay
  open.sort((a, b) => (b.date || 0) - (a.date || 0));
  posRng.clearContent();

  // 3) Fill top N slots unconditionally
  open.slice(0, nr).forEach((t, i) => {
    const dstRow = pr + i;
    const bayRow = buildBayRow(src, t.row, dstRow);
    sum.getRange(dstRow, pc, 1, nc).setValues([bayRow]);
  });
}

/**
 * Sort Assets sheet by Value (col F) descending,
 * moving columns C through L together.
 */
function sortAssetsSheet() {
  const ss       = SpreadsheetApp.getActive();
  const sh       = ss.getSheetByName(ASSETS_SHEET);
  const mainRng  = ss.getRangeByName(NR.mainTable); // C4:L32
  const startRow = mainRng.getRow();                // 4
  const startCol = mainRng.getColumn();             // 3
  const numRows  = mainRng.getNumRows();            // 29
  const numCols  = mainRng.getNumColumns();         // 10

  // If the top‐left cell of the block is blank, there's nothing to sort.
  if (!sh.getRange(startRow, startCol).getValue()) return;

  // Sort that exact block by the sheet‐wide column F (6) descending:
  sh
    .getRange(startRow, startCol, numRows, numCols)
    .sort([{ column: 6, ascending: false }]);
}

/**
 * Populate the top 18 assets from Assets sheet into
 * Assets Summary in three vertical 6‑row blocks:
 *
 * Block 1 → rows 1–6 in cols A:B, C:D, E:F, G:H, I:J, K:L  (ranks 1–6)
 * Block 2 → rows 8–13 in same col‑pairs (7–12)
 * Block 3 → rows 15–20                          (13–18)
 *
 * Within each block, down the six rows go:
 *   Asset
 *   Tokens
 *   Value
 *   PnL (USD)
 *   % PnL
 *   Location(s)
 */
function updateTopAssetsSummary() {
  const ss   = SpreadsheetApp.getActive();
  const asht = ss.getSheetByName(ASSETS_SHEET);
  const rng  = ss.getRangeByName(NR.topAssets);  // Assets!C3:N22

  // 1) Pull and sort your top‑18 rows from C4:L…
  const last = asht.getLastRow();
  if (last < 4) return;
  const rows = asht
    .getRange(4, 3, last - 3, 10)   // C4:L…
    .getValues()
    .filter(r => r[0] && Number(r[3]) > 0)
    .sort((a, b) => b[3] - a[3])
    .slice(0, 18);

  // 2) Clear out the entire 20×12 block C3:N22
  rng.clearContent();

  // 3) Define your labels & data‑indexes
  const FIELDS   = ['Asset','Tokens','Value','PnL (USD)','% PnL','Location(s)'];
  const DATA_IDX = [   0,      2,     3,        5,        6,            7   ];
  // each asset uses a 6‑row × 2‑col sub‑block inside C3:N22

  // 4) Populate into Top_18_Assets
  rows.forEach((rdata, idx) => {
    const block = Math.floor(idx / 6);  // which vertical block: 0,1,2
    const slot  = idx % 6;              // which horizontal slot: 0–5

    // rowOffset: within the 20 rows, each block is 6 rows tall
    const baseRowOffset = block * 6;

    // colOffset: each slot is two columns wide
    const baseColOffset = slot * 2;

    FIELDS.forEach((label, f) => {
      // f = 0…5 corresponds to the six rows in each block
      const cell = rng
        .offset(baseRowOffset + f, baseColOffset, 1, 1);

      // write label in the left-hand column
      cell.setValue(label);

      // write the data value in the column immediately to the right
      cell.offset(0, 1).setValue(rdata[ DATA_IDX[f] ]);
    });
  });
}

/**
 * ─── MERGE SPOT + MANUAL HISTORY (no Sets) ─────────────────────────────────
 * Only writes back columns:
 *   • E (tokens)
 *   • G (cost basis)
 *   • J (tags)
 * preserving any formulas in H & I—and overriding E when the
 * "Hyperliquid (Perps)" tag is present.
 */
function mergeSpotAndManualHistory() {
  const ss           = SpreadsheetApp.getActive();
  const sh           = ss.getSheetByName(ASSETS_SHEET);
  const spot         = ss.getSheetByName(SPOT_SHEET);
  const perpsBalance = ss.getSheetByName(TRADING_SUMMARY_SHEET)
                         .getRange('G4').getValue();

  // 1) Build spotTags lookup: { "asset|chain" → [tag,tag,…] }
  const spotTags = {};
  [
    { namesRow:  1, tagsRow: 16 },
    { namesRow: 28, tagsRow: 43 }
  ].forEach(({ namesRow, tagsRow }) => {
    const names = spot.getRange(namesRow, 1, 1, spot.getLastColumn()).getValues()[0];
    const tags  = spot.getRange(tagsRow,  1, 1, spot.getLastColumn()).getValues()[0];
    names.forEach((rawAsset, i) => {
      const asset = normalize(rawAsset).toLowerCase();
      const tag   = normalize(tags[i]);
      if (!asset || !tag) return;
      const key = `${asset}|${normalize(tags[i]).toLowerCase()}`;
      if (!spotTags[asset]) spotTags[asset] = [];
      if (spotTags[asset].indexOf(tag) === -1) {
        spotTags[asset].push(tag);
      }
    });
  });

  // 2) Read manual history (hard-coded C5:H1000)
  const historySheet = ss.getSheetByName(HISTORY_SHEET);
  if (!historySheet) throw new Error(`Cannot find sheet: ${HISTORY_SHEET}`);
  const history = historySheet.getRange('C5:H1000').getValues(); // [asset,chain,tok,cost,tag,date]

  const manualTok  = {};
  const manualCost = {};
  const manualTags = {};

  history.forEach(row => {
    const rawA = String(row[0] || '').trim();
    const rawB = String(row[1] || '').trim();
    if (!rawA || !rawB) return;
    const asset = rawA.toUpperCase();
    const chain = rawB;
    const key   = (asset + '|' + chain).toLowerCase();

    manualTok[key]  = (manualTok[key]  || 0) + (Number(row[2]) || 0);
    manualCost[key] = (manualCost[key] || 0) + (Number(row[3]) || 0);

    const tag = String(row[4]).trim();
    if (tag) {
      manualTags[key] = manualTags[key] || [];
      if (manualTags[key].indexOf(tag) === -1) {
        manualTags[key].push(tag);
      }
    }
  });

  // 2.5) Add any new <asset|chain> from history into Assets_Main
  const mainRng   = ss.getRangeByName(NR.mainTable);
  const startRow  = mainRng.getRow();
  const startCol  = mainRng.getColumn();
  const numRows   = mainRng.getNumRows();
  const mainPairs = sh
    .getRange(startRow, startCol, numRows, 2)
    .getValues()
    .map(r => [String(r[0]||'').trim(), String(r[1]||'').trim()].join('|').toLowerCase());

  Object.keys(manualTok).forEach(key => {
    if (!mainPairs.includes(key)) {
      const blankIdx = mainPairs.findIndex(p => p === '|' /* both empty */);
      let targetRow;
      if (blankIdx >= 0) {
        targetRow = startRow + blankIdx;
      } else {
        sh.insertRowBefore(startRow);
        targetRow = startRow;
      }
      const [asset, chain] = key.split('|');
      sh.getRange(targetRow, startCol   ).setValue(asset.toUpperCase());
      const chainCap = chain.charAt(0).toUpperCase() + chain.slice(1).toLowerCase();
      sh.getRange(targetRow, startCol+1 ).setValue(chainCap);
      mainPairs.splice(blankIdx >= 0 ? blankIdx : 0, 0, key);
    }
  });

  // 3) Read mainTable into block
  const numCols = mainRng.getNumColumns();
  const block   = sh
    .getRange(startRow, startCol, numRows, numCols)
    .getValues();

  // 4) Merge tokens, costs & tags
  const outTokens = [];
  const outCost   = [];
  const outTags   = [];

  block.forEach(row => {
    const [ asset, chain, oldTokRaw, , oldCostRaw, , , existingRaw ] = row;
    const oldTok   = Number(oldTokRaw)  || 0;
    const oldCost  = Number(oldCostRaw) || 0;
    const existing = String(existingRaw||'')
                     .split(',').map(t=>t.trim()).filter(Boolean);

    if (!asset || !chain) {
      outTokens.push(['']);
      outCost.push(['']);
      outTags.push(['']);
      return;
    }

    const key    = `${asset}|${chain}`.toLowerCase();
    const merged = existing.slice();
    const skey   = asset.toLowerCase();

    (spotTags[skey]    || []).forEach(t => { if (!merged.includes(t)) merged.push(t); });
    (manualTags[key]   || []).forEach(t => { if (!merged.includes(t)) merged.push(t); });

    const hasPerps = merged.includes('Hyperliquid (Perps)');
    const newTok   = hasPerps ? perpsBalance : oldTok + (manualTok[key] || 0);
    const newCost  = oldCost  + (manualCost[key] || 0);

    outTokens.push([ newTok ]);
    outCost.push([ newCost ]);
    outTags.push([ merged.join(', ') ]);
  });

  // 5) Write back into E, G and J
  sh.getRange(startRow, startCol + 2, numRows, 1).setValues(outTokens);
  sh.getRange(startRow, startCol + 4, numRows, 1).setValues(outCost);
  sh.getRange(startRow, startCol + 7, numRows, 1).setValues(outTags);
}

/**
 * Re‑apply PnL, % PnL, first‑purchase date and age
 * into columns H, I, K and L of your Assets_Main block (C4:L32).
 */
function syncAssetFormulas() {
  const ss       = SpreadsheetApp.getActive();
  const sh       = ss.getSheetByName(ASSETS_SHEET);
  const mainRng  = ss.getRangeByName(NR.mainTable); // C4:L32
  const startRow = mainRng.getRow();                // 4
  const numRows  = mainRng.getNumRows();            // 29

  if (numRows < 1) return;

  // Build four arrays of R1C1‑style formulas:
  const fH = [], fI = [], fK = [], fL = [];
  for (let i = 0; i < numRows; i++) {
    // For each relative row in the block...
    // RC[-5] is column C → asset name
    // RC[-2] is column F → USD value
    // RC[-1] is column G → cost basis
    // RC[-8] is column C again (for spot lookup)
    // RC[-7] is column D → blockchain
    const guard = `IF(RC[-5]="","",`; // only run if C not blank

    // H: PnL USD = F – G
    fH.push([ guard + `IF(AND(RC[-2]<>"",RC[-1]<>""),RC[-2]-RC[-1],"") )` ]);

    // I: % PnL = (F–G)/G
    fI.push([ guard + `IF(AND(RC[-1]<>"",RC[-2]<>"",RC[-2]<>0),RC[-1]/RC[-2],"") )` ]);

    // K: first‑purchase date = MIN( 2 spot lookups, manual history lookup )
    //   history is in N9:S32, date now in S9:S32 → that's R9C19:R32C19 in R1C1
    //   filter on N9:N32 = RC[-8] (asset) and O9:O32 = RC[-7] (chain)
    fK.push([ guard +
      `MIN(` +
        `IFERROR(INDEX('Spot Trading Log'!R4C1:R4C26, MATCH(RC[-8],'Spot Trading Log'!R1C1:R1C26,0)),9.99E+307),` +
        `IFERROR(INDEX('Spot Trading Log'!R31C1:R31C26, MATCH(RC[-8],'Spot Trading Log'!R28C1:R28C26,0)),9.99E+307),` +
        `IFERROR(` +
          `MIN(` +
            `FILTER(R9C19:R32C19, R9C14:R32C14=RC[-8], R9C15:R32C15=RC[-7])` +
          `),9.99E+307)` +
      `)` +
    `)` ]);

    // L: age since that date
    //   TEXT(INT(NOW()–K),…) etc.
    fL.push([ guard +
      `TEXT(INT(NOW()-RC[-1]),"00")&":"&` +
      `TEXT(HOUR(NOW()-RC[-1]),"00")&":"&` +
      `TEXT(MINUTE(NOW()-RC[-1]),"00")` +
    `)` ]);
  }

  // Write them into H4:H32, I4:I32, K4:K32 and L4:L32
  sh.getRange(startRow,  8, numRows, 1).setFormulasR1C1(fH);
  sh.getRange(startRow,  9, numRows, 1).setFormulasR1C1(fI);
  sh.getRange(startRow, 11, numRows, 1).setFormulasR1C1(fK);
  sh.getRange(startRow, 12, numRows, 1).setFormulasR1C1(fL);
}

/**
 * Whenever N5:R5 is fully filled (or tagged “Hyperliquid (Perps)”):
 *  • If tag=Hyperliquid (Perps), grab live balance from Trading Summary!G4 and use it as tok
 *  • Archive that input into your full‑history sheet (HISTORY_SHEET)
 *  • Shift N5:R5 → N9:S9, pushing the rest of N9:S32 down one (24‑row in‑sheet history)
 *  • Clear N5:R5
 */
function recordManualHistory() {
  const ss       = SpreadsheetApp.getActive();
  const assets   = ss.getSheetByName(ASSETS_SHEET);

  // <-- NEW: grab asset, chain, tok, cost, tag from N5:R5 -->
  const inputRng = ss.getRangeByName(NR.inputRow);
  const [assetRaw, chainRaw, tokRawRaw, costRawRaw, tagRaw] =
    inputRng.getValues()[0];

  const asset   = String(assetRaw ).trim();
  const chain   = String(chainRaw ).trim();
  const tokRaw  = String(tokRawRaw).trim();
  const costRaw = String(costRawRaw).trim();
  const tag     = String(tagRaw   ).trim();

  // 2) bail unless fully filled OR perps tag
  if (!( (asset && chain && tokRaw && costRaw && tag)
      || tag === 'Hyperliquid (Perps)' )) {
    return;
  }

  // 3) resolve “Perps”-override
  let tok = tokRaw;
  if (tag === 'Hyperliquid (Perps)') {
    tok = String(
      ss.getSheetByName(TRADING_SUMMARY_SHEET)
        .getRange('G4').getValue()
    );
  }

// 4) FULL-ARCHIVE → shift row 5 down, then write into C5:H5
  const fullSheet = ss.getRangeByName('Full_Assets_Manual_Input_History').getSheet();

  // ── shift existing row 5 (C5:H5) into row 6 (and push everything below down) ──
  fullSheet.insertRowBefore(5);

  // now row 5 is empty—stamp your new record there:
  fullSheet
    .getRange('C5:H5')
    .setValues([[ asset, chain, tok, costRaw, tag, new Date() ]]);

  // 5) IN-SHEET 24-ROW HISTORY → shift into your Assets_History named range
  const histRng  = ss.getRangeByName(NR.historyTable);
  const histVals = histRng.getValues();
  histVals.pop();   // drop the oldest
  histVals.unshift([ asset, chain, tok, costRaw, tag, new Date() ]);
  histRng.setValues(histVals);
}

/**
 * Fetch latest USD prices for your watchlist
 * and write into the rightmost column of Asset_Watchlist.
 */
function refreshAllCMCPrices() {
  const ss       = SpreadsheetApp.getActive();
  const sum      = ss.getSheetByName(ASSETS_SUMMARY_SHEET);
  const watchRng = ss.getRangeByName(NR.watchlist);     // Assets!P21:R33

  // 1) Read symbols from column P of the watchlist block
  const watchVals = watchRng.getValues();               // [ [sym, …, price], … ]
  const syms = watchVals
    .map(row => String(row[0]).trim().toUpperCase())
    .filter(s => s);

  if (syms.length === 0) return;

  // 2) Query CMC API
  const url = 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest';
  const params = {
    method: 'get',
    headers: { 'X-CMC_PRO_API_KEY': CMC_API_KEY },
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch(url + '?symbol=' + syms.join(','), params);
  const payload  = JSON.parse(response.getContentText());

  // 3) Map each symbol → USD price
  const prices = syms.map(sym => {
    const entry = payload.data && payload.data[sym];
    return (entry && entry.quote && entry.quote.USD && entry.quote.USD.price) || 0;
  });

  // 4) Write prices back into column R of Asset_Watchlist
  //    that's an offset of 0 rows, 2 cols from P
  const priceRng = watchRng.offset(0, 2, prices.length, 1);
  priceRng.setValues(prices.map(p => [p]));

  // 5) Trigger downstream updates
  overridePerpsTokens();
  syncAssetTotalValue();
  updateTopAssetsSummary();
}

function overridePerpsTokens() {
  Logger.log('🔥 overridePerpsTokens() fired');

  const ss           = SpreadsheetApp.getActive();
  const sh           = ss.getSheetByName(ASSETS_SHEET);
  const perpsSheet   = ss.getSheetByName(TRADING_SUMMARY_SHEET);
  const livePerpsBal = perpsSheet.getRange('G4').getValue();

  // ── 1) Pull manual history (C5:H1000) ───────────────────────────────────────
  const historySheet = ss.getSheetByName(HISTORY_SHEET);
  if (!historySheet) throw new Error(`Sheet not found: ${HISTORY_SHEET}`);

  const history = historySheet
    .getRange('C5:H1000')
    .getValues(); // [Asset, Chain, Tok, Cost, Tag, Date]

  // ── 2) Build set of keys to override ─────────────────────────────────────────
  //    (only rows tagged 'Hyperliquid (Perps)')
  const perpsKeys = new Set();
  history.forEach(row => {
    const [rawAsset, rawChain, , , rawTag] = row;
    if (String(rawTag||'').trim() === 'Hyperliquid (Perps)') {
      const asset = String(rawAsset||'').trim().toLowerCase();
      const chain = String(rawChain||'').trim().toLowerCase();
      if (asset && chain) {
        perpsKeys.add(`${asset}|${chain}`);
      }
    }
  });

  // ── 3) Read your main-table C:D:E (Asset, Chain, Tokens) ────────────────────
  const mainRng   = ss.getRangeByName(NR.mainTable);
  const mRow      = mainRng.getRow();
  const mCol      = mainRng.getColumn();    // col C = 3
  const mNumRows  = mainRng.getNumRows();
  const mainVals  = sh.getRange(mRow, mCol, mNumRows, 3).getValues();

  // ── 4) Loop and only overwrite column E where the key matches ───────────────
  const eCol = mCol + 2; // column E
  mainVals.forEach((row, i) => {
    const asset = String(row[0]||'').trim().toLowerCase();
    const chain = String(row[1]||'').trim().toLowerCase();
    const key   = `${asset}|${chain}`;
    if (perpsKeys.has(key)) {
      // only override this one cell
      sh.getRange(mRow + i, eCol).setValue(livePerpsBal);
    }
  });

  Logger.log('✅ Finished overridePerpsTokens');
}

/**
 * In your manual‐history area (L7:P30), whenever
 * column P has the tag "Hyperliquid (Perps)",
 * overwrite the token count in column N with
 * the live Trading Summary!G4 value.
 */
function overrideHistoryTokens() {
  const ss            = SpreadsheetApp.getActive();
  const assetsSheet   = ss.getSheetByName(ASSETS_SHEET);
  const historySheet  = ss.getSheetByName(HISTORY_SHEET);
  const perpsSheet    = ss.getSheetByName(TRADING_SUMMARY_SHEET);
  const livePerpsBal  = perpsSheet.getRange('G4').getValue();

  // Define the two target ranges:
  const targets = [
    { 
      sheet: assetsSheet, 
      rangeA1: 'N9:S32', 
      desc: 'Assets-sheet manual history (N9:S32)' 
    },
    { 
      sheet: historySheet, 
      rangeA1: 'C5:H1000', 
      desc: 'Full manual-history sheet (C5:H1000)' 
    }
  ];

  targets.forEach(({ sheet, rangeA1, desc }) => {
    if (!sheet) {
      throw new Error(`Cannot find sheet for ${desc}`);
    }
    const rng   = sheet.getRange(rangeA1);
    const data  = rng.getValues(); 
    let changed = false;

    // data[i] is [Asset, Chain, Tokens, Cost, Tag, …]
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][4]||'').trim() === 'Hyperliquid (Perps)') {
        data[i][2] = livePerpsBal;  // overwrite Tokens
        changed = true;
      }
    }

    if (changed) {
      rng.setValues(data);
    }
  });
}

function updateSolValueInTradingSummary() {
  const ss           = SpreadsheetApp.getActive();
  const assetsSheet  = ss.getSheetByName(ASSETS_SUMMARY_SHEET);
  const tradingSheet = ss.getSheetByName(TRADING_SUMMARY_SHEET);

  // 1) Fetch lamports via JSON‑RPC getBalance
  const payload = {
    jsonrpc: "2.0",
    id:      1,
    method:  "getBalance",
    params: [ WALLET_1 ]
  };
  const resp = UrlFetchApp.fetch(
    HELIUS_RPC_URL,
    {
      method:      'post',
      contentType: 'application/json',
      payload:     JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );
  const js = JSON.parse(resp.getContentText());
  if (!js.result || js.result.value == null) {
    throw new Error(
      'Could not fetch SOL balance for ' + WALLET_1 
      + '. Full response: ' + resp.getContentText()
    );
  }
  const solBalance = js.result.value / LAMPORTS_PER_SOL;

  // 2) Find SOL price in Assets Summary!P21:P33 → R21:R33
  const names  = assetsSheet.getRange('P21:P33').getValues().flat();
  const prices = assetsSheet.getRange('R21:R33').getValues().flat();
  const solPrice = names.reduce((acc, val, i) => val === 'SOL' ? prices[i] : acc, 0);

  if (solPrice === 0) {
    throw new Error('SOL price not found in ' + ASSETS_SUMMARY_SHEET + '!N19:N31');
  }

  // 3) Multiply and write to Trading Summary!L3
  const totalValue = solBalance * solPrice;
  tradingSheet.getRange('L3').setValue(totalValue);
}

/**
 * @customfunction
 * Calculates average time between your first trade date and
 * (last trade date + 7 days), divided by the number of trades.
 * Returns "[d]:hh:mm" or an error message.
 */
function TradeFrequency() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sh      = ss.getSheetByName('Memecoin Trading Data Table');
  if (!sh) return "Sheet not found";

  const lastRow = sh.getLastRow();
  if (lastRow < 3) return "No trade data available";

  // only pull rows 2..lastRow
  const dates = sh.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const times = sh.getRange(2, 2, lastRow - 1, 1).getValues().flat();

  // build JS Date objects only where both exist
  const datetimes = dates
    .map((d,i) => ({ d, t: times[i] }))
    .filter(x => x.d instanceof Date && x.t instanceof Date)
    .map(x => new Date(
      x.d.getFullYear(), x.d.getMonth(), x.d.getDate(),
      x.t.getHours(),    x.t.getMinutes(), x.t.getSeconds()
    ))
    .sort((a,b) => a - b);

  if (datetimes.length < 2) 
    return "Not enough active trades to calculate meaningful frequency";

  const first = datetimes[0];
  const last  = datetimes[datetimes.length - 1];
  const end   = new Date(last.getTime() + 7*24*60*60*1000);
  const count = datetimes.filter(d => d >= first && d <= end).length;
  const span  = end - first;

  if (count <= 1 || span <= 0) 
    return "Not enough active trades to calculate meaningful frequency";

  // average ms per trade
  const avgMs     = span / count;
  const totalDays = Math.floor(avgMs / (1000*60*60*24));
  const remMs     = avgMs - totalDays*24*60*60*1000;
  const hrs       = Math.floor(remMs / (1000*60*60));
  const mins      = Math.floor((remMs - hrs*60*60*1000) / (1000*60));

  return Utilities.formatString("%d:%02d:%02d", totalDays, hrs, mins);
}

/**
 * @customfunction
 * Calculates average time between trades on the
 * Perps Data Table sheet, over a 7-day window after the last trade.
 * Returns "[d]:hh:mm" or an explanatory message.
 */
function PerpsFrequency() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sh      = ss.getSheetByName('Perps Data Table');
  if (!sh) return 'Sheet "Perps Data Table" not found';

  const lastRow = sh.getLastRow();
  if (lastRow < 3) return "No trade data available";

  // only pull rows 2..lastRow
  const dates = sh.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const times = sh.getRange(2, 2, lastRow - 1, 1).getValues().flat();

  // build JS Date objects only where both exist
  const datetimes = dates
    .map((d,i) => ({ d, t: times[i] }))
    .filter(x => x.d instanceof Date && x.t instanceof Date)
    .map(x => new Date(
      x.d.getFullYear(), x.d.getMonth(), x.d.getDate(),
      x.t.getHours(),    x.t.getMinutes(), x.t.getSeconds()
    ))
    .sort((a,b) => a - b);

  if (datetimes.length < 2) 
    return "Not enough active trades to calculate meaningful frequency";

  const first = datetimes[0];
  const last  = datetimes[datetimes.length - 1];
  const end   = new Date(last.getTime() + 7*24*60*60*1000);
  const count = datetimes.filter(d => d >= first && d <= end).length;
  const span  = end - first;

  if (count <= 1 || span <= 0) 
    return "Not enough active trades to calculate meaningful frequency";

  // average milliseconds per trade
  const avgMs   = span / count;
  const days    = Math.floor(avgMs / (24*60*60*1000));
  const remMs   = avgMs - days * 24*60*60*1000;
  const hours   = Math.floor(remMs / (60*60*1000));
  const minutes = Math.floor((remMs - hours*60*60*1000) / (60*1000));

  return Utilities.formatString("%d:%02d:%02d", days, hours, minutes);
}

/**
 * @customfunction
 * Calculates average number of trades per week
 * on the Perps Data Table sheet.
 * Returns a number (trades/week) or a message if no data.
 */
function Perps_Avg_Trades_Per_Week() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sh    = ss.getSheetByName('Perps Data Table');
  if (!sh) return 'Sheet "Perps Data Table" not found';
  
  // pull all dates from A2:A
  const rawDates = sh.getRange('A2:A').getValues().flat();
  // filter to only actual dates
  const dates = rawDates
    .filter(d => d instanceof Date)
    // normalize to midnight to ignore time
    .map(d => new Date(d.getFullYear(), d.getMonth(), d.getDate()))
    .sort((a,b) => a - b);
  
  const tradeCount = dates.length;
  if (tradeCount === 0) return 'No trade dates found';
  
  const firstDate = dates[0];
  const lastDate  = dates[dates.length - 1];
  const msPerDay  = 24 * 60 * 60 * 1000;
  // inclusive day count
  const daySpan   = Math.round((lastDate - firstDate) / msPerDay) + 1;
  
  // convert days to weeks
  const weekSpan  = daySpan / 7;
  
  return tradeCount / weekSpan;
}

/**
 * @customfunction
 * Calculates average number of trades per month
 * on the Spot Trading Log sheet, using combined date and time
 * found in row 4.
 * Returns a number (trades/month) or a message if no data.
 */
function Spot_Avg_Trades_Per_Month() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Spot Trading Log');
  if (!sh) return 'Sheet "Spot Trading Log" not found';

  // retrieve all datetime values from row 4 (combined date & time)
  const lastCol       = sh.getLastColumn();
  const rawDateTimes  = sh.getRange(4, 1, 1, lastCol).getValues()[0];

  // filter only valid Date objects
  const datetimes = rawDateTimes
    .filter(d => d instanceof Date)
    .sort((a, b) => a - b);

  if (datetimes.length === 0) return 'No trade date/time found';

  // compute spans
  const first      = datetimes[0];
  const last       = datetimes[datetimes.length - 1];
  const tradeCount = datetimes.length;

  // calculate inclusive month span
  const yearDiff  = last.getFullYear() - first.getFullYear();
  const monthDiff = last.getMonth() - first.getMonth();
  const monthSpan = yearDiff * 12 + monthDiff + 1;

  return tradeCount / monthSpan;
}

/**
 * @customfunction
 * Calculates frequency of spot trades formatted [d]:hh:mm
 * on the Spot Trading Log sheet, using combined date and time
 * found in row 4.
 * Returns the average interval between trades over a 7-day window
 * after the last trade, formatted as DD:HH:MM, or an explanatory message.
 */
function Spot_Trade_Frequency() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sh    = ss.getSheetByName('Spot Trading Log');
  if (!sh) return 'Sheet "Spot Trading Log" not found';

  // pull all datetime values from row 4 (combined date & time)
  const lastCol      = sh.getLastColumn();
  const rawDateTimes = sh.getRange(4, 1, 1, lastCol).getValues()[0];
  // filter only valid Date objects
  const datetimes    = rawDateTimes
    .filter(d => d instanceof Date)
    .sort((a, b) => a - b);

  const count = datetimes.length;
  if (count === 0) return 'No trade date/time found';

  const first = datetimes[0];
  const last  = datetimes[count - 1];
  // extend end date by 7 days
  const end   = new Date(last.getTime() + 7 * 24 * 60 * 60 * 1000);

  // count trades within this window
  const withinCount = datetimes.filter(d => d >= first && d <= end).length;
  const durationMs  = end - first;

  if (withinCount <= 1 || durationMs <= 0)
    return 'Not enough active trades to calculate meaningful frequency';

  const avgMs   = durationMs / withinCount;
  const days    = Math.floor(avgMs / (24 * 60 * 60 * 1000));
  const remMs   = avgMs - days * 24 * 60 * 60 * 1000;
  const hours   = Math.floor(remMs / (60 * 60 * 1000));
  const minutes = Math.floor((remMs - hours * 60 * 60 * 1000) / (60 * 1000));

  return Utilities.formatString("%d:%02d:%02d", days, hours, minutes);
}

function recordDailyTotal() {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const summary = ss.getSheetByName('Assets Summary');
    const history = ss.getSheetByName('Assets Total History');
    if (!summary || !history) {
      throw new Error("Make sure both tabs exist: 'Assets Summary' & 'Assets Total History'");
    }

    // 1) Read your total and timestamp
    const total = summary.getRange('D24').getValue();
    const now   = new Date();

    // 2) Figure out which two columns to use this month
    //    April (4) → A/B (1,2)
    //    May   (5) → C/D (3,4)
    //    June  (6) → E/F (5,6)
    //    July  (7) → G/H (7,8)
    const month = now.getMonth() + 1;
    const startCol = (month - 4) * 2 + 1;
    const colDate  = startCol;
    const colTotal = startCol + 1;
    if (startCol < 1 || colTotal > history.getMaxColumns()) {
      throw new Error(`No columns set up for month=${month}`);
    }

    // 3) Scan down from row 2 in the date‑column to find the first blank
    const lastRow = history.getMaxRows();
    const dateVals = history
      .getRange(2, colDate, lastRow - 1, 1)
      .getValues()
      .flat();
    const nextIdx = dateVals.findIndex(cell => cell === '' || cell === null);

    // 4) Compute the actual row to write into
    //    If we found an empty slot, row = 2 + index; otherwise append at bottom+1
    const targetRow = nextIdx >= 0
      ? 2 + nextIdx
      : lastRow + 1;

    // 5) Write timestamp and total
    history.getRange(targetRow, colDate ).setValue(now);
    history.getRange(targetRow, colTotal).setValue(total);

    // Optional: log to Stackdriver so you can inspect later
    console.log(`Wrote ${now} → col ${colDate}, and ${total} → col ${colTotal}, row ${targetRow}`);
  }
  catch (e) {
    // If you run manually you'll see a pop‑up with the exact message
    SpreadsheetApp.getUi().alert('recordDailyTotal error:\n' + e.message);
    throw e;
  }
}
