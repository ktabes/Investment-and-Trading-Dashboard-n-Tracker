// HYPERLIQUID PERPS - MULTI-ASSET TABLE REBUILD (ASCII-safe)E
// Requires V8 runtime. File is self-contained. Create a sheet with asset names in:
//   - C3, Q3, AE3, AS3, etc. (every 14 columns)
//   - Data tables start at row 5 for each asset

'use strict';

const HL_API  = 'https://api.hyperliquid.xyz/info';
const HL_USER = __SP.getProperty('HL_USER');

const SHEET_NAME    = 'Test Perps Data Table'; 
const TZ            = __SP.getProperty('TIME_ZONE');
const LOOKBACK_DAYS = 180;        // change as needed
const FEES_INCLUDE_SPOT = false;  // fees helper only

// Table configuration - asset names are in these columns, data starts 2 rows below
const ASSET_NAME_ROW = 3;
const DATA_START_ROW = 5;
const TABLE_WIDTH = 13;  // C to O = 13 columns
const TABLE_SPACING = 14; // Gap between tables (C to Q = 14 column gap)

// --------------------------------- QUICK PROBES ---------------------------------
function _ping() {
  Logger.log('Ping start ' + new Date().toISOString());
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sh.getRange('A1').setValue('Ping @ ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'));
  Logger.log('Ping wrote to sheet successfully.');
}

function _probeHL() {
  Logger.log('HL probe start');
  const res = hlPost_({ type: 'clearinghouseState', user: HL_USER });
  Logger.log('HL keys: ' + JSON.stringify(Object.keys(res || {})));
}

// --------------------------------- PUBLIC RUNNERS ---------------------------------
function Rebuild_Perps_Data_Table() {
  Logger.log('Rebuild start ' + new Date().toISOString());
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + SHEET_NAME + '" not found');

  // Find asset tables and their positions
  const assetTables = findAssetTables_(sh);
  Logger.log('Found asset tables: ' + JSON.stringify(assetTables.map(function(t){ return t.asset; })));
  
  if (assetTables.length === 0) {
    Logger.log('No asset tables found. Make sure asset names are in cells C3, Q3, AE3, etc.');
    return;
  }

  // Fetch fills (perps-only), ASC
  const win   = getWindowMs_(LOOKBACK_DAYS);
  const fills = getPerpFillsAsc_(HL_USER, win.startMs, win.endMs);
  Logger.log('fills kept (perps-only): ' + fills.length);

  if (fills.length === 0) {
    Logger.log('No fills in window; clearing all tables.');
    clearAssetTables_(sh, assetTables);
    return;
  }

  // Live state & orders
  const positionData = getCurrentLiqPxMap_();
  const liqMap = positionData.liquidation;
  const leverageMap = positionData.leverage;
  const collateralMap = positionData.collateral;
  const openOrdMap = getOpenOrdersMap_();

  // Funding events by coin (for fee column)
  const fundingByCoin = getFundingEventsByCoin_(HL_USER, win.startMs, win.endMs);
  
  // Group fills by asset and annotate
  const fillsByAsset = groupFillsByAsset_(fills, assetTables);
  const annotatedByAsset = {};
  Object.keys(fillsByAsset).forEach(function(asset){
    annotatedByAsset[asset] = annotateAssetFills_(fillsByAsset[asset], fundingByCoin[asset] || []);
  });
  
  // Clear all tables first
  clearAssetTables_(sh, assetTables);
  
  // Process each asset table
  for (var i = 0; i < assetTables.length; i++) {
    const table = assetTables[i];
    const assetFills = annotatedByAsset[table.asset] || [];
    const mergedFills = mergeFillsByTimestamp_(assetFills);
    
    if (mergedFills.length === 0) {
      Logger.log('No fills for asset: ' + table.asset);
      continue;
    }
    
    Logger.log('Processing ' + mergedFills.length + ' merged trades for asset: ' + table.asset);
    
    // Convert merged fills to rows for this asset
    const rows = mergedFills.map(function(fill) {
      return fillToTableRow_(fill, liqMap, leverageMap, collateralMap, openOrdMap);
    });

    // Sort by time (newest first)
    rows.sort(function(a,b){ return b[0] - a[0]; });

    const displayRows = rows.map(formatRowDisplay_);

    // (moved up) Determine dynamic table width and inject formulas

    // Pad or trim to width
    const startCol = table.startCol;
    const endCol = table.endCol;
    // We always output 17 columns for the data table; clamp to available space if narrower
    const OUTPUT_COLS = 17;
    const width = Math.min(Math.max(1, endCol - startCol + 1), OUTPUT_COLS);

    // Inject R:R spreadsheet formula per row so it auto-updates if TP/SL edit in sheet
    (function injectRRFormulas(){
      const RR_COL_IDX = 14; // zero-based index in our output array
      const priceCol = startCol + 5;    // Entry/Exit Price column index in sheet
      const dirCol   = startCol + 4;    // Direction (Long/Short)
      const tpCol    = startCol + 11;   // Take Profit
      const slCol    = startCol + 12;   // Stop Loss
      for (var r = 0; r < displayRows.length; r++) {
        var rowNum = DATA_START_ROW + r;
        var a1P  = colToA1_(priceCol) + rowNum;
        var a1D  = colToA1_(dirCol)   + rowNum;
        var a1TP = colToA1_(tpCol)    + rowNum;
        var a1SL = colToA1_(slCol)    + rowNum;
        var formula = "=IF(OR(NOT(ISNUMBER("+a1P+")),NOT(ISNUMBER("+a1TP+")),NOT(ISNUMBER("+a1SL+"))),\"\",IF("+a1D+"=\"Long\",IF(("+a1P+"-"+a1SL+")>0,ROUND(("+a1TP+"-"+a1P+")/("+a1P+"-"+a1SL+"),2),\"\"),IF(("+a1SL+"-"+a1P+")>0,ROUND(("+a1P+"-"+a1TP+")/("+a1SL+"-"+a1P+"),2),\"\")))";
        displayRows[r][RR_COL_IDX] = formula;
      }
    })();

    // Pad or trim to width
    const normalizedRows = displayRows.map(function(r){
      if (r.length > width) return r.slice(0, width);
      if (r.length < width) return r.concat(Array(width - r.length).fill(''));
      return r;
    });
    
    const startRow = DATA_START_ROW;
    sh.getRange(startRow, startCol, normalizedRows.length, width).setValues(normalizedRows);
    // Format PnL (%) column as percent if present
    if (width >= 17) {
      sh.getRange(startRow, startCol + 16, normalizedRows.length, 1).setNumberFormat('0.00%');
    }
    Logger.log('Wrote ' + normalizedRows.length + ' rows for asset: ' + table.asset + ' width=' + width);
  }
  
  Logger.log('Rebuild complete.');
}

// --------------------------------- TABLE MANAGEMENT ---------------------------------
function findAssetTables_(sh) {
  const tables = [];
  const maxCols = sh.getMaxColumns();

  // Read the entire header row once; detect any non-empty string as an asset name starting from col 3
  var row;
  try { row = sh.getRange(ASSET_NAME_ROW, 1, 1, maxCols).getValues()[0]; }
  catch(e) { return tables; }

  const starts = [];
  for (var col = 3; col <= maxCols; col++) {
    const v = row[col - 1];
    if (typeof v === 'string' && v.trim()) {
      starts.push({ col: col, name: v.trim() });
    }
  }
  if (starts.length === 0) return tables;

  // Build table objects; endCol is just before the next start, last one clamps to maxCols
  for (var i = 0; i < starts.length; i++) {
    const startCol = starts[i].col;
    const asset = starts[i].name;
    var endCol;
    if (i + 1 < starts.length) endCol = Math.max(startCol, starts[i+1].col - 1);
    else endCol = Math.min(maxCols, startCol + 32); // give the last table ample room without spanning the whole sheet
    tables.push({ asset: asset, startCol: startCol, endCol: endCol });
  }
  return tables;
}

function clearAssetTables_(sh, assetTables) {
  for (var i = 0; i < assetTables.length; i++) {
    const table = assetTables[i];
    const lastRow = sh.getLastRow();
    if (lastRow >= DATA_START_ROW) {
      const rowsToClear = lastRow - DATA_START_ROW + 1;
      const width = Math.max(1, table.endCol - table.startCol + 1);
      sh.getRange(DATA_START_ROW, table.startCol, rowsToClear, width).clearContent();
    }
  }
}

function groupFillsByAsset_(fills, assetTables) {
  const assetNames = assetTables.map(function(t) { return t.asset; });
  const fillsByAsset = {};
  
  for (var i = 0; i < assetNames.length; i++) fillsByAsset[assetNames[i]] = [];
  
  for (var j = 0; j < fills.length; j++) {
    const fill = fills[j];
    const fillCoin = fill.coin;
    var matched = false;
    for (var k = 0; k < assetNames.length; k++) {
      const assetName = assetNames[k];
      if (fillCoin === assetName || fillCoin.indexOf(assetName) !== -1 || assetName.indexOf(fillCoin) !== -1) {
        fillsByAsset[assetName].push(fill);
        matched = true; break;
      }
    }
    if (!matched) Logger.log('Warning: Could not match fill coin "' + fillCoin + '" to any asset table');
  }
  return fillsByAsset;
}

// --------------------------------- OPEN vs CLOSED - ANNOTATION ---------------------------------
function annotateAssetFills_(fills, fundingEvents) {
  if (!Array.isArray(fills) || fills.length === 0) return [];

  // Copy & sort ASC
  const arr = fills.slice().sort(function(a,b){ return a.time - b.time; });

  // FIFO lots for PnL attribution
  const lotsLong = [];   // {px, qty}
  const lotsShort = [];  // {px, qty}

  // Funding events (ASC): [{time, usdc}]
  const fEvts = (fundingEvents || []).slice().sort(function(a,b){ return a.time - b.time; });
  var fIdx = 0;          // pointer
  var fBucket = 0;       // accumulated funding since last allocation

  var pos = 0; // signed position accumulator (tokens)

  for (var i = 0; i < arr.length; i++) {
    const f = arr[i];
    const parts = parseDirParts_(f.dir);
    const price = Number(f.px);
    const qtyAbs = Math.abs(Number(f.sz));

    const startPosMaybe = toNum_(f.startPosition);
    const posBefore = (startPosMaybe != null ? startPosMaybe : pos);

    // Ingest funding up to this fill's time; only accrue when there is exposure open
    while (fIdx < fEvts.length && Number(fEvts[fIdx].time) <= Number(f.time)) {
      if (Math.abs(posBefore) > 0) fBucket += Number(fEvts[fIdx].usdc) || 0;
      fIdx++;
    }

    var signedDelta = 0;
    var realizedPnl = 0;
    var realizedBasis = 0; // sum(entryPx * matchedQty)

    if (parts.side === 'long' && parts.action === 'open') {
      lotsLong.push({ px: price, qty: qtyAbs });
      signedDelta = qtyAbs;
    } else if (parts.side === 'long' && parts.action === 'close') {
      var need = Math.min(qtyAbs, Math.abs(posBefore));
      // match FIFO from long lots
      while (need > 0 && lotsLong.length > 0) {
        var lot = lotsLong[0];
        var take = Math.min(need, lot.qty);
        realizedPnl += (price - lot.px) * take;
        realizedBasis += lot.px * take;
        lot.qty -= take; need -= take;
        if (lot.qty <= 0.0000001) lotsLong.shift();
      }
      signedDelta = -qtyAbs;
    } else if (parts.side === 'short' && parts.action === 'open') {
      lotsShort.push({ px: price, qty: qtyAbs });
      signedDelta = -qtyAbs;
    } else if (parts.side === 'short' && parts.action === 'close') {
      var needS = Math.min(qtyAbs, Math.abs(posBefore));
      while (needS > 0 && lotsShort.length > 0) {
        var lotS = lotsShort[0];
        var takeS = Math.min(needS, lotS.qty);
        realizedPnl += (lotS.px - price) * takeS; // short wins when price down
        realizedBasis += lotS.px * takeS;
        lotS.qty -= takeS; needS -= takeS;
        if (lotS.qty <= 0.0000001) lotsShort.shift();
      }
      signedDelta = qtyAbs; // closing short reduces negative exposure
    }

    const posAfter = posBefore + signedDelta;
    const absBefore = Math.abs(posBefore);
    const absAfter  = Math.abs(posAfter);

    f._posBefore = posBefore;
    f._posAfter  = posAfter;
    f._isOpening = absAfter > absBefore;
    f._isClosing = absAfter < absBefore;
    f._action    = (parts.actionLabel || 'Unknown');
    f._openAdded = f._isOpening ? (absAfter - absBefore) : 0;
    f._closeQty  = f._isClosing ? (absBefore - absAfter) : 0;
    f._side      = parts.side;

    // Funding allocation proportional to qty closed
    var fundingAlloc = 0;
    if (f._isClosing && absBefore > 0) {
      fundingAlloc = (fBucket * (f._closeQty / absBefore)) || 0;
      fBucket -= fundingAlloc;
    }
    f._fundingAlloc = fundingAlloc;

    // PnL (prefer API closedPnl if present for close fills)
    if (f._isClosing) {
      var apiClosed = toNum_(f.closedPnl);
      f._realizedPnl = (apiClosed != null ? apiClosed : realizedPnl);
      f._realizedBasis = realizedBasis;
      f._realizedPnlPct = (realizedBasis > 0 ? (f._realizedPnl / realizedBasis) * 100 : null);
    } else {
      f._realizedPnl = 0;
      f._realizedPnlPct = null;
    }

    // Total fees = trade fee + funding allocated
    var tradeFee = Number(f.fee);
    f._feeTotal = (Number.isFinite(tradeFee) ? tradeFee : 0) + (Number(fundingAlloc) || 0);

    pos = posAfter;
  }

  // Mark which openings still contribute to current open position (unchanged)
  const finalPos = pos;
  var needOpen = Math.abs(finalPos);
  const wantSide = (finalPos > 0 ? 'long' : (finalPos < 0 ? 'short' : null));
  for (var j = arr.length - 1; j >= 0 && needOpen > 0 && wantSide; j--) {
    const f2 = arr[j];
    f2._contributesOpen = false;
    f2._openRemainder   = 0;
    if (f2._isOpening && f2._side === wantSide) {
      const add2 = (f2._openAdded || Math.abs(Number(f2.sz)) || 0);
      const take2 = Math.min(needOpen, add2);
      if (take2 > 0) { f2._contributesOpen = true; f2._openRemainder = take2; needOpen -= take2; }
    }
  }
  for (var k = 0; k < arr.length; k++) if (arr[k]._contributesOpen !== true) arr[k]._contributesOpen = false;

  // Compute position-start notional (USD) from the first contributing opening fill
  var startIdx = -1;
  for (var s = 0; s < arr.length; s++) {
    var ff = arr[s];
    if (ff._contributesOpen && ff._isOpening) { startIdx = s; break; }
  }
  if (startIdx !== -1) {
    var f0 = arr[startIdx];
    var startTokens = (Number(f0._openRemainder) || Math.abs(Number(f0.sz)) || 0);
    var startPx = Number(f0.px);
    var startNotional = (Number.isFinite(startPx) && startTokens) ? startPx * startTokens : null;
    for (var t = 0; t < arr.length; t++) {
      if (arr[t]._contributesOpen) {
        arr[t]._positionStartNotional = startNotional;
        arr[t]._positionStartPx = startPx;
        arr[t]._positionStartTokens = startTokens;
      }
    }
  }

  return arr;
}

// Merge multiple fills with the exact same timestamp into a single trade (per asset)
function mergeFillsByTimestamp_(fills) {
  if (!Array.isArray(fills) || fills.length === 0) return [];
  var byTs = {};
  for (var i=0;i<fills.length;i++) {
    var f = fills[i];
    var k = String(f.time);
    if (!byTs[k]) byTs[k] = [];
    byTs[k].push(f);
  }
  var times = Object.keys(byTs).map(function(k){ return Number(k); }).sort(function(a,b){ return a - b; });
  var out = [];
  for (var t=0;t<times.length;t++) {
    var ts = times[t];
    var list = byTs[String(ts)];
    if (!list || list.length === 0) continue;

    var coin = String(list[0].coin || '');
    var sumAbsSz = 0, sumNotional = 0, sumWeightedPx = 0;
    var tradeFeeSum = 0, fundingSum = 0;
    var realizedPnlSum = 0, realizedBasisSum = 0;
    var openLong=0, openShort=0, closeLong=0, closeShort=0;
    var contributes = false;
    var startNotional = null;

    for (var j=0;j<list.length;j++) {
      var f = list[j];
      var qtyAbs = Math.abs(Number(f.sz)) || 0;
      var px = Number(f.px);
      if (Number.isFinite(px) && qtyAbs>0) { sumWeightedPx += px * qtyAbs; sumAbsSz += qtyAbs; sumNotional += px * qtyAbs; }
      tradeFeeSum += Number(f.fee) || 0;
      fundingSum  += Number(f._fundingAlloc) || 0;
      if (f._isClosing) { realizedPnlSum += Number(f._realizedPnl) || 0; realizedBasisSum += Number(f._realizedBasis) || 0; }
      if (f._isOpening && f._side==='long') openLong += Number(f._openAdded) || qtyAbs;
      if (f._isOpening && f._side==='short') openShort += Number(f._openAdded) || qtyAbs;
      if (f._isClosing && f._side==='long') closeLong += Number(f._closeQty) || qtyAbs;
      if (f._isClosing && f._side==='short') closeShort += Number(f._closeQty) || qtyAbs;
      contributes = contributes || !!f._contributesOpen;
      if (startNotional == null && Number.isFinite(f._positionStartNotional)) startNotional = Number(f._positionStartNotional);
    }

    var vwap = (sumAbsSz>0 ? sumWeightedPx / sumAbsSz : null);

    var openQty = openLong + openShort;
    var closeQty = closeLong + closeShort;
    var action = '';
    var side = '';
    if (openQty > closeQty) {
      action = 'open'; side = (openLong >= openShort ? 'long' : 'short');
    } else if (closeQty > openQty) {
      action = 'close'; side = (closeLong >= closeShort ? 'long' : 'short');
    } else {
      if (closeQty > 0) { action='close'; side=(closeLong>=closeShort?'long':'short'); }
      else if (openQty > 0) { action='open'; side=(openLong>=openShort?'long':'short'); }
      else { action='open'; side='long'; }
    }

    out.push({
      time: ts,
      coin: coin,
      dir: (action==='open' ? (side==='short' ? 'Open Short' : 'Open Long') : (side==='short' ? 'Close Short' : 'Close Long')),
      px: (vwap != null ? vwap : (list[0] && Number(list[0].px))),
      sz: sumAbsSz,
      fee: tradeFeeSum,
      _fundingAlloc: fundingSum,
      _isClosing: (closeQty > 0),
      _isOpening: (openQty > 0),
      _realizedPnl: realizedPnlSum,
      _realizedBasis: realizedBasisSum,
      _contributesOpen: contributes,
      _positionStartNotional: startNotional
    });
  }
  return out;
}

function parseDirParts_(dir) {
  const d = String(dir || '').toLowerCase();
  var action = null, side = null, label = null;

  if (d.indexOf('open') !== -1 && d.indexOf('long') !== -1) { action = 'open'; side = 'long'; label = 'Open Long'; }
  else if (d.indexOf('close') !== -1 && d.indexOf('long') !== -1) { action = 'close'; side = 'long'; label = 'Close Long'; }
  else if (d.indexOf('open') !== -1 && d.indexOf('short') !== -1) { action = 'open'; side = 'short'; label = 'Open Short'; }
  else if (d.indexOf('close') !== -1 && d.indexOf('short') !== -1) { action = 'close'; side = 'short'; label = 'Close Short'; }
  else if (d === 'buy') { action = 'open'; side = 'long'; label = 'Buy'; }
  else if (d === 'sell') { action = 'close'; side = 'long'; label = 'Sell'; }
  else { action = 'open'; side = 'long'; label = 'Long'; }
  return { action: action, side: side, actionLabel: label };
}

// --------------------------------- INDIVIDUAL TRADE ROW MAPPER ---------------------------------
function fillToTableRow_(fill, liqMap, leverageMap, collateralMap, openOrdMap) {
  const tradeTime = new Date(fill.time);
  const coin = fill.coin;
  const price = Number(fill.px);
  const size = Number(fill.sz);
  const notional = (Number.isFinite(price) && Number.isFinite(size)) ? price * size : null;

  const parts = parseDirParts_(fill.dir);
  const openClose = (parts.action === 'open' ? 'Open' : (parts.action === 'close' ? 'Close' : ''));
  const direction = (parts.side === 'short' ? 'Short' : 'Long');

  // Live position context
  const liqPriceLive = liqMap[coin];
  const curLeverage  = leverageMap[coin];
  const curCollateral= collateralMap[coin];

  // TP/SL only if this fill contributes to current open position
  var takeProfitLevel = '';
  var stopLossLevel = '';
  if (fill._contributesOpen) {
    const orders = openOrdMap[coin] || [];
    var bestTP = null, bestSL = null;
    for (var i = 0; i < orders.length; i++) {
      const o = orders[i];
      if (parts.side === 'long') {
        if (o.side === 'sell') {
          if (o.px > price) { if (bestTP == null || o.px < bestTP) bestTP = o.px; }
          if (o.px < price) { if (bestSL == null || o.px > bestSL) bestSL = o.px; }
        }
      } else { // short
        if (o.side === 'buy') {
          if (o.px < price) { if (bestTP == null || o.px > bestTP) bestTP = o.px; }
          if (o.px > price) { if (bestSL == null || o.px < bestSL) bestSL = o.px; }
        }
      }
    }
    takeProfitLevel = (bestTP != null ? bestTP : '');
    stopLossLevel   = (bestSL != null ? bestSL : '');
  }

  // Risk:Reward
  var riskReward = '';
  if (takeProfitLevel !== '' && stopLossLevel !== '' && Number.isFinite(price)) {
    var reward, risk;
    if (parts.side === 'long') { reward = Number(takeProfitLevel) - price; risk = price - Number(stopLossLevel); }
    else { reward = price - Number(takeProfitLevel); risk = Number(stopLossLevel) - price; }
    if (risk > 0 && Number.isFinite(reward)) riskReward = (reward / risk).toFixed(2);
  }

  // Collateral/Leverage/Liq only for currently-open contributing fills
  var displayCollateral = '';
  var displayLeverage = '';
  var liqPrice = '';
  if (fill._contributesOpen) {
    // Starting collateral = position-start notional / current leverage (approximation)
    var startNotional = Number(fill._positionStartNotional);
    if (Number.isFinite(startNotional) && Number.isFinite(curLeverage) && curLeverage > 0) {
      displayCollateral = startNotional / curLeverage;
    } else if (Number.isFinite(curCollateral)) {
      displayCollateral = curCollateral; // fallback
    }
    if (Number.isFinite(curLeverage)) displayLeverage = String(curLeverage) + 'x';
    if (Number.isFinite(liqPriceLive)) liqPrice = liqPriceLive;
  }

  // Fees column: trade fee + allocated funding trade fee + allocated funding
  var tradeFee = Number(fill.fee);
  var fundingAlloc = Number(fill._fundingAlloc) || 0;
  var feesTotal = (Number.isFinite(tradeFee) ? tradeFee : 0) - fundingAlloc; // net cost: fee minus funding credit

  // Realized PnL shown only on close fills (net of fees)
  var pnlUsd = '';
  var pnlPct = '';
  if (fill._isClosing) {
    var pnlUsdGross = Number(fill._realizedPnl) || 0;
    var basis = (fill._realizedBasis != null) ? Number(fill._realizedBasis) : null;
    pnlUsd = pnlUsdGross - feesTotal;
    pnlPct = (basis && basis > 0) ? (pnlUsd / basis) : '';
  }

  const row = [
    tradeTime,                                  // C: Date
    tradeTime,                                  // D: Time 
    'Hyperliquid',                              // E: Protocol
    openClose,                                  // F: Open/Close
    direction,                                  // G: Direction (Long/Short)
    Number.isFinite(price) ? price : '',        // H: Entry/Exit Price
    Number.isFinite(size) ? size : '',          // I: Position Size (tokens)
    Number.isFinite(notional) ? notional : '',  // J: Position Size (USD)
    displayCollateral,                          // K: Collateral (only if current open)
    displayLeverage,                            // L: Leverage (only if current open)
    feesTotal,                                  // M: Fees (trade + funding)
    takeProfitLevel,                            // N: Take Profit
    stopLossLevel,                              // O: Stop Loss
    liqPrice,                                   // P: Liquidation Price
    riskReward,                                 // Q: R:R (Risk:Reward)
    pnlUsd,                                     // R: PnL ($)
    (pnlPct !== '' ? pnlPct : '') // S: PnL (%) as decimal; formatted as % in sheet
  ];
  return row;
}

// --------------------------------- HTTP CORE (RETRY & BACKOFF) ---------------------------------
function hlPost_(payload) {
  const url = HL_API;
  const body = JSON.stringify(payload);
  const maxAttempts = 5;
  var wait = 200;

  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: body,
        muteHttpExceptions: true,
        headers: { 'Accept': 'application/json', 'Cache-Control': 'no-cache' }
      });
      const code = res.getResponseCode();
      const txt  = res.getContentText();
      if (code >= 300) throw new Error('HTTP ' + code + (payload && payload.type ? ' [' + payload.type + ']' : '') + ' ' + (txt ? txt.slice(0, 500) : ''));
      const json = JSON.parse(txt);
      if (json && json.error) throw new Error('API error ' + (payload && payload.type ? '[' + payload.type + ']' : '') + ' ' + JSON.stringify(json.error).slice(0, 500));
      return json;
    } catch (e) {
      if (attempt === maxAttempts) throw new Error('hlPost_ failed after ' + attempt + ' attempts for ' + (payload && payload.type ? payload.type : 'unknown') + ': ' + e.message);
      Utilities.sleep(wait);
      wait = Math.min(wait * 2, 2000);
    }
  }
  throw new Error('hlPost_ fell through.');
}

// --------------------------------- DATA FETCH ---------------------------------
function getPerpFillsAsc_(user, startMs, endMs) {
  const DAY = 24*60*60*1000, WINDOW = 30*DAY;
  var end = endMs, pages = 0, fetched = 0;
  const seen = new Set(), bag = [];

  while (end >= startMs && pages < 120 && fetched < 8000) {
    const start = Math.max(startMs, end - WINDOW);
    const page = hlPost_({ type:'userFillsByTime', user:user, startTime:start, endTime:end });

    if (Array.isArray(page) && page.length) {
      var minTime = Infinity;
      for (var i=0; i<page.length; i++) {
        var f = page[i];
        const coin = String((f && f.coin) || '');
        const isSpot = coin.indexOf('@') === 0 || coin.indexOf('/') !== -1;
        if (isSpot) continue;

        const tid = (f && f.tid != null) ? String(f.tid)
                  : (String(f.time) + '-' + f.oid + '-' + f.px + '-' + f.sz + '-' + f.dir);
        if (seen.has(tid)) continue;
        seen.add(tid);

        bag.push({
          time: Number(f.time),
          coin: coin,
          dir: (f && f.dir) || '',              // e.g., "Open Long", "Close Short"
          px: Number(f.px),
          sz: Number(f.sz),
          fee: Number(f.fee),
          startPosition: toNum_(f && f.startPosition),
          closedPnl: toNum_(f && f.closedPnl),
          oid: f.oid || '',
          tid: f.tid || tid
        });

        if (Number.isFinite(f.time) && f.time < minTime) minTime = f.time;
      }
      fetched += page.length;
      end = Number.isFinite(minTime) ? (minTime - 1) : (start - 1);
    } else {
      end = start - 1;
    }

    pages++;
    Utilities.sleep(120); // throttle per page
  }

  bag.sort(function(a,b){ return a.time - b.time; });
  return bag;
}

function getCurrentLiqPxMap_() {
  // Supports both clearinghouseState and batchClearinghouseStates shapes
  const chs = hlPost_({ type: 'clearinghouseState', user: HL_USER });

  // Normalized maps keyed by coin
  const liqMap = {};            // liquidationPx
  const leverageMap = {};       // leverage.value
  const collateralMap = {};     // marginUsed
  const entryPxMap = {};        // entryPx
  const posSzMap = {};          // szi (signed size)

  // Helper to ingest one position object
  function ingest(posObj) {
    if (!posObj) return;
    const coin = String(posObj.coin || '');
    if (!coin) return;
    const liq = Number(posObj.liquidationPx);
    const lev = (posObj.leverage && Number(posObj.leverage.value)) || null;
    const mgn = Number(posObj.marginUsed);
    const entry = Number(posObj.entryPx);
    const szi = Number(posObj.szi);

    if (Number.isFinite(liq)) liqMap[coin] = liq;
    if (lev != null && Number.isFinite(lev)) leverageMap[coin] = lev;
    if (Number.isFinite(mgn)) collateralMap[coin] = mgn;
    if (Number.isFinite(entry)) entryPxMap[coin] = entry;
    if (Number.isFinite(szi)) posSzMap[coin] = szi;
  }

  // API response may be { assetPositions: [{ position: {...}, type: 'oneWay' }, ...] }
  // Older clients sometimes exposed openPerpPositions directly; handle both.
  const assetPositions = (chs && chs.assetPositions) || (chs && chs.openPerpPositions) || [];
  for (var i = 0; i < assetPositions.length; i++) {
    const item = assetPositions[i];
    if (item && item.position) ingest(item.position); else ingest(item);
  }

  return { liquidation: liqMap, leverage: leverageMap, collateral: collateralMap, entryPx: entryPxMap, posSz: posSzMap };
}

function getOpenOrdersMap_() {
  var raw = [];
  try { raw = hlPost_({ type:'frontendOpenOrders', user:HL_USER }) || []; } catch(e) {}
  if (!Array.isArray(raw) || raw.length === 0) {
    try { raw = hlPost_({ type:'openOrders', user:HL_USER }) || []; } catch(e2) { raw = []; }
  }

  const norm = [];
  for (var i=0;i<raw.length;i++) {
    const o = raw[i] || {};
    const coin = String(o.coin || '');
    if (!coin) continue;

    // Normalize side: HL may use B (bid/buy) and A (ask/sell)
    var side;
    const sideRaw = String(o.side || '').toUpperCase();
    if (sideRaw === 'B' || sideRaw === 'BUY') side = 'buy';
    else if (sideRaw === 'A' || sideRaw === 'SELL') side = 'sell';
    else side = String(o.side || '').toLowerCase();
    if (!(side === 'buy' || side === 'sell')) continue;

    const limitPx   = toNum_(o.limitPx != null ? o.limitPx : o.px);
    const triggerPx = toNum_(o.triggerPx);
    const isTrigger = !!(o.isTrigger || o.triggerPx || o.trigger);
    const px = isTrigger ? (triggerPx != null ? triggerPx : limitPx) : (limitPx != null ? limitPx : triggerPx);
    if (!Number.isFinite(px)) continue;

    norm.push({ 
      coin: coin,
      side: side,
      px: px,
      type: (o.orderType ? String(o.orderType) : (isTrigger ? 'trigger' : 'limit')).toLowerCase(),
      reduceOnly: !!o.reduceOnly,
      isPositionTpsl: !!o.isPositionTpsl,
      sz: toNum_(o.sz != null ? o.sz : o.origSz) || 0
    });
  }

  const byCoin = {};
  for (var k=0;k<norm.length;k++) {
    const ord = norm[k];
    if (!byCoin[ord.coin]) byCoin[ord.coin] = [];
    byCoin[ord.coin].push(ord);
  }
  return byCoin;
}

// Funding events grouped by coin for fee attribution
function getFundingEventsByCoin_(user, startMs, endMs) {
  const events = [];
  const seen = new Set();

  var start = startMs; var pages = 0;
  while (start <= endMs && pages < 800) {
    var rows;
    try { rows = hlPost_({ type:'userFunding', user:user, startTime:start, endTime:endMs }); }
    catch(e){ start = start + 1; pages++; continue; }
    if (!Array.isArray(rows) || rows.length === 0) break;

    var maxTime = -Infinity;
    for (var i=0; i<rows.length; i++) {
      var r = rows[i] || {};
      const t = Number(r.time);
      const c = String(r && r.delta && r.delta.coin || '');
      const amt = Number(r && r.delta && r.delta.usdc);
      if (!c || !Number.isFinite(t) || !Number.isFinite(amt)) continue;

      const id = t + ':' + c + ':' + amt;
      if (seen.has(id)) continue;
      seen.add(id);

      events.push({ time: t, coin: c, usdc: amt });
      if (t > maxTime) maxTime = t;
    }
    pages++;
    if (!Number.isFinite(maxTime) || maxTime <= start) start = start + 1; else start = maxTime + 1;
    if (rows.length < 500) break;
  }

  // Group by coin
  const byCoin = {};
  for (var j=0; j<events.length; j++) {
    const e = events[j];
    if (!byCoin[e.coin]) byCoin[e.coin] = [];
    byCoin[e.coin].push({ time: e.time, usdc: e.usdc });
  }
  // sort each
  Object.keys(byCoin).forEach(function(k){ byCoin[k].sort(function(a,b){ return a.time - b.time; }); });
  return byCoin;
}

// --------------------------------- UTILITIES ---------------------------------
function getWindowMs_(days){
  const DAY=24*60*60*1000;
  const endMs=Date.now();
  const startMs=endMs - days*DAY - 60*1000;
  return { startMs:startMs, endMs:endMs };
}

function toNum_(x){ 
  var n=Number(x); 
  return Number.isFinite(n)?n:null; 
}

function colToA1_(col){
  var s = "";
  while (col > 0) { var r = (col - 1) % 26; s = String.fromCharCode(65 + r) + s; col = Math.floor((col - 1) / 26); }
  return s;
}

function formatRowDisplay_(row){
  // Date/time are always the first two entries
  const out = row.slice();
  if (out[0] instanceof Date) out[0] = Utilities.formatDate(out[0], TZ, 'yyyy-MM-dd');
  if (out[1] instanceof Date) out[1] = Utilities.formatDate(out[1], TZ, 'HH:mm:ss');
  // Round numeric convenience for cosmetic cleanliness
  function neat(n){ return (typeof n === 'number' && Number.isFinite(n)) ? Number((Math.abs(n) < 0.000001 ? 0 : n).toFixed(6)) : n; }
  for (var i=2;i<out.length;i++) out[i] = neat(out[i]);
  return out;
}

// --------------------------------- FEES & FUNDING (PRESERVED) ---------------------------------
function isSpotCoin_(coin) {
  return typeof coin === 'string' && (coin.indexOf('@') === 0 || coin.indexOf('/') !== -1);
}

function sumTradeFees_(user, startMs, endMs) {
  const DAY=24*60*60*1000, WINDOW=30*DAY;
  var end=endMs, total=0, fillsCount=0, perpsFees=0, spotFees=0, pages=0, fetched=0;
  const seen=new Set();
  while (end>=startMs && pages<200 && fetched<10000) {
    const start=Math.max(startMs, end-WINDOW);
    const fills=hlPost_({ type:'userFillsByTime', user:user, startTime:start, endTime:end });
    if (Array.isArray(fills) && fills.length) {
      var minTime=Infinity;
      for (var i=0;i<fills.length;i++) {
        var f=fills[i];
        const tid=(f && f.tid!=null)?String(f.tid):(String(f.time)+'-'+f.oid+'-'+f.px+'-'+f.sz+'-'+f.dir);
        if (seen.has(tid)) continue;
        seen.add(tid);

        const fee=Number(f.fee);
        if (!Number.isFinite(fee)) continue;

        const isSpot=isSpotCoin_(f && f.coin);
        if (isSpot) { spotFees+=fee; if (FEES_INCLUDE_SPOT) { total+=fee; fillsCount++; } }
        else { perpsFees+=fee; total+=fee; fillsCount++; }

        const t=Number(f.time);
        if (Number.isFinite(t) && t<minTime) minTime=t;
      }
      fetched+=fills.length;
      end = Number.isFinite(minTime) ? (minTime-1) : (start-1);
    } else {
      end = start-1;
    }
    pages++;
  }
  return { total:total, perpsFees:perpsFees, spotFees:spotFees, fillsCount:fillsCount, start: startMs, end: endMs };
}

function sumFundingNet_(user, startMs, endMs) {
  const startedAt=Date.now(), MAX_MS=25*1000;
  var start=startMs, net=0, paid=0, received=0, events=0, pages=0;
  const seen=new Set();
  while (start<=endMs && pages<800) {
    if (Date.now()-startedAt>MAX_MS) break;
    var rows;
    try { rows = hlPost_({ type:'userFunding', user:user, startTime:start, endTime:endMs }); }
    catch(e){ start=start+1; pages++; continue; }
    if (!Array.isArray(rows) || rows.length===0) break;

    var maxTime=-Infinity;
    for (var i=0;i<rows.length;i++) {
      var r=rows[i];
      const t=Number(r && r.time);
      const amt=Number(r && r.delta && r.delta.usdc);
      if (Number.isFinite(t) && t>maxTime) maxTime=t;

      const id=String(t)+':'+(r && r.delta && r.delta.coin)+':'+amt;
      if (seen.has(id)) continue;
      seen.add(id);

      if (Number.isFinite(amt)) {
        net += amt;
        if (amt<0) paid += -amt; else received += amt;
        events++;
      }
    }
    pages++;
    if (!Number.isFinite(maxTime) || maxTime<=start) start=start+1; else start=maxTime+1;
    if (rows.length<500) break;
  }
  return { net:net, paid:paid, received:received, events:events };
}

function UpdateTradingSummary_Hyperliquid() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Trading Summary');
  if (!sh) throw new Error('Sheet "Trading Summary" not found');

  sh.getRangeList(['G4','G15','G16','G22']).clearContent();

  const chs = hlPost_({ type: 'clearinghouseState', user: HL_USER });
  const accVal = toNum_(chs && chs.marginSummary && chs.marginSummary.accountValue);
  const withdrawable = toNum_(chs && chs.withdrawable);
  sh.getRange('G4').setValue(accVal != null ? accVal : 0);
  sh.getRange('G22').setValue(withdrawable != null ? withdrawable : 0);

  const win = getWindowMs_(LOOKBACK_DAYS);
  const fees = sumTradeFees_(HL_USER, win.startMs, win.endMs);
  sh.getRange('G15').setValue((fees.total || 0) * -1);

  const f = sumFundingNet_(HL_USER, win.startMs, win.endMs);
  sh.getRange('G16').setValue(f.net || 0);
}