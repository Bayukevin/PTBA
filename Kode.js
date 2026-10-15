const SHEETS = {
  MASTER: 'DB_Master_Material',
  HIST_SOH: 'Histori_SOH',
  HIST_TX: 'Histori_Transaksi',
  LOGS: 'Log_Upload',
  APL_SERVICE: 'DB_APL_Service',
  PO_REPORT: 'DB_PO'
};

const SPREADSHEET_ID = "1Y4lvjoF2nH_mzyByhsfdLDpiUtcs9yaoN_nGigc19cw";

function doGet(e) {
  checkAndSetupDatabase();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SWI PTBA - Inventory Hub')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function checkAndSetupDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const structures = [
    {
      name: SHEETS.MASTER,
      headers: [
        'No Urut', 'Produk', 'Komoditi', 'Material Utama', 
        'ITC Material 1', 'ITC Material 2', 'ITC Material 3', 'ITC Material 4', 'ITC Material 5',
        'Deskripsi Material', 'Indikator', 'Min Qty', 'Max Qty', 'Remark', 
        'Stok On Hand', 'Intransit', 'Pending Intransit', 'Status', 
        'Reorder Point', 'Qty Order', 'Suggest Max Qty', 'Suggest Min Qty'
      ]
    },
    {
      name: SHEETS.HIST_SOH,
      headers: ['Tanggal Upload', 'Timestamp', 'Material Utama', 'SOH', 'Intransit', 'Pending Intransit', 'Remark', 'Status']
    },
    {
      name: SHEETS.HIST_TX,
      headers: ['Tanggal Transaksi', 'No Transaksi', 'Material/PN_ITC', 'Qty', 'Harga Jual', 'HPP/COGS', 'Revenue', 'Gross Profit', 'Produk', 'Komoditi']
    },
    {
      name: SHEETS.LOGS,
      headers: ['Timestamp', 'Tipe Upload', 'Nama File', 'Jumlah Baris', 'Status']
    },
    {
      name: SHEETS.APL_SERVICE,
      headers: ['Material', 'Deskripsi', 'Qty', 'SOH', 'SOH Min', 'SOH Max', 'Indikator']
    },
    // Header Data Base PO
    {
      name: SHEETS.PO_REPORT,
      headers: [
        'Release status', 'PO history/release documentation', 'Purchasing Document', 'Item', 
        'Deletion Indicator', 'Release indicator', 'Document Date', 'Vendor/supplying', 
        'Plant', 'Material', 'Short Text', 'Order Quantity', 'Order Unit', 'Net price', 
        'Currency', 'Still to be delivered (qty)', 'Still to be delivered (value)', 
        'Storage Location', 'No. of Positions', 'Still to be invoiced (qty)', 
        'Still to be invoiced (val.)', 'Net Order Value', 'Requisitioner', 
        'Qty Released to Date', 'Plant 2', 'Purchasing Doc. Type'
      ]
    }
  ];

  structures.forEach(struct => {
    let sheet = ss.getSheetByName(struct.name);
    if (!sheet) {
      sheet = ss.insertSheet(struct.name);
    }
    if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() === '') {
      sheet.getRange(1, 1, 1, struct.headers.length).setValues([struct.headers])
           .setFontWeight('bold')
           .setBackground('#C8102E')
           .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
  });
}

/**
 * Jalankan fungsi ini dari editor Apps Script untuk mengisi header sheet DB_PO
 */
function setupDBPOHeaders() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const headers = [
    'Release status', 'PO history/release documentation', 'Purchasing Document', 'Item',
    'Deletion Indicator', 'Release indicator', 'Document Date', 'Vendor/supplying',
    'Plant', 'Material', 'Short Text', 'Order Quantity', 'Order Unit', 'Net price',
    'Currency', 'Still to be delivered (qty)', 'Still to be delivered (value)',
    'Storage Location', 'No. of Positions', 'Still to be invoiced (qty)',
    'Still to be invoiced (val.)', 'Net Order Value', 'Requisitioner',
    'Qty Released to Date', 'Plant 2', 'Purchasing Doc. Type'
  ];

  let sheet = ss.getSheetByName(SHEETS.PO_REPORT) || ss.getSheetByName('DB_PO');
  if (!sheet) {
    sheet = ss.insertSheet('DB_PO');
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight('bold')
       .setBackground('#C8102E')
       .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);

  Logger.log('Header DB_PO berhasil dipasang (' + headers.length + ' kolom).');
}

function calculateSOHStatus(soh, minQty, maxQty) {
  soh = Number(soh) || 0;
  minQty = Number(minQty) || 0;
  maxQty = Number(maxQty) || 0;

  if (soh >= maxQty && maxQty > 0) return 'A';
  if (soh >= minQty && minQty > 0) return 'B';
  if (soh > 0) return 'C';
  return 'D';
}

function getDashboardData(filterDate, filterProduct, filterCommodity) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName(SHEETS.MASTER);
  
  const masterValues = masterSheet.getDataRange().getValues();
  masterValues.shift();
  
  let filteredData = masterValues.filter(row => {
    let rowProd = String(row[1] || '').trim().toUpperCase();
    let rowComm = String(row[2] || '').trim().toUpperCase();
    
    let filterP = String(filterProduct || 'ALL').trim().toUpperCase();
    let filterC = String(filterCommodity || 'ALL').trim().toUpperCase();

    let matchProd = !filterProduct || filterP === 'ALL' || rowProd === filterP;
    let matchComm = !filterCommodity || filterC === 'ALL' || rowComm === filterC;
    return matchProd && matchComm;
  });

  let totalItems = filteredData.length;
  let reorderItemCount = 0;
  
  let remarkCount = {
    LIB: { A:0, B:0, C:0, D:0, total:0 },
    ADD: { A:0, B:0, C:0, D:0, total:0 },
    ALL: { A:0, B:0, C:0, D:0, total:0 }
  };

  filteredData.forEach(row => {
    let soh = Number(row[14]) || 0;
    let rop = Number(row[18]) || 0;
    
    let remark = String(row[13] || '').toUpperCase().trim();
    let rawStatus = String(row[17] || '').toUpperCase().trim();
    
    let status = ['A','B','C','D'].includes(rawStatus) 
      ? rawStatus 
      : calculateSOHStatus(soh, row[11], row[12]);

    if (soh <= rop && rop > 0) {
      reorderItemCount++;
    }

    if (['LIB', 'ADD'].includes(remark)) {
      if (remarkCount[remark] && remarkCount[remark][status] !== undefined) {
        remarkCount[remark][status]++;
        remarkCount[remark].total++;
      }
    }
    
    if (remarkCount.ALL[status] !== undefined) {
      remarkCount.ALL[status]++;
      remarkCount.ALL.total++;
    }
  });

  let readinessLIB = remarkCount.LIB.total > 0 
    ? (((remarkCount.LIB.A + remarkCount.LIB.B) / remarkCount.LIB.total) * 100).toFixed(2) 
    : "0.00";

  let readinessADD = remarkCount.ADD.total > 0 
    ? (((remarkCount.ADD.A + remarkCount.ADD.B) / remarkCount.ADD.total) * 100).toFixed(2) 
    : "0.00";

  return {
    kpi: {
      totalItems: totalItems,
      reorderItemCount: reorderItemCount,
      readinessLIB: readinessLIB,
      readinessADD: readinessADD
    },
    remarkBreakdown: remarkCount,
    filterOptions: getDropdownOptions()
  };
}

function getDropdownOptions() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName(SHEETS.MASTER);
  const data = masterSheet.getDataRange().getValues();
  data.shift();

  let products = new Set();
  let commodities = new Set();

  data.forEach(r => {
    if (r[1]) products.add(r[1]);
    if (r[2]) commodities.add(r[2]);
  });

  return {
    products: Array.from(products).sort(),
    commodities: Array.from(commodities).sort()
  };
}

function getHistoricalReadinessData(mode, startDate, endDate) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const histSheet = ss.getSheetByName(SHEETS.HIST_SOH);
  const data = histSheet.getDataRange().getValues();
  data.shift();

  let groupedData = {};

  data.forEach(row => {
    let rawDate = new Date(row[0]);
    if (isNaN(rawDate.getTime())) return;

    if (startDate && rawDate < new Date(startDate)) return;
    if (endDate && rawDate > new Date(endDate)) return;

    let year = rawDate.getFullYear();
    let month = String(rawDate.getMonth() + 1).padStart(2, '0');
    let day = String(rawDate.getDate()).padStart(2, '0');
    let periodKey = (mode === 'MONTHLY') ? `${year}-${month}` : `${year}-${month}-${day}`;

    let remark = String(row[6]).toUpperCase().trim();
    let status = String(row[7]).toUpperCase().trim();

    if (!groupedData[periodKey]) {
      groupedData[periodKey] = {
        LIB: { A:0, B:0, C:0, D:0, total:0, readiness:0 },
        ADD: { A:0, B:0, C:0, D:0, total:0, readiness:0 },
        ALL: { A:0, B:0, C:0, D:0, total:0, readiness:0 }
      };
    }

    if (['A', 'B', 'C', 'D'].includes(status)) {
      if (['LIB', 'ADD'].includes(remark)) {
        groupedData[periodKey][remark][status]++;
        groupedData[periodKey][remark].total++;
      }
      groupedData[periodKey].ALL[status]++;
      groupedData[periodKey].ALL.total++;
    }
  });

  let periods = Object.keys(groupedData).sort();
  let tableRows = [];

  periods.forEach(p => {
    ['LIB', 'ADD', 'ALL'].forEach(cat => {
      let d = groupedData[p][cat];
      d.readiness = d.total > 0 ? (((d.A + d.B) / d.total) * 100).toFixed(2) : 0;
    });

    tableRows.push({
      period: p,
      lib: groupedData[p].LIB,
      add: groupedData[p].ADD,
      all: groupedData[p].ALL
    });
  });

  return {
    periods: periods,
    groupedData: groupedData,
    tableRows: tableRows
  };
}

function getInventoryData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.MASTER);
  const values = sheet.getDataRange().getValues();
  values.shift();
  return values;
}

function runABCXYZAnalysis() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName(SHEETS.MASTER);
  const txSheet = ss.getSheetByName(SHEETS.HIST_TX);

  // Ambil nilai data dan rumus pada sheet Master
  const masterRange = masterSheet.getDataRange();
  const masterData = masterRange.getValues();
  const masterFormulas = masterRange.getFormulas(); // <-- Ambil rumus yang ada
  
  const masterHeaders = masterData.shift();
  masterFormulas.shift(); // Buang header rumus agar sejajar dengan masterData

  const txData = txSheet.getDataRange().getValues();
  txData.shift();

  let pnToMainMap = {};
  masterData.forEach((row) => {
    let mainPN = String(row[3]).trim();
    if (!mainPN) return;
    
    pnToMainMap[mainPN] = mainPN;
    for (let col = 4; col <= 8; col++) {
      if (row[col]) {
        pnToMainMap[String(row[col]).trim()] = mainPN;
      }
    }
  });

  let materialStats = {};
  txData.forEach(r => {
    let rawPN = String(r[2]).trim();
    let mainPN = pnToMainMap[rawPN] || rawPN;
    let qty = Number(r[3]) || 0;
    let revenue = Number(r[6]) || 0;
    let txDate = new Date(r[0]);

    if (!materialStats[mainPN]) {
      materialStats[mainPN] = { totalRevenue: 0, totalQty: 0, monthlyDemand: {} };
    }

    materialStats[mainPN].totalRevenue += revenue;
    materialStats[mainPN].totalQty += qty;

    if (!isNaN(txDate.getTime())) {
      let monthKey = txDate.getFullYear() + '-' + String(txDate.getMonth() + 1).padStart(2, '0');
      materialStats[mainPN].monthlyDemand[monthKey] = (materialStats[mainPN].monthlyDemand[monthKey] || 0) + qty;
    }
  });

  let sortedByRev = Object.keys(materialStats).sort((a,b) => materialStats[b].totalRevenue - materialStats[a].totalRevenue);
  let grandTotalRevenue = sortedByRev.reduce((acc, pn) => acc + materialStats[pn].totalRevenue, 0);

  let cumulativeRevenue = 0;
  sortedByRev.forEach(pn => {
    cumulativeRevenue += materialStats[pn].totalRevenue;
    let cumPercentage = grandTotalRevenue > 0 ? (cumulativeRevenue / grandTotalRevenue) * 100 : 100;
    
    if (cumPercentage <= 80) materialStats[pn].abc = 'A';
    else if (cumPercentage <= 95) materialStats[pn].abc = 'B';
    else materialStats[pn].abc = 'C';
  });

  Object.keys(materialStats).forEach(pn => {
    let demands = Object.values(materialStats[pn].monthlyDemand);
    if (demands.length < 2) {
      materialStats[pn].xyz = 'Z';
      return;
    }

    let mean = demands.reduce((a, b) => a + b, 0) / 12;
    let variance = demands.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / 12;
    let stdDev = Math.sqrt(variance);
    let cv = mean > 0 ? stdDev / mean : 999;

    if (cv <= 0.5) materialStats[pn].xyz = 'X';
    else if (cv <= 1.0) materialStats[pn].xyz = 'Y';
    else materialStats[pn].xyz = 'Z';
  });

  let updatedRows = masterData.map((row, index) => {
    let mainPN = String(row[3]).trim();
    let stats = materialStats[mainPN];

    let finalIndicator = 'D. Non Moving';
    let suggestMin = Number(row[11]) || 0;
    let suggestMax = Number(row[12]) || 0;
    let rop = Number(row[18]) || 0;
    let qtyOrder = 0;

    if (stats) {
      let combineClass = `${stats.abc}${stats.xyz}`;
      
      if (['AX', 'AY', 'BX'].includes(combineClass)) finalIndicator = 'A. Fast Moving';
      else if (['AZ', 'BY', 'CX'].includes(combineClass)) finalIndicator = 'B. Medium Moving';
      else if (['BZ', 'CY', 'CZ'].includes(combineClass)) finalIndicator = 'C. Slow Moving';

      let avgDailyDemand = stats.totalQty / 365;
      let leadTimeDays = 14;
      let safetyStockMultiplier = stats.xyz === 'X' ? 1.2 : stats.xyz === 'Y' ? 1.5 : 2.0;

      let safetyStock = Math.ceil(avgDailyDemand * leadTimeDays * (safetyStockMultiplier - 1));
      suggestMin = Math.ceil((avgDailyDemand * leadTimeDays) + safetyStock);
      suggestMax = Math.ceil(suggestMin * 2.5);
      
      rop = suggestMin;
      let currentSOH = Number(row[14]) || 0;
      qtyOrder = currentSOH <= rop ? Math.max(0, suggestMax - currentSOH) : 0;
    }

    let currentSOH = Number(row[14]) || 0;
    let status = calculateSOHStatus(currentSOH, suggestMin, suggestMax);

    row[10] = finalIndicator; 
    row[11] = suggestMin;      
    row[12] = suggestMax;      
    row[17] = status;          
    row[18] = rop;             
    row[19] = qtyOrder;        
    row[20] = suggestMax;      
    row[21] = suggestMin;      

    // Restore rumus untuk kolom O (14), P (15), dan Q (16) jika ada
    [14, 15, 16].forEach(colIdx => {
      if (masterFormulas[index][colIdx]) {
        row[colIdx] = masterFormulas[index][colIdx];
      }
    });

    return row;
  });

  if (updatedRows.length > 0) {
    masterSheet.getRange(2, 1, updatedRows.length, masterHeaders.length).setValues(updatedRows);
  }

  return { success: true, message: 'Analisis ABC-XYZ & Recalculate Min/Max berhasil diperbarui!' };
}

function getROPReportData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("DB_Master_Material"); // Sesuaikan nama sheet Anda
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    // Mengambil seluruh data (asumsi Baris 1 adalah Header)
    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    var result = [];
    //Logger.log("disini");
    data.forEach(function(row, index) {
      // Ambil variabel penting dengan fallback nilai 0 / String Kosong jika null/blank
      var materialUtama = row[3] || "";    // Kolom D (Material Utama)
      var deskripsi     = row[9] || "";    // Kolom J (Deskripsi)
      var produk        = row[1] || "";    // Kolom B (Produk)
      var komoditi      = row[2] || "";    // Kolom C (Komoditi)
      var soh           = Number(row[14]) || 0; // Kolom O (SOH)
      var rop           = Number(row[18]) || 0; // Kolom S (ROP)
      var maxQty        = Number(row[12]) || 0; // Kolom M (Max Qty)
      var pendingPO     = Number(row[16]) || 0; // Kolom Q (Pending)
      var qtyRecommend  = Number(row[19]) || 0; // Kolom T (Qty Order)
      var statusROP     = (soh <= rop) ? "REORDER NOW" : "SAFE";

      // Memasukkan array yang rapi dan aman ke hasil
      result.push([
        index + 1,        // [0] No
        materialUtama,    // [1] Material Utama
        deskripsi,        // [2] Deskripsi
        produk,           // [3] Produk
        komoditi,         // [4] Komoditi
        soh,              // [5] SOH
        rop,              // [6] ROP
        maxQty,           // [7] Max Qty
        pendingPO,        // [8] Pending PO
        qtyRecommend,     // [9] Qty Recommendation
        statusROP         // [10] Status ROP
      ]);
    });
    Logger.log (result);
    return result;
  } catch (e) {
    //Logger.log("gagal");
    throw new Error("Gagal mengambil data ROP: " + e.message);
  }
}

function getMinMaxComparisonData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName(SHEETS.MASTER);
  
  // Mengambil data beserta nilai hasil evaluasi rumus
  const data = masterSheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Antisipasi jika sheet kosong / hanya ada header
  
  // Hapus baris header
  data.shift();

  // Helper untuk konversi aman angka (termasuk hasil dari rumus)
  const parseNum = (val) => {
    if (val === "" || val === null || val === undefined) return 0;
    const num = Number(val);
    return isNaN(num) ? 0 : num;
  };

  let comparisonList = [];

  data.forEach(r => {
    // Ambil nilai (indeks array 0-based: r[11] = Kolom L, r[12] = Kolom M, dst.)
    let minExisting = parseNum(r[11]);
    let maxExisting = parseNum(r[12]);
    let minSuggest  = parseNum(r[21]);
    let maxSuggest  = parseNum(r[20]);

    let deviasiMin = minSuggest - minExisting;
    let deviasiMax = maxSuggest - maxExisting;

    // Logika Evaluasi Rekomendasi
    let recommendation = 'OPTIMAL';
    if (deviasiMin > 0 || deviasiMax > 0) {
      recommendation = 'UNDERSTOCKED_PARAM';
    } else if (deviasiMin < 0 || deviasiMax < 0) {
      recommendation = 'OVERSTOCKED_PARAM';
    }

    comparisonList.push({
      produk: r[1],
      komoditi: r[2],
      materialUtama: r[3],
      deskripsi: r[9],
      indikator: r[10],
      minExisting: minExisting,
      maxExisting: maxExisting,
      minSuggest: minSuggest,
      maxSuggest: maxSuggest,
      deviasiMin: deviasiMin,
      deviasiMax: deviasiMax,
      recommendation: recommendation
    });
  });

  return comparisonList;
}

function applyMinMaxSuggestions() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName(SHEETS.MASTER);
  const masterData = masterSheet.getDataRange().getValues();
  
  // Hapus header
  masterData.shift();

  let updatedCount = 0;

  masterData.forEach(row => {
    let minSuggest = Number(row[21]) || 0;
    let maxSuggest = Number(row[20]) || 0;

    if (minSuggest > 0 || maxSuggest > 0) {
      row[11] = minSuggest; // Kolom L (Indeks 11 / Kolom 12) -> Min Existing
      row[12] = maxSuggest; // Kolom M (Indeks 12 / Kolom 13) -> Max Existing
      
      let soh = Number(row[14]) || 0; // Membaca data dari Kolom O (Indeks 14)
      row[17] = calculateSOHStatus(soh, minSuggest, maxSuggest); // Kolom R (Indeks 17 / Kolom 18) -> Status SOH
      updatedCount++;
    }
  });

  if (masterData.length > 0) {
    // 1. Tulis HANYA Kolom L & M (Kolom ke-12 dan 13)
    const minMaxData = masterData.map(row => [row[11], row[12]]);
    masterSheet.getRange(2, 12, minMaxData.length, 2).setValues(minMaxData);

    // 2. Tulis HANYA Kolom R (Kolom ke-18)
    // Kolom O, P, Q (Kolom ke-15, 16, 17) DILOMPATI dan TIDAK TERSENTUH
    const statusData = masterData.map(row => [row[17]]);
    masterSheet.getRange(2, 18, statusData.length, 1).setValues(statusData);
  }

  return { 
    success: true, 
    count: updatedCount, 
    message: `Berhasil memperbarui ${updatedCount} parameter Min/Max Existing!` 
  };
}

function getTransactionData(startDate, endDate, product, commodity) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const txSheet = ss.getSheetByName(SHEETS.HIST_TX);
  const txData = txSheet.getDataRange().getValues();
  txData.shift();

  let filtered = txData.filter(r => {
    let tDate = new Date(r[0]);
    let matchStart = !startDate || tDate >= new Date(startDate);
    let matchEnd = !endDate || tDate <= new Date(endDate);
    let matchProd = !product || product === 'ALL' || r[8] === product;
    let matchComm = !commodity || commodity === 'ALL' || r[9] === commodity;
    return matchStart && matchEnd && matchProd && matchComm;
  });

  let totalRevenue = 0;
  let totalCOGS = 0;
  let totalGP = 0;

  let monthlyTrend = {};
  let productRev = {};

  filtered.forEach(r => {
    let rev = Number(r[6]) || 0;
    let gp = Number(r[7]) || 0;
    let cogs = Number(r[5]) || 0;
    let prod = r[8] || 'Uncategorized';
    let d = new Date(r[0]);

    totalRevenue += rev;
    totalCOGS += cogs;
    totalGP += gp;

    let monthStr = isNaN(d.getTime()) ? 'Unknown' : `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    if (!monthlyTrend[monthStr]) monthlyTrend[monthStr] = { rev: 0, gp: 0 };
    monthlyTrend[monthStr].rev += rev;
    monthlyTrend[monthStr].gp += gp;

    productRev[prod] = (productRev[prod] || 0) + rev;
  });

  let gpMargin = totalRevenue > 0 ? (totalGP / totalRevenue) * 100 : 0;

  return {
    kpi: {
      totalRevenue: totalRevenue,
      totalGP: totalGP,
      gpMargin: gpMargin.toFixed(2)
    },
    monthlyTrend: monthlyTrend,
    productContribution: productRev
  };
}



function processTransactionUpload(rowsData, fileName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const txSheet = ss.getSheetByName(SHEETS.HIST_TX);
  const logSheet = ss.getSheetByName(SHEETS.LOGS);
  const timestamp = new Date();

  let formattedRows = rowsData.map(r => {
    let qty = Number(r.qty || r['Qty']) || 0;
    let hargaJual = Number(r.hargaJual || r['Harga Jual']) || 0;
    let rev = Number(r.revenue || r['Revenue']) || (qty * hargaJual);
    let cogs = Number(r.hpp || r['HPP/COGS']) || 0;
    let gp = Number(r.grossProfit || r['Gross Profit']) || (rev - (cogs * qty));

    return [
      r.tanggalTransaksi || r['Tanggal Transaksi'],
      r.noTransaksi || r['No Transaksi'],
      r.materialPN || r['Material/PN_ITC'],
      qty,
      hargaJual,
      cogs,
      rev,
      gp,
      r.produk || r['Produk'],
      r.komoditi || r['Komoditi']
    ];
  });

  if (formattedRows.length > 0) {
    txSheet.getRange(txSheet.getLastRow() + 1, 1, formattedRows.length, formattedRows[0].length).setValues(formattedRows);
  }

  logSheet.appendRow([timestamp, 'Transaction Upload', fileName, rowsData.length, 'SUCCESS']);

  return { success: true, count: rowsData.length };
}

function getUploadLogs() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = ss.getSheetByName(SHEETS.LOGS);
  const data = logSheet.getDataRange().getValues();
  data.shift(); 
  return data.reverse(); 
}

function getAPLServiceData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEETS.APL_SERVICE);
    
    // Jika sheet belum ada, setup otomatis
    if (!sheet) {
      checkAndSetupDatabase();
      sheet = ss.getSheetByName(SHEETS.APL_SERVICE);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    
    return data.map((row, index) => ({
      no: index + 1,
      material: String(row[0] || '').trim(),
      deskripsi: String(row[1] || '').trim(),
      qty: Number(row[2]) || 0,
      soh: Number(row[3]) || 0,
      sohMin: Number(row[4]) || 0,
      sohMax: Number(row[5]) || 0,
      indikator: String(row[6] || 'Unassigned').trim()
    }));
  } catch (e) {
    throw new Error("Gagal mengambil data APL Service: " + e.message);
  }
}

function getPOReportData() {
  try {
    const ss = SpreadsheetApp.openById("1Y4lvjoF2nH_mzyByhsfdLDpiUtcs9yaoN_nGigc19cw");
    let sheet = ss.getSheetByName(SHEETS.PO_REPORT) || ss.getSheetByName('DB_PO');
    if (sheet && sheet.getLastRow() <= 1) {
      const alt = ss.getSheetByName('DB_PO') || ss.getSheetByName('DB_PO_Report');
      if (alt && alt.getLastRow() > 1) sheet = alt;
    }

    if (!sheet) {
      Logger.log("ERROR: Sheet PO tidak ditemukan!");
      return {
        items: [],
        kpi: { totalPO: 0, openPO: 0, closedPO: 0, totalVal: 0, openVal: 0 }
      };
    }

    Logger.log("Sheet dibaca: " + sheet.getName() + " | LastRow: " + sheet.getLastRow());

    const data = sheet.getDataRange().getValues();
    Logger.log("Data total baris: " + data.length);

    if (data.length <= 1) {
      Logger.log("Data kosong atau hanya header.");
      return {
        items: [],
        kpi: { totalPO: 0, openPO: 0, closedPO: 0, totalVal: 0, openVal: 0 }
      };
    }

    const rows = data.slice(1);

    let kpi = {
      totalPO: rows.length,
      openPO: 0,
      closedPO: 0,
      totalVal: 0,
      openVal: 0
    };

    const items = rows.map((row, index) => {
      // Menangani Indeks Kolom Sesuai Struktur Sheet PO
      const stillDelivQty = parseFloat(row[15]) || 0; // Still to be delivered (qty)
      const stillDelivVal = parseFloat(row[16]) || 0; // Still to be delivered (value)
      const netOrderVal   = parseFloat(row[21]) || 0; // Net Order Value
      const relInd        = String(row[5] || '').trim(); // Release indicator

      const poStatus = (stillDelivQty > 0) ? 'OPEN' : 'CLOSED';
      if (poStatus === 'OPEN') {
        kpi.openPO++;
        kpi.openVal += stillDelivVal;
      } else {
        kpi.closedPO++;
      }

      kpi.totalVal += netOrderVal;

      let releaseStatusText = 'Hold';
      if (relInd === '1') releaseStatusText = 'Approved';
      if (relInd === '0') releaseStatusText = 'Belum Approved';

      return {
        no: index + 1,
        poStatus: poStatus,
        releaseStatus: String(row[0] || ''),
        poHistoryDoc: String(row[1] || ''),
        purchasingDoc: String(row[2] || ''),
        item: String(row[3] || ''),
        deletionIndicator: String(row[4] || ''),
        releaseIndicator: relInd,
        releaseStatusText: releaseStatusText,
        docDate: (row[6] && !isNaN(new Date(row[6]).getTime())) ? Utilities.formatDate(new Date(row[6]), Session.getScriptTimeZone(), "yyyy-MM-dd") : String(row[6] || ''),
        vendor: String(row[7] || ''),
        plant: String(row[8] || ''),
        material: String(row[9] || ''),
        shortText: String(row[10] || ''),
        orderQty: parseFloat(row[11]) || 0,
        orderUnit: String(row[12] || ''),
        netPrice: parseFloat(row[13]) || 0,
        currency: String(row[14] || ''),
        stillDelivQty: stillDelivQty,
        stillDelivVal: stillDelivVal,
        storageLoc: String(row[17] || ''),
        noOfPositions: String(row[18] || ''),
        stillInvoiceQty: parseFloat(row[19]) || 0,
        stillInvoiceVal: parseFloat(row[20]) || 0,
        netOrderVal: netOrderVal,
        requisitioner: String(row[22] || ''),
        qtyReleasedToDate: String(row[23] || ''),
        plant2: String(row[24] || ''),
        purchasingDocType: String(row[25] || '')
      };
    });

    Logger.log("Sukses mapping: " + items.length + " item. KPI: " + JSON.stringify(kpi));

    return {
      items: items,
      kpi: kpi
    };

  } catch (e) {
    throw new Error("Gagal mengambil data PO Report: " + e.message);
  }
}

function processPOReportUpload(rowData, fileName) {
  try {
    const ss = SpreadsheetApp.openById("1Y4lvjoF2nH_mzyByhsfdLDpiUtcs9yaoN_nGigc19cw");
    let sheet = ss.getSheetByName(SHEETS.PO_REPORT);

    if (!sheet) {
      checkAndSetupDatabase();
      sheet = ss.getSheetByName(SHEETS.PO_REPORT);
    }

    if (!rowData || rowData.length === 0) {
      throw new Error("File Excel kosong atau format tidak sesuai.");
    }

    // Hitung jumlah kolom maksimum dari data Excel yang diunggah
    const maxCols = Math.max(...rowData.map(row => row.length));

    // Normalisasi setiap baris agar memiliki jumlah kolom yang seragam
    const formattedRows = rowData.map(row => {
      let newRow = new Array(maxCols).fill('');
      for (let i = 0; i < row.length; i++) {
        newRow[i] = (row[i] !== null && row[i] !== undefined) ? row[i] : '';
      }
      return newRow;
    });

    // Hapus data lama di sheet (kecuali Baris 1 / Header)
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, Math.max(lastCol, maxCols)).clearContent();
    }

    // Masukkan data baru hasil upload Excel
    if (formattedRows.length > 0) {
      sheet.getRange(2, 1, formattedRows.length, maxCols).setValues(formattedRows);
    }

    // Catat Riwayat ke Log Upload
    let logSheet = ss.getSheetByName(SHEETS.LOGS);
    if (logSheet) {
      logSheet.appendRow([
        new Date(),
        'PO_REPORT_UPLOAD',
        fileName || 'PO_Data.xlsx',
        formattedRows.length,
        'SUCCESS'
      ]);
    }

    return {
      status: 'success',
      message: `Berhasil mengunggah ${formattedRows.length} baris data PO Report.`,
      totalRows: formattedRows.length
    };

  } catch (e) {
    throw new Error("Gagal memproses upload data PO: " + e.message);
  }
}

function getProkerData() {
  try {
    const ss = SpreadsheetApp.openById("1Y4lvjoF2nH_mzyByhsfdLDpiUtcs9yaoN_nGigc19cw");
    let sheet = ss.getSheetByName("DB_Proker");
    
    // Jika nama sheet default di Spreadsheet adalah 'Proker', gunakan fallback
    if (!sheet) {
      sheet = ss.getSheetByName('Proker');
    }

    if (!sheet) {
      return {
        items: [],
        kpi: { totalItems: 0, openCount: 0, closeCount: 0, pctStatusOpen: 0, pctStatusClose: 0, totalProker: 0, pctProkerClose: 0 },
        filterOptions: { prokerList: [], materialList: [], statusList: [], unitList: [] },
        prokerCloseSummary: []
      };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return {
        items: [],
        kpi: { totalItems: 0, openCount: 0, closeCount: 0, pctStatusOpen: 0, pctStatusClose: 0, totalProker: 0, pctProkerClose: 0 },
        filterOptions: { prokerList: [], materialList: [], statusList: [], unitList: [] },
        prokerCloseSummary: []
      };
    }

    const rows = data.slice(1);

    let openCount = 0;
    let closeCount = 0;

    let prokerSet = new Set();
    let materialSet = new Set();
    let statusSet = new Set();
    let unitSet = new Set();

    let prokerStats = {}; // { [prokerName]: { total: 0, close: 0 } }

    const items = rows.map((row, index) => {
      const status = String(row[0] || '').trim().toUpperCase();
      const proker = String(row[1] || '').trim();
      const modelUnit = String(row[2] || '').trim();
      const kodeUnit = String(row[3] || '').trim();
      const material = String(row[4] || '').trim();
      const desc = String(row[5] || '').trim();
      const qty = parseFloat(row[6]) || 0;
      const planInstall = row[7] ? String(row[7]).trim() : '';
      const tglSupplay = row[8] ? String(row[8]).trim() : '';
      const qtySupplay = parseFloat(row[9]) || 0;
      const openQty = parseFloat(row[10]) || 0;
      const remark = String(row[11] || '').trim();

      if (status === 'CLOSE') closeCount++;
      else if (status === 'OPEN') openCount++;

      if (proker) prokerSet.add(proker);
      if (material) materialSet.add(material);
      if (status) statusSet.add(status);
      if (kodeUnit) unitSet.add(kodeUnit);

      // Hitung per Proker
      if (proker) {
        if (!prokerStats[proker]) {
          prokerStats[proker] = { total: 0, close: 0 };
        }
        prokerStats[proker].total++;
        if (status === 'CLOSE') {
          prokerStats[proker].close++;
        }
      }

      return {
        no: index + 1,
        status: status,
        proker: proker,
        modelUnit: modelUnit,
        kodeUnit: kodeUnit,
        material: material,
        desc: desc,
        qty: qty,
        planInstall: planInstall,
        tglSupplay: tglSupplay,
        qtySupplay: qtySupplay,
        openQty: openQty,
        remark: remark
      };
    });

    const totalItems = items.length;
    const pctStatusClose = totalItems > 0 ? parseFloat(((closeCount / totalItems) * 100).toFixed(2)) : 0;
    const pctStatusOpen = totalItems > 0 ? parseFloat(((openCount / totalItems) * 100).toFixed(2)) : 0;

    // Ringkasan Persentase CLOSE per Proker (Dashboard 2)
    let prokerCloseSummary = [];
    let totalProkerCount = Object.keys(prokerStats).length;
    let totalProkerClosePct = 0;

    Object.keys(prokerStats).forEach(pName => {
      let stat = prokerStats[pName];
      let pct = stat.total > 0 ? parseFloat(((stat.close / stat.total) * 100).toFixed(2)) : 0;
      prokerCloseSummary.push({
        prokerName: pName,
        total: stat.total,
        close: stat.close,
        pctClose: pct
      });
      totalProkerClosePct += pct;
    });

    const avgProkerClosePct = totalProkerCount > 0 ? parseFloat((totalProkerClosePct / totalProkerCount).toFixed(2)) : 0;

    return {
      items: items,
      kpi: {
        totalItems: totalItems,
        openCount: openCount,
        closeCount: closeCount,
        pctStatusOpen: pctStatusOpen,
        pctStatusClose: pctStatusClose,
        totalProker: totalProkerCount,
        pctProkerClose: avgProkerClosePct
      },
      prokerCloseSummary: prokerCloseSummary,
      filterOptions: {
        prokerList: Array.from(prokerSet).sort(),
        materialList: Array.from(materialSet).sort(),
        statusList: Array.from(statusSet).sort(),
        unitList: Array.from(unitSet).sort()
      }
    };
  } catch (e) {
    throw new Error("Gagal mengambil data Proker: " + e.message);
  }
}

function getProkerData() {
  try {
    const ss = SpreadsheetApp.openById("1Y4lvjoF2nH_mzyByhsfdLDpiUtcs9yaoN_nGigc19cw");
    let sheet = ss.getSheetByName("DB_Proker");
    
    // Jika nama sheet default di Spreadsheet adalah 'Proker', gunakan fallback
    if (!sheet) {
      sheet = ss.getSheetByName('Proker');
    }

    if (!sheet) {
      return {
        items: [],
        kpi: { totalItems: 0, openCount: 0, closeCount: 0, pctStatusOpen: 0, pctStatusClose: 0, totalProker: 0, pctProkerClose: 0 },
        filterOptions: { prokerList: [], materialList: [], statusList: [], unitList: [] },
        prokerCloseSummary: []
      };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return {
        items: [],
        kpi: { totalItems: 0, openCount: 0, closeCount: 0, pctStatusOpen: 0, pctStatusClose: 0, totalProker: 0, pctProkerClose: 0 },
        filterOptions: { prokerList: [], materialList: [], statusList: [], unitList: [] },
        prokerCloseSummary: []
      };
    }

    const rows = data.slice(1);

    let openCount = 0;
    let closeCount = 0;

    let prokerSet = new Set();
    let materialSet = new Set();
    let statusSet = new Set();
    let unitSet = new Set();

    let prokerStats = {}; // { [prokerName]: { total: 0, close: 0 } }

    const items = rows.map((row, index) => {
      const status = String(row[0] || '').trim().toUpperCase();
      const proker = String(row[1] || '').trim();
      const modelUnit = String(row[2] || '').trim();
      const kodeUnit = String(row[3] || '').trim();
      const material = String(row[4] || '').trim();
      const desc = String(row[5] || '').trim();
      const qty = parseFloat(row[6]) || 0;
      const planInstall = row[7] ? String(row[7]).trim() : '';
      const tglSupplay = row[8] ? String(row[8]).trim() : '';
      const qtySupplay = parseFloat(row[9]) || 0;
      const openQty = parseFloat(row[10]) || 0;
      const remark = String(row[11] || '').trim();

      if (status === 'CLOSE') closeCount++;
      else if (status === 'OPEN') openCount++;

      if (proker) prokerSet.add(proker);
      if (material) materialSet.add(material);
      if (status) statusSet.add(status);
      if (kodeUnit) unitSet.add(kodeUnit);

      // Hitung per Proker
      if (proker) {
        if (!prokerStats[proker]) {
          prokerStats[proker] = { total: 0, close: 0 };
        }
        prokerStats[proker].total++;
        if (status === 'CLOSE') {
          prokerStats[proker].close++;
        }
      }

      return {
        no: index + 1,
        status: status,
        proker: proker,
        modelUnit: modelUnit,
        kodeUnit: kodeUnit,
        material: material,
        desc: desc,
        qty: qty,
        planInstall: planInstall,
        tglSupplay: tglSupplay,
        qtySupplay: qtySupplay,
        openQty: openQty,
        remark: remark
      };
    });

    const totalItems = items.length;
    const pctStatusClose = totalItems > 0 ? parseFloat(((closeCount / totalItems) * 100).toFixed(2)) : 0;
    const pctStatusOpen = totalItems > 0 ? parseFloat(((openCount / totalItems) * 100).toFixed(2)) : 0;

    // Ringkasan Persentase CLOSE per Proker (Dashboard 2)
    let prokerCloseSummary = [];
    let totalProkerCount = Object.keys(prokerStats).length;
    let totalProkerClosePct = 0;

    Object.keys(prokerStats).forEach(pName => {
      let stat = prokerStats[pName];
      let pct = stat.total > 0 ? parseFloat(((stat.close / stat.total) * 100).toFixed(2)) : 0;
      prokerCloseSummary.push({
        prokerName: pName,
        total: stat.total,
        close: stat.close,
        pctClose: pct
      });
      totalProkerClosePct += pct;
    });

    const avgProkerClosePct = totalProkerCount > 0 ? parseFloat((totalProkerClosePct / totalProkerCount).toFixed(2)) : 0;

    return {
      items: items,
      kpi: {
        totalItems: totalItems,
        openCount: openCount,
        closeCount: closeCount,
        pctStatusOpen: pctStatusOpen,
        pctStatusClose: pctStatusClose,
        totalProker: totalProkerCount,
        pctProkerClose: avgProkerClosePct
      },
      prokerCloseSummary: prokerCloseSummary,
      filterOptions: {
        prokerList: Array.from(prokerSet).sort(),
        materialList: Array.from(materialSet).sort(),
        statusList: Array.from(statusSet).sort(),
        unitList: Array.from(unitSet).sort()
      }
    };
  } catch (e) {
    throw new Error("Gagal mengambil data Proker: " + e.message);
  }
}

// Scope Global / Paling Atas File
const EXCEL_TARGET_SHEETS = ["DB_PO", "DB_No_Ibd", "DO to BA", "DB_Intrs"];

/**
 * Memproses Upload File Excel dengan Multi-Sheet
 */
function processExcelUpload(formObject) {
  try {
    const blob = formObject.excelFile; 
    if (!blob || blob.getBytes().length === 0) {
      throw new Error("Pilih file Excel terlebih dahulu!");
    }

    const tempFile = DriveApp.createFile(blob);
    const tempFileId = tempFile.getId();

    const resource = {
      title: '[TEMP_CONVERT] ' + blob.getName(),
      mimeType: MimeType.GOOGLE_SHEETS
    };
    
    const convertedFile = DriveApp.getFileById(
      Drive.Files.insert(resource, blob, { convert: true }).id
    );
    const uploadedSs = SpreadsheetApp.openById(convertedFile.getId());
    const targetSs = SpreadsheetApp.openById("1Y4lvjoF2nH_mzyByhsfdLDpiUtcs9yaoN_nGigc19cw");

    let logs = [];
    let updatedSheets = [];

    EXCEL_TARGET_SHEETS.forEach(sheetName => {
      const sourceSheet = uploadedSs.getSheetByName(sheetName);
      let targetSheet = targetSs.getSheetByName(sheetName);

      if (!targetSheet) {
        targetSheet = targetSs.insertSheet(sheetName);
      }

      if (sourceSheet) {
        const sourceData = sourceSheet.getDataRange().getValues();

        const lastRowTarget = targetSheet.getLastRow();
        if (lastRowTarget > 1) {
          targetSheet.getRange(2, 1, lastRowTarget - 1, targetSheet.getLastColumn()).clearContent();
        }

        if (sourceData.length > 1) {
          const rowsToInsert = sourceData.slice(1);
          const numRows = rowsToInsert.length;
          const numCols = rowsToInsert[0].length;

          if (targetSheet.getLastRow() === 0 || targetSheet.getRange(1, 1).getValue() === '') {
            targetSheet.getRange(1, 1, 1, sourceData[0].length).setValues([sourceData[0]])
                       .setFontWeight('bold');
            targetSheet.setFrozenRows(1);
          }

          targetSheet.getRange(2, 1, numRows, numCols).setValues(rowsToInsert);
          logs.push(`Sheet "${sheetName}": Berhasil diperbarui (${numRows} baris data disalin).`);
        } else {
          logs.push(`Sheet "${sheetName}": Berhasil dikosongkan (tidak ada data baru).`);
        }
        updatedSheets.push(sheetName);
      } else {
        logs.push(`Sheet "${sheetName}": Tidak ditemukan dalam file Excel yang diunggah.`);
      }
    });

    convertedFile.setTrashed(true);
    tempFile.setTrashed(true);

    const userEmail = Session.getActiveUser().getEmail() || "User System";
    const statusSummary = `Upload Excel Selesai. Sheet diproses: ${updatedSheets.join(", ")}.`;

    // Pemanggilan fungsi
    writeSimpleLog(userEmail, "Bulk Upload Excel Multi-Sheet", statusSummary);

    return {
      success: true,
      message: statusSummary,
      logs: logs
    };

  } catch (e) {
    throw new Error("Gagal memproses upload Excel: " + e.message);
  }
}

/**
 * Fungsi Pembantu Log (Harus berada di luar fungsi processExcelUpload)
 */
function writeSimpleLog(user, aksi, keterangan) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEETS.LOGS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.LOGS);
      sheet.appendRow(["Tanggal & Waktu", "User", "Aksi", "Keterangan"]);
    }
    sheet.appendRow([new Date(), user, aksi, keterangan]);
  } catch (e) {
    Logger.log("Gagal menulis log: " + e.message);
  }
}

/**
 * Fungsi Ambil Log (Harus berada di luar fungsi processExcelUpload)
 */
function getUploadLogs() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.LOGS);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const rows = data.slice(1);
    return rows.reverse().map((r, i) => ({
      no: i + 1,
      timestamp: r[0] ? Utilities.formatDate(new Date(r[0]), "GMT+7", "yyyy-MM-dd HH:mm:ss") : "-",
      user: r[1] || "-",
      aksi: r[2] || "-",
      keterangan: r[3] || "-"
    }));
  } catch (e) {
    return [];
  }
}

// ================= AUTHENTICATION & USER MANAGEMENT =================
const AUTH_USERS_SHEET = 'System_Users';
const AUTH_SESSION_HOURS = 12;
function authUsersSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(AUTH_USERS_SHEET);
  if (!sh) { sh = ss.insertSheet(AUTH_USERS_SHEET); sh.appendRow(['Name','Email','Role','PasswordHash','Salt','Active','GoogleSub','CreatedAt']); sh.setFrozenRows(1); sh.getRange(1,1,1,8).setFontWeight('bold').setBackground('#ef512e').setFontColor('#fff'); }
  return sh;
}
function setupInitialAdmin() {
  const sh = authUsersSheet_();
  if (sh.getLastRow() > 1) throw new Error('User sudah ada. Untuk keamanan, setup admin awal hanya boleh dijalankan pada database user kosong.');
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) throw new Error('Email akun Google tidak tersedia. Tambahkan admin melalui sheet System_Users atau jalankan dengan akun Workspace yang teridentifikasi.');
  const tempPassword = Utilities.getUuid().replace(/-/g,'') + 'A!';
  const salt = Utilities.getUuid();
  sh.appendRow(['Administrator',email,'Admin',passwordHash_(tempPassword,salt),salt,true,'',new Date()]);
  console.log('Admin dibuat: '+email+' | Password sementara: '+tempPassword+' | Segera login dan ganti password dengan memperbarui user.');
  return 'Admin dibuat. Lihat execution log untuk password sementara.';
}
function passwordHash_(password,salt) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt+'|'+password, Utilities.Charset.UTF_8)); }
function hashPassword(password, salt) { return passwordHash_(password, salt); }

function addUser(userData) {
  try {
    if (!userData || typeof userData !== 'object') {
      return { success: false, message: 'Data user tidak valid.' };
    }
    const name = String(userData.name || '').trim();
    const email = String(userData.email || '').trim().toLowerCase();
    const role = String(userData.role || 'Karyawan').trim();
    const password = String(userData.password || '');

    if (!name) {
      return { success: false, message: 'Nama tidak boleh kosong.' };
    }
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!email || !emailRegex.test(email)) {
      return { success: false, message: 'Format email tidak valid.' };
    }
    const validRoles = ['Admin', 'Karyawan', 'Customer'];
    if (!validRoles.includes(role)) {
      return { success: false, message: 'Role hanya boleh Admin, Karyawan, atau Customer.' };
    }
    if (!password || password.length < 8) {
      return { success: false, message: 'Password minimal 8 karakter.' };
    }

    const sh = authUsersSheet_();
    const existing = authFindUser_(email);
    if (existing) {
      return { success: false, message: 'Email sudah terdaftar' };
    }

    const salt = Utilities.getUuid();
    const passwordHash = passwordHash_(password, salt);

    // Urutan kolom: Name, Email, Role, PasswordHash, Salt, Active, GoogleSub, CreatedAt
    sh.appendRow([
      name,
      email,
      role,
      passwordHash,
      salt,
      true,
      '',
      new Date()
    ]);

    return { success: true, message: 'User berhasil ditambahkan' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getUsers() {
  try {
    const sh = authUsersSheet_();
    const values = sh.getDataRange().getValues();
    if (values.length <= 1) return [];
    return values.slice(1).map(function (r) {
      return {
        name: String(r[0] || ''),
        email: String(r[1] || ''),
        role: String(r[2] || 'Customer'),
        active: r[5] === true || String(r[5]).toLowerCase() === 'true'
      };
    });
  } catch (error) {
    return [];
  }
}

function getUserByEmail(email) {
  try {
    email = String(email || '').trim().toLowerCase();
    if (!email) {
      return { success: false, message: 'Email tidak boleh kosong.' };
    }
    const u = authFindUser_(email);
    if (!u) {
      return { success: false, message: 'User tidak ditemukan' };
    }
    return {
      success: true,
      data: {
        name: u.name,
        email: u.email,
        role: u.role,
        active: u.active
      }
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function updateUser(userData) {
  try {
    if (!userData || typeof userData !== 'object') {
      return { success: false, message: 'Data user tidak valid.' };
    }
    const originalEmail = String(userData.originalEmail || '').trim().toLowerCase();
    const name = String(userData.name || '').trim();
    const email = String(userData.email || '').trim().toLowerCase();
    const role = String(userData.role || '').trim();
    const password = String(userData.password || '');
    const activeBool = userData.active === true || String(userData.active).toLowerCase() === 'true';

    if (!originalEmail) {
      return { success: false, message: 'Original email tidak valid.' };
    }
    if (!name) {
      return { success: false, message: 'Nama tidak boleh kosong.' };
    }
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!email || !emailRegex.test(email)) {
      return { success: false, message: 'Format email tidak valid.' };
    }
    const validRoles = ['Admin', 'Karyawan', 'Customer'];
    if (!validRoles.includes(role)) {
      return { success: false, message: 'Role hanya boleh Admin, Karyawan, atau Customer.' };
    }

    const sh = authUsersSheet_();
    const existing = authFindUser_(originalEmail);
    if (!existing) {
      return { success: false, message: 'User tidak ditemukan' };
    }

    // Cek duplikasi email jika email diubah
    if (email !== originalEmail) {
      const duplicate = authFindUser_(email);
      if (duplicate && duplicate.row !== existing.row) {
        return { success: false, message: 'Email sudah digunakan user lain' };
      }
    }

    // Update password jika diisi, jika kosong pertahankan PasswordHash dan Salt lama
    if (password) {
      if (password.length < 8) {
        return { success: false, message: 'Password minimal 8 karakter.' };
      }
      const salt = Utilities.getUuid();
      const passwordHash = passwordHash_(password, salt);
      sh.getRange(existing.row, 1, 1, 6).setValues([[name, email, role, passwordHash, salt, activeBool]]);
    } else {
      sh.getRange(existing.row, 1, 1, 3).setValues([[name, email, role]]);
      sh.getRange(existing.row, 6).setValue(activeBool);
    }

    return { success: true, message: 'User berhasil diperbarui' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}
function authFindUser_(email) {
  const sh=authUsersSheet_(), values=sh.getDataRange().getValues(); email=String(email||'').trim().toLowerCase();
  for(let i=1;i<values.length;i++) if(String(values[i][1]).trim().toLowerCase()===email) return {row:i+1,name:String(values[i][0]||''),email:String(values[i][1]),role:String(values[i][2]||'Customer'),hash:String(values[i][3]||''),salt:String(values[i][4]||''),active:values[i][5]===true||String(values[i][5]).toLowerCase()==='true',googleSub:String(values[i][6]||'')};
  return null;
}
function createAuthSession_(u) { const token=Utilities.getUuid()+Utilities.getUuid(); CacheService.getScriptCache().put('auth:'+token,JSON.stringify({email:u.email,role:u.role,name:u.name}),AUTH_SESSION_HOURS*3600); return {success:true,token:token,user:{name:u.name,email:u.email,role:u.role},expiresIn:AUTH_SESSION_HOURS*3600}; }
function loginWithPassword(email,password) { const u=authFindUser_(email); if(!u||!u.active||!u.hash||passwordHash_(String(password||''),u.salt)!==u.hash) return {success:false,message:'Email atau password tidak valid, atau akun nonaktif.'}; return createAuthSession_(u); }
function validateSession(token) { const raw=CacheService.getScriptCache().get('auth:'+String(token||'')); if(!raw)return {success:false,message:'Sesi telah berakhir. Silakan login kembali.'}; const d=JSON.parse(raw),u=authFindUser_(d.email); if(!u||!u.active)return {success:false,message:'Akun tidak aktif.'}; return {success:true,token:token,user:{name:u.name,email:u.email,role:u.role},expiresIn:AUTH_SESSION_HOURS*3600}; }
function logoutSession(token) { if(token)CacheService.getScriptCache().remove('auth:'+String(token)); return {success:true}; }
function requireAdmin_(token) { const s=validateSession(token); if(!s.success)throw new Error('Sesi tidak valid. Silakan login kembali.'); if(s.user.role!=='Admin')throw new Error('Akses ditolak. Hanya Admin yang dapat mengelola pengguna.'); return s.user; }
function listUsers(token) { requireAdmin_(token); return getUsers(); }
function getUserForAdmin(token,email) { requireAdmin_(token); const u=authFindUser_(email); if(!u)throw new Error('User tidak ditemukan.'); return {name:u.name,email:u.email,role:u.role,active:u.active}; }
function saveUserForAdmin(token,d) { requireAdmin_(token); d=d||{}; const name=String(d.name||'').trim(),email=String(d.email||'').trim().toLowerCase(),role=String(d.role||''); if(!name||!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)||!['Admin','Karyawan','Customer'].includes(role))return {success:false,message:'Periksa nama, email, dan role.'}; const sh=authUsersSheet_(),existing=d.originalEmail?authFindUser_(d.originalEmail):null,duplicate=authFindUser_(email); if(duplicate&&(!existing||duplicate.email!==existing.email))return {success:false,message:'Email sudah terdaftar.'}; const password=String(d.password||''); if(!existing&&password.length<8)return {success:false,message:'Password minimal 8 karakter.'}; if(existing){sh.getRange(existing.row,1,1,3).setValues([[name,email,role]]);if(password){const salt=Utilities.getUuid();sh.getRange(existing.row,4,1,2).setValues([[passwordHash_(password,salt),salt]]);}if(email!==existing.email)CacheService.getScriptCache().removeAll([]);}else{const salt=Utilities.getUuid();sh.appendRow([name,email,role,passwordHash_(password,salt),salt,true,'',new Date()]);}return {success:true}; }
function setUserActive(token,email,active) { requireAdmin_(token); const u=authFindUser_(email); if(!u)throw new Error('User tidak ditemukan.'); authUsersSheet_().getRange(u.row,6).setValue(!!active); return {success:true}; }
function loginWithGoogle(credential) {
  const clientId = PropertiesService.getScriptProperties()
    .getProperty('GOOGLE_CLIENT_ID');

  if (!clientId) {
    return {
      success:false,
      message:'Login Google belum dikonfigurasi.'
    };
  }

  const resp = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' +
    encodeURIComponent(String(credential || '')),
    {
      muteHttpExceptions:true
    }
  );

  if (resp.getResponseCode() !== 200) {
    return {
      success:false,
      message:'Token Google tidak valid.'
    };
  }


  const payload = JSON.parse(resp.getContentText());


  // validasi client id
  if (
    payload.aud !== clientId ||
    payload.email_verified !== 'true'
  ) {
    return {
      success:false,
      message:'Akun Google tidak dapat diverifikasi.'
    };
  }


  const email = String(payload.email)
    .trim()
    .toLowerCase();


  let user = authFindUser_(email);


  // ======================================
  // AUTO CREATE CUSTOMER GOOGLE LOGIN
  // ======================================
  if (!user) {

    const sh = authUsersSheet_();

    sh.appendRow([
      payload.name || email.split('@')[0], // Name
      email,                              // Email
      'Customer',                         // DEFAULT ROLE
      '',                                 // PasswordHash
      '',                                 // Salt
      true,                               // Active
      payload.sub,                        // GoogleSub
      new Date()                          // CreatedAt
    ]);


    user = authFindUser_(email);

  } else {

    // jika user lama belum punya Google ID
    if (!user.googleSub) {

      authUsersSheet_()
        .getRange(user.row,7)
        .setValue(payload.sub);

    }

    // keamanan akun
    if (
      user.googleSub &&
      user.googleSub !== payload.sub
    ) {

      return {
        success:false,
        message:'Akun Google tidak sesuai.'
      };

    }

  }

  // Validasi status akun sebelum membuat sesi
  if (!user || !user.active) {
    return {
      success: false,
      message: 'Akun Anda tidak aktif'
    };
  }

  return createAuthSession_(user);
}

/**
 * Endpoint alternatif googleLogin sesuai spesifikasi PRD
 * Menerima payload langsung { email, name, googleSub }
 */
function googleLogin(data) {
  try {
    if (!data || typeof data !== 'object') {
      return { success: false, message: 'Data Google login tidak valid.' };
    }

    const email = String(data.email || '').trim().toLowerCase();
    const name = String(data.name || '').trim();
    const googleSub = String(data.googleSub || '').trim();

    if (!email) {
      return { success: false, message: 'Email Google tidak ditemukan' };
    }

    let user = findUserByEmail(email);

    if (!user) {
      user = createGoogleCustomer(name, email, googleSub);
    } else {
      if (!user.googleSub && googleSub) {
        authUsersSheet_().getRange(user.row, 7).setValue(googleSub);
        user.googleSub = googleSub;
      }
      if (user.googleSub && googleSub && user.googleSub !== googleSub) {
        return { success: false, message: 'Akun Google tidak sesuai.' };
      }
    }

    if (!user.active) {
      return { success: false, message: 'Akun Anda tidak aktif' };
    }

    return createAuthSession_(user);
  } catch (error) {
    return { success: false, message: 'Login Google gagal: ' + error.message };
  }
}

function findUserByEmail(email) {
  return authFindUser_(email);
}

function createGoogleCustomer(name, email, googleSub) {
  const sh = authUsersSheet_();
  sh.appendRow([
    name || email.split('@')[0],
    email,
    'Customer',
    '',
    '',
    true,
    googleSub || '',
    new Date()
  ]);
  return authFindUser_(email);
}

/**
 * Login langsung menggunakan akun Google yang sedang aktif di browser
 * Bebas dari issue origin_mismatch Google Cloud OAuth
 */
function loginWithActiveGoogleAccount() {
  try {
    let email = '';
    try {
      email = Session.getActiveUser().getEmail();
    } catch (e) {
      email = '';
    }

    if (!email) {
      return {
        success: false,
        message: 'Tidak dapat mendeteksi akun Google aktif. Pastikan Anda sudah login ke browser dengan akun Google.'
      };
    }

    email = String(email).trim().toLowerCase();
    let user = authFindUser_(email);

    if (!user) {
      user = createGoogleCustomer(email.split('@')[0], email, '');
    }

    if (!user || !user.active) {
      return {
        success: false,
        message: 'Akun Anda tidak aktif'
      };
    }

    return createAuthSession_(user);
  } catch (err) {
    return {
      success: false,
      message: 'Gagal login Google: ' + err.message
    };
  }
}

function getGoogleClientId(){return PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_ID')||'';}


/** Reset password Admin secara manual. Hapus fungsi ini setelah digunakan. */
function resetAdminPassword() {
  const emailAdmin = 'bayukevinfarindra23@gmail.com';
  const passwordBaru = 'admin123';

  const user = authFindUser_(emailAdmin);

  if (!user) {
    throw new Error('Akun Admin tidak ditemukan: ' + emailAdmin);
  }

  if (user.role !== 'Admin') {
    throw new Error('Akun tersebut bukan Admin.');
  }

  const salt = Utilities.getUuid();
  const hash = passwordHash_(passwordBaru, salt);
  const sheet = authUsersSheet_();

  sheet.getRange(user.row, 4, 1, 2)
    .setValues([[hash, salt]]);

  console.log('Password Admin berhasil direset: ' + user.email);
}
