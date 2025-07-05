// --- Global Constants (Konstanta Global) ---
const SPREADSHEET_ID = '1eZxURW6yArfKbksTBQ_3fAKfcwS4m9drArefsuiYRdU'; // Ganti dengan ID Spreadsheet Anda
const DATA_MASTER_SHEET_NAME = 'MASTER LENDING'; // Nama sheet utama yang berisi data transaksi lending
const MASTER_KARYAWAN_SHEET_NAME = 'MASTER KARYAWAN'; // Nama sheet untuk data marketing/target

// --- Web App Entry Point ---
/**
 * Fungsi untuk melayani permintaan web (GET).
 * Ini akan memuat file HTML.
 */
function doGet() {
  const htmlOutput = HtmlService
    .createHtmlOutputFromFile('PerformanceMonitoring.html') // Pastikan nama file HTML benar
    .setTitle('System Monitoring Kinerja Lending Glory Victory Anandita')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return htmlOutput;
}

// --- Public Functions (Dipanggil dari Frontend/HTML) ---

/**
 * Fungsi untuk mendapatkan semua nama marketing aktif dari MASTER KARYAWAN.
 * Dipanggil oleh HTML untuk mengisi filter marketing.
 * @returns {string[]} Array of marketing names.
 */
function getAllMarketingNames() {
  Logger.log('Memulai getAllMarketingNames');
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const marketingSheet = ss.getSheetByName(MASTER_KARYAWAN_SHEET_NAME);

    if (!marketingSheet) {
      throw new Error(`Sheet "${MASTER_KARYAWAN_SHEET_NAME}" tidak ditemukan.`);
    }

    const data = marketingSheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log('MASTER KARYAWAN sheet kosong atau hanya berisi header.');
      return [];
    }

    const headers = data[0];
    const nameColIdx = getColumnIndex(headers, 'NAMA');
    const statusColIdx = getColumnIndex(headers, 'STATUS');

    const uniqueNames = new Set();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = (row[statusColIdx] || '').toString().trim().toLowerCase();
      const name = (row[nameColIdx] || '').toString().trim();
      if (status === 'aktif' && name) {
        uniqueNames.add(name);
      }
    }
    const sortedNames = [...Array.from(uniqueNames).sort()];
    Logger.log(`Marketing aktif ditemukan: ${sortedNames.join(', ')}`);
    return sortedNames;
  } catch (e) {
    Logger.log(`ERROR in getAllMarketingNames: ${e.message}`);
    throw e;
  }
}

/**
 * Fungsi utama untuk mendapatkan data dashboard berdasarkan filter.
 * @param {number} month Bulan yang dipilih (1-12).
 * @param {number} year Tahun yang dipilih.
 * @param {string} selectedMarketingParam Nama marketing yang dipilih, atau "All".
 * @param {string} periodType Tipe periode ('monthly' atau 'weekly').
 * @param {number} weekNumber Nomor minggu (jika periodType 'weekly').
 * @param {number} weeklyYear Tahun untuk nomor minggu (jika periodType 'weekly').
 * @returns {Object} Objek berisi ringkasan, data performa, dan data grafik.
 */
function getDashboardData(month, year, selectedMarketingParam, periodType, weekNumber, weeklyYear) {
  Logger.log(`[START] getDashboardData with params: month=${month}, year=${year}, selectedMarketingParam=${selectedMarketingParam}, periodType=${periodType}, weekNumber=${weekNumber}, weeklyYear=${weeklyYear}`);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const lendingSheet = ss.getSheetByName(DATA_MASTER_SHEET_NAME);
  const karyawanSheet = ss.getSheetByName(MASTER_KARYAWAN_SHEET_NAME);

  if (!lendingSheet || !karyawanSheet) {
    Logger.log("Spreasheet 'MASTER LENDING' or 'MASTER KARYAWAN' not found. Returning default data.");
    return getDefaultDashboardData();
  }

  const lendingRows = lendingSheet.getDataRange().getDisplayValues();
  const karyawanRows = karyawanSheet.getDataRange().getDisplayValues();

  if (lendingRows.length <= 1 || karyawanRows.length <= 1) {
    Logger.log('One or both sheets are empty or only contain headers. Returning default data.');
    return getDefaultDashboardData();
  }

  const lendingHeader = lendingRows[0];
  const karyawanHeader = karyawanRows[0];

  const lendingColMap = createHeaderMap(lendingHeader, DATA_MASTER_SHEET_NAME, ["TANGGAL PENCAIRAN", "MARKETING", "STATUS", "PINJAMAN DITERIMA", "FEE MARKETING", "PLAFON DIAJUKAN"]);
  const karyawanColMap = createHeaderMap(karyawanHeader, MASTER_KARYAWAN_SHEET_NAME, ["NAMA", "TARGET", "PENAWARAN GAJI", "STATUS"]);

  const initialMarketingPerformance = initializeMarketingPerformance(karyawanRows, karyawanColMap);
  let activeMarketersCount = Object.keys(initialMarketingPerformance).length;

  // Filter data lending dan hitung total prospek/closing serta prospek per marketing
  const { filteredLendingRows, totalProspectCount, totalClosingCount, marketingSpecificCounts } = filterLendingData(
    lendingRows, lendingColMap, month, year, selectedMarketingParam, periodType, weekNumber, weeklyYear
  );

  // Buat salinan marketingPerformance untuk perhitungan inflow/outflow
  let marketingPerformance = JSON.parse(JSON.stringify(initialMarketingPerformance));

  // Agregasi data inflow/outflow dari baris yang difilter ke marketingPerformance
  aggregateLendingData(filteredLendingRows, lendingColMap, marketingPerformance);

  // Update totalProspect per marketing dari marketingSpecificCounts
  for (const marketerName in marketingSpecificCounts) {
    if (marketingPerformance[marketerName]) {
      marketingPerformance[marketerName].totalProspect = marketingSpecificCounts[marketerName].totalProspect || 0;
    }
  }

  // Hitung total keseluruhan target dari semua marketing aktif
  let totalOverallTarget = 0;
  for (const marketerName in initialMarketingPerformance) {
    totalOverallTarget += initialMarketingPerformance[marketerName].target;
  }

  // Hitung gaji dan komisi
  const { performanceData, gajiKomisiData, totalGajiKeseluruhan, totalKomisiKeseluruhan } = calculateGajiKomisi(
    marketingPerformance
  );

  // Siapkan data untuk grafik
  const chartData = prepareChartData(performanceData, marketingPerformance);

  const summary = {
    totalGrowth: marketingPerformance.hasOwnProperty('All') ? marketingPerformance['All'].inflow - marketingPerformance['All'].outflow : 0, // Ini mungkin perlu disesuaikan jika 'All' tidak ada
    overallTargetPercentage: totalOverallTarget > 0 ? (getSumOfInflow(marketingPerformance) / totalOverallTarget) * 100 : 0,
    totalInflow: getSumOfInflow(marketingPerformance), // Mengambil total inflow dari semua marketer
    totalOutflow: getSumOfOutflow(marketingPerformance), // Mengambil total outflow dari semua marketer
    totalGaji: totalGajiKeseluruhan,
    totalKomisi: totalKomisiKeseluruhan,
    activeMarketers: activeMarketersCount,
    totalProspect: totalProspectCount, // ini total prospek global dari filterLendingData
    totalClosing: totalClosingCount, // ini total closing global dari filterLendingData
    totalTarget: totalOverallTarget
  };

  const result = {
    summary: summary,
    performanceData: performanceData,
    gajiKomisiData: gajiKomisiData,
    top5Kinerja: chartData.top5Kinerja,
    bottom5Kinerja: chartData.bottom5Kinerja,
    top5Growth: chartData.top5Growth,
    negativeGrowth: chartData.negativeGrowth,
    growthVsTarget: chartData.growthVsTarget
  };

  Logger.log("[END] getDashboardData. Returning object (excerpt performanceData): " + JSON.stringify(result.performanceData.slice(0, 5)));
  Logger.log("[END] getDashboardData. Returning object (excerpt gajiKomisiData): " + JSON.stringify(result.gajiKomisiData.slice(0, 5)));

  return result;
}

// --- Helper Functions (Fungsi Pembantu) ---

/**
 * Mengembalikan objek map header ke index kolom.
 * Memvalidasi keberadaan header yang diperlukan.
 * @param {string[]} headers Array of header names.
 * @param {string} sheetName Nama sheet untuk logging error.
 * @param {string[]} requiredHeaders Array of required header names.
 * @returns {Object} Map of header (trimmed and uppercased) to column index.
 * @throws {Error} Jika ada header yang diperlukan tidak ditemukan.
 */
function createHeaderMap(headers, sheetName, requiredHeaders = []) {
  const headerMap = {};
  headers.forEach((header, index) => {
    headerMap[header.toString().trim().toUpperCase()] = index;
  });

  for (const header of requiredHeaders) {
    if (!(header.toUpperCase() in headerMap)) {
      throw new Error(`Kolom "${header}" tidak ditemukan di sheet "${sheetName}". Pastikan penulisan header di Google Sheet sama persis.`);
    }
  }
  Logger.log(`Header map for ${sheetName}: ${JSON.stringify(headerMap)}`);
  return headerMap;
}

/**
 * Mengambil indeks kolom dari map yang sudah dibuat.
 * @param {string[]} headers Array of header names.
 * @param {string} headerName Nama header yang dicari.
 * @returns {number} Index kolom.
 * @throws {Error} Jika header tidak ditemukan.
 */
function getColumnIndex(headers, headerName) {
  const index = headers.indexOf(headerName);
  if (index === -1) {
    throw new Error(`Kolom '${headerName}' tidak ditemukan.`);
  }
  return index;
}

/**
 * Menginisialisasi objek kinerja marketing dari data MASTER KARYAWAN.
 * @param {Array<Array>} masterKaryawanRawData Data dari MASTER KARYAWAN sheet.
 * @param {Object} karyawanColMap Map header to column index for MASTER KARYAWAN.
 * @returns {Object} Objek kinerja marketing.
 */
function initializeMarketingPerformance(masterKaryawanRawData, karyawanColMap) {
  const marketingPerformance = {};
  const colKaryawanName = karyawanColMap['NAMA'];
  const colKaryawanGajiKontrak = karyawanColMap['PENAWARAN GAJI'];
  const colKaryawanTarget = karyawanColMap['TARGET'];
  const colKaryawanStatusAktif = karyawanColMap['STATUS'];

  for (let i = 1; i < masterKaryawanRawData.length; i++) {
    const row = masterKaryawanRawData[i];
    const name = (row[colKaryawanName] || '').toString().trim();
    const status = (row[colKaryawanStatusAktif] || '').toString().trim().toLowerCase();
    const gajiKontrak = parseNumber(row[colKaryawanGajiKontrak]);
    const target = parseNumber(row[colKaryawanTarget]);

    if (status === 'aktif' && name) {
      marketingPerformance[name] = {
        inflow: 0,
        outflow: 0,
        feeMarketingCair: 0,
        totalProspect: 0, // NEW: For storing prospect count per marketer
        target: target,
        gajiKontrakAwal: gajiKontrak,
        totalGaji: 0,
        totalKomisi: 0
      };
    }
  }
  Logger.log(`Initial marketingPerformance setup: ${JSON.stringify(marketingPerformance)}`);
  return marketingPerformance;
}

/**
 * Memfilter data lending berdasarkan parameter yang diberikan
 * dan menghitung total prospect serta closing.
 * @param {Array<Array>} masterLendingRawData Data dari MASTER LENDING sheet.
 * @param {Object} lendingColMap Map header to column index for MASTER LENDING.
 * @param {number} month Bulan yang dipilih.
 * @param {number} year Tahun yang dipilih.
 * @param {string} selectedMarketingParam Nama marketing yang dipilih atau 'All'.
 * @param {string} periodType Tipe periode ('monthly' atau 'weekly').
 * @param {number} weekNumber Nomor minggu.
 * @param {number} weeklyYear Tahun untuk nomor minggu.
 * @returns {Object} Objek berisi baris yang difilter, total prospect, total closing, dan marketing-specific counts.
 */
function filterLendingData(masterLendingRawData, lendingColMap, month, year, selectedMarketingParam, periodType, weekNumber, weeklyYear) {
  const filteredLendingRows = [];
  let totalClosingCount = 0;
  let totalProspectCount = 0;
  const marketingSpecificCounts = {}; // NEW: Object to store prospect count per marketing

  const colLendingTanggalPencairan = lendingColMap['TANGGAL PENCAIRAN'];
  const colLendingMarketingName = lendingColMap['MARKETING'];
  const colLendingStatus = lendingColMap['STATUS'];

  const validClosingStatuses = ['approve', 'claimed'];
  const validProspectStatuses = ['pending', 'verifikasi', 'survey'];

  for (let i = 1; i < masterLendingRawData.length; i++) {
    const row = masterLendingRawData[i];
    const dateCell = row[colLendingTanggalPencairan];
    const date = (dateCell instanceof Date) ? dateCell : new Date(dateCell);

    if (isNaN(date.getTime())) {
      // Logger.log(`Skipping row ${i} due to invalid date: ${dateCell}`);
      continue;
    }

    const status = (row[colLendingStatus] || '').toString().trim().toLowerCase();
    const marketingNameInRow = (row[colLendingMarketingName] || '').toString().trim();

    // Check date match
    let isDateMatch = false;
    if (periodType === 'monthly' && month !== null && year !== null) {
      isDateMatch = date.getMonth() + 1 === parseInt(month) && date.getFullYear() === parseInt(year);
    } else if (periodType === 'weekly' && weekNumber !== null && weeklyYear !== null) {
      const rowWeekNumber = getWeekNumber(date);
      isDateMatch = weekNumberCalculated === parseInt(weekNumber) && date.getFullYear() === parseInt(weeklyYear);
    } else {
      // If no periodType or incomplete date params, allow all dates for now
      // Or you might want to default to `false` based on desired behavior
      isDateMatch = true; // Example: if no filter applied, all dates match
    }


    // Check marketing match
    const isMarketingMatch = (selectedMarketingParam === 'All' || !selectedMarketingParam || marketingNameInRow === selectedMarketingParam);

    // Aggregate based on date and marketing filters
    if (isDateMatch && isMarketingMatch) {
      if (validClosingStatuses.includes(status)) {
        filteredLendingRows.push(row);
        totalClosingCount++;
      }
    }

    // Always count prospects for ALL marketers regardless of the selectedMarketingParam filter,
    // but within the selected date period.
    if (isDateMatch && validProspectStatuses.includes(status)) {
      totalProspectCount++; // Global prospect count for the period
      if (!marketingSpecificCounts[marketingNameInRow]) {
        marketingSpecificCounts[marketingNameInRow] = { totalProspect: 0 };
      }
      marketingSpecificCounts[marketingNameInRow].totalProspect++;
    }
  }

  Logger.log(`[END] filterLendingData. Filtered Rows (for aggregation): ${filteredLendingRows.length}. Total Closing Count: ${totalClosingCount}. Total Prospect Count: ${totalProspectCount}`);
  Logger.log("Marketing Specific Counts (from filterLendingData): " + JSON.stringify(marketingSpecificCounts));
  return { filteredLendingRows, totalProspectCount, totalClosingCount, marketingSpecificCounts };
}


/**
 * Mengagregasi data inflow, outflow, dan fee marketing dari baris yang difilter.
 * Memperbarui objek marketingPerformance.
 * @param {Array<Array>} filteredLendingRows Baris data lending yang sudah difilter.
 * @param {Object} lendingColMap Map header to column index for MASTER LENDING.
 * @param {Object} marketingPerformance Objek kinerja marketing yang akan diperbarui.
 * @returns {Object} Objek berisi total inflow global dan total outflow global.
 */
function aggregateLendingData(filteredLendingRows, lendingColMap, marketingPerformance) {
  let totalInflowGlobal = 0;
  let totalOutflowGlobal = 0;

  const colLendingPlafonDiajukan = lendingColMap['PLAFON DIAJUKAN'];
  const colLendingPinjamanDiterima = lendingColMap['PINJAMAN DITERIMA'];
  const colLendingFeeMarketing = lendingColMap['FEE MARKETING'];
  const colLendingMarketingName = lendingColMap['MARKETING'];
  const colLendingStatus = lendingColMap['STATUS'];

  const validClosingStatuses = ['approve', 'claimed'];

  Logger.log(`Starting aggregateLendingData with ${filteredLendingRows.length} filtered rows.`);

  for (const row of filteredLendingRows) {
    const marketingNameInRow = (row[colLendingMarketingName] || '').toString().trim();
    const statusInRow = (row[colLendingStatus] || '').toString().trim().toLowerCase();

    // Only aggregate inflow and feeMarketingCair for valid closing statuses
    if (validClosingStatuses.includes(statusInRow)) {
      const pinjaman = parseNumber(row[colLendingPinjamanDiterima]);
      const feeMarketingRaw = row[colLendingFeeMarketing];
      const feePercentage = parseNumber(feeMarketingRaw, true);

      const feeMarketingAmount = pinjaman * feePercentage;

      if (marketingPerformance[marketingNameInRow]) {
        marketingPerformance[marketingNameInRow].inflow += pinjaman;
        marketingPerformance[marketingNameInRow].feeMarketingCair += feeMarketingAmount;
        totalInflowGlobal += pinjaman;
        Logger.log(`Aggregating for ${marketingNameInRow} (Closing): Inflow=${pinjaman}, Fee=${feeMarketingAmount}. Current Inflow for ${marketingNameInRow}: ${marketingPerformance[marketingNameInRow].inflow}`);
      } else {
        Logger.log(`Marketing name "${marketingNameInRow}" from lending data not found in MASTER KARYAWAN. Skipping performance aggregation for this closing entry: ${JSON.stringify(row)}`);
      }
    }

    // Outflow (if it represents canceled/rejected loans) - needs specific status
    // Assuming 'PLAFON DIAJUKAN' is for inflow calculation for now, adjust if it's for outflow
    // If 'Plafon Diajukan' is always considered, include it here.
    // const plafon = parseNumber(row[colLendingPlafonDiajukan]);
    // totalOutflowGlobal += plafon; // Uncomment and adjust if plafon represents outflow
  }
  // This function currently focuses on inflow and fee for marketingPerformance.
  // totalOutflowGlobal calculation logic might need to be refined based on 'outflow' definition in your sheet.
  return { totalInflowGlobal, totalOutflowGlobal };
}


/**
 * Menghitung gaji dan komisi untuk setiap marketing.
 * Memperbarui objek marketingPerformance dengan totalGaji dan totalKomisi.
 * @param {Object} marketingPerformance Objek kinerja marketing yang sudah diperbarui dengan inflow/outflow.
 * @returns {Object} Objek berisi data performa, data gaji/komisi, dan total gaji/komisi keseluruhan.
 */
function calculateGajiKomisi(marketingPerformance) {
  let totalGajiKeseluruhan = 0;
  let totalKomisiKeseluruhan = 0;
  const performanceData = []; // Untuk tabel utama
  const gajiKomisiData = []; // Untuk tabel rincian gaji/komisi

  Logger.log(`Starting calculateGajiKomisi. Marketing performance keys: ${Object.keys(marketingPerformance).join(', ')}`);

  for (const name in marketingPerformance) {
    const data = marketingPerformance[name];
    const target = data.target;
    const pencapaianInflow = data.inflow;
    const kinerja = target > 0 ? (pencapaianInflow / target) : 0;

    Logger.log(`Calculating for Marketing ${name}: Inflow=${pencapaianInflow}, Target=${target}, Kinerja=${kinerja.toFixed(2)}`);

    // Rule 1: Gaji Kontrak (Total Gaji) - Dibayarkan pro-rata jika pencapaian < 50%
    let gajiFinal = data.gajiKontrakAwal;
    if (kinerja < 0.5) {
      gajiFinal = data.gajiKontrakAwal * kinerja;
      Logger.log(`Marketing ${name}: Kinerja ${kinerja.toFixed(2)} (<0.5). Gaji pro-rata dari ${data.gajiKontrakAwal} menjadi ${gajiFinal.toFixed(2)}`);
    } else {
      Logger.log(`Marketing ${name}: Kinerja ${kinerja.toFixed(2)} (>=0.5). Gaji penuh: ${gajiFinal.toFixed(2)}`);
    }
    data.totalGaji = gajiFinal;
    totalGajiKeseluruhan += gajiFinal;

    // Rule 3: Fee 0.25% dari setiap kredit cair (ini adalah 'FEE MARKETING' di sheet)
    let komisiFeeKreditCair = data.feeMarketingCair; // Ini sudah dihitung di aggregateLendingData

    // Rule 2: Reward Tambahan 0.25% jika achievement > 100%
    let rewardTambahan = 0;
    if (kinerja > 1.0) {
      rewardTambahan = 0.0025 * pencapaianInflow;
      Logger.log(`Marketing ${name}: Kinerja ${kinerja.toFixed(2)} (>1.0). Reward tambahan: ${rewardTambahan.toFixed(2)}`);
    }

    let totalKomisiPerMarketing = komisiFeeKreditCair + rewardTambahan;
    data.totalKomisi = totalKomisiPerMarketing;
    totalKomisiKeseluruhan += totalKomisiPerMarketing;

    performanceData.push({
      name: name,
      performance: kinerja, // Sebagai desimal, misal 0.85 untuk 85%
      commission: totalKomisiPerMarketing,
      inflow: pencapaianInflow,
      target: target,
      totalProspect: data.totalProspect || 0 // Include totalProspect here
    });

    gajiKomisiData.push({
      name: name,
      gajiKontrakAwal: data.gajiKontrakAwal,
      kinerja: kinerja,
      gajiFinal: gajiFinal,
      feeMarketingCair: komisiFeeKreditCair,
      rewardTambahan: rewardTambahan,
      totalKomisi: totalKomisiPerMarketing,
      totalDibayarkan: gajiFinal + totalKomisiPerMarketing
    });
  }
  Logger.log(`[END] calculateGajiKomisi. Final totalGajiKeseluruhan: ${totalGajiKeseluruhan}, totalKomisiKeseluruhan: ${totalKomisiKeseluruhan}`);
  // Sort performanceData by performance in descending order for main table display
  performanceData.sort((a, b) => b.performance - a.performance);

  return { performanceData, gajiKomisiData, totalGajiKeseluruhan, totalKomisiKeseluruhan };
}

/**
 * Menyiapkan data untuk berbagai grafik.
 * @param {Array<Object>} performanceData Data performa marketing yang sudah dihitung.
 * @param {Object} marketingPerformance Objek kinerja marketing mentah (untuk growthVsTarget).
 * @returns {Object} Objek berisi data untuk top 5 kinerja, bottom 5 kinerja, top 5 growth, negative growth, dan growth vs target.
 */
function prepareChartData(performanceData, marketingPerformance) {
  const sortedByPerformance = [...performanceData].sort((a, b) => b.performance - a.performance);
  const sortedByInflow = [...performanceData].sort((a, b) => b.inflow - a.inflow);

  const top5Kinerja = sortedByPerformance.slice(0, 5).map(item => ({
    label: item.name,
    value: item.performance * 100 // Convert to percentage
  }));
  // Untuk bottom 5, pastikan mengambil dari data yang sudah disortir dan mungkin membalik urutannya jika ingin dari terendah
  const bottom5Kinerja = sortedByPerformance.filter(item => item.performance > 0).slice(-5).reverse().map(item => ({
    label: item.name,
    value: item.performance * 100
  }));

  const top5Growth = sortedByInflow.slice(0, 5).map(item => ({
    label: item.name,
    value: item.inflow
  }));

  const negativeGrowthList = performanceData.filter(item => item.performance < 0.5 && item.inflow > 0) // Kinerja < 50% dan ada inflow
    .map(item => ({
      name: item.name,
      growth: item.inflow,
      target: item.target,
      kinerja: item.performance * 100
    }))
    .sort((a, b) => a.kinerja - b.kinerja); // Urutkan dari kinerja terendah

  const growthVsTargetLabels = [];
  const growthVsTargetGrowthValues = [];
  const growthVsTargetTargetValues = [];

  for (const name in marketingPerformance) {
    const data = marketingPerformance[name];
    growthVsTargetLabels.push(name);
    growthVsTargetGrowthValues.push(data.inflow);
    growthVsTargetTargetValues.push(data.target);
  }
  const growthVsTargetData = {
    labels: growthVsTargetLabels,
    growth: growthVsTargetGrowthValues,
    target: growthVsTargetTargetValues
  };

  Logger.log(`[END] prepareChartData. Top 5 Kinerja: ${JSON.stringify(top5Kinerja)}`);

  return {
    top5Kinerja: { labels: top5Kinerja.map(d => d.label), values: top5Kinerja.map(d => d.value) },
    bottom5Kinerja: { labels: bottom5Kinerja.map(d => d.label), values: bottom5Kinerja.map(d => d.value) },
    top5Growth: { labels: top5Growth.map(d => d.label), values: top5Growth.map(d => d.value) },
    negativeGrowth: negativeGrowthList,
    growthVsTarget: growthVsTargetData
  };
}


/**
 * Menghitung nomor minggu dalam setahun untuk tanggal yang diberikan.
 * @param {Date} d Objek tanggal.
 * @returns {number} Nomor minggu.
 */
function getWeekNumber(d) {
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  // Set to nearest Thursday: current date + 4 - current day number.
  // Make Sunday's day number 7
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  // Get first day of year
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  // Calculate full weeks to the nearest Thursday
  var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return weekNo;
}

/**
 * Mengubah nilai string menjadi angka, menangani format mata uang dan persentase.
 * @param {*} value Nilai yang akan di-parse.
 * @param {boolean} isPercentage Apakah nilai ini adalah persentase.
 * @returns {number} Nilai angka yang di-parse.
 */
function parseNumber(value, isPercentage = false) {
  if (typeof value === 'number') {
    return value;
  }
  if (typeof value === 'string') {
    let cleanValue = value.replace(/[Rp$,.]/g, '').trim(); // Remove Rp, $, and . (for thousands separator)
    // Handle comma as decimal separator if it's the only one
    if (cleanValue.includes(',') && cleanValue.indexOf(',') === cleanValue.lastIndexOf(',')) {
      cleanValue = cleanValue.replace(/,/g, '.');
    }
    if (isPercentage) {
      cleanValue = cleanValue.replace(/%/g, ''); // Remove % for percentages
    }
    const num = parseFloat(cleanValue);
    if (!isNaN(num)) {
      return isPercentage ? num / 100 : num;
    }
  }
  // Logger.log(`Failed to parse number: "${value}". Returning 0.`); // Uncomment for debugging
  return 0; // Default to 0 if parsing fails
}

/**
 * Mengembalikan data dashboard default jika terjadi kesalahan atau sheet kosong.
 * @returns {Object} Objek data dashboard default.
 */
function getDefaultDashboardData() {
  Logger.log('Returning default dashboard data due to empty sheets or initial condition.');
  return {
    summary: {
      totalGrowth: 0,
      overallTargetPercentage: 0,
      totalInflow: 0,
      totalOutflow: 0,
      totalGaji: 0,
      totalKomisi: 0,
      activeMarketers: 0,
      totalProspect: 0,
      totalClosing: 0,
      totalTarget: 0
    },
    performanceData: [],
    gajiKomisiData: [],
    top5Kinerja: { labels: [], values: [] },
    bottom5Kinerja: { labels: [], values: [] },
    top5Growth: { labels: [], values: [] },
    negativeGrowth: [],
    growthVsTarget: { labels: [], growth: [], target: [] }
  };
}

/**
 * Helper to sum inflow from marketingPerformance object.
 * @param {Object} marketingPerformance
 * @returns {number}
 */
function getSumOfInflow(marketingPerformance) {
  let sum = 0;
  for (const name in marketingPerformance) {
    sum += marketingPerformance[name].inflow;
  }
  return sum;
}

/**
 * Helper to sum outflow from marketingPerformance object.
 * @param {Object} marketingPerformance
 * @returns {number}
 */
function getSumOfOutflow(marketingPerformance) {
  let sum = 0;
  for (const name in marketingPerformance) {
    sum += marketingPerformance[name].outflow;
  }
  return sum;
}