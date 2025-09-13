// =============================================================================
// KONSTANTA GLOBAL - SESUAIKAN DENGAN ID DAN NAMA SHEET ANDA
// =============================================================================
const SPREADSHEET_ID = '1eZxURW6yArfKbksTBQ_3fAKfcwS4m9drArefsuiYRdU';
const SPREADSHEET_ABSENSI_ID = '1MwDFnkUBDbAhSKqW4Y_dRTuivZdX-TCqmfviC4xcJN0';

const DATA_MASTER_SHEET_NAME = 'MASTER LENDING'; // Nama sheet master data pinjaman
const MASTER_KARYAWAN_SHEET_NAME = 'MASTER KARYAWAN'; // Nama sheet master data karyawan/marketing
const ABSENSI_SHEET_NAME = 'Form Responses 1'; // Nama sheet Absensi (sesuai image_921002.jpg)

// =============================================================================
// WEB APP ENTRY POINT
// =============================================================================
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard Monitoring - Glory Victory Anandita');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

// =============================================================================
// FUNCTION UNTUK GET MARKETING AKTIF
// =============================================================================

/**
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

// =============================================================================
// FUNGSI UTAMA DASHBOARD
// =============================================================================

/**
 * @param {number} month Bulan yang dipilih (1-12).
 * @param {number} year Tahun yang dipilih.
 * @param {string} selectedMarketingParam Nama marketing yang dipilih, atau "All".
 * @param {string} periodType Tipe periode ('monthly' atau 'weekly').
 * @param {number} weekNumber Nomor minggu (jika periodType 'weekly').
 * @param {number} weeklyYear Tahun untuk nomor minggu (jika periodType 'weekly').
 * @returns {Object} Objek berisi ringkasan, data performa marketing, data gaji/komisi,
 * data performa recruiter, data grafik, absensi, dan detail pinjaman.
 */

function getDashboardData(month, year, selectedMarketingParam, periodType, weekNumber, weeklyYear) {
  Logger.log(`[START] getDashboardData with params: month=${month}, year=${year}, selectedMarketingParam=${selectedMarketingParam}, periodType=${periodType}, weekNumber=${weekNumber}, weeklyYear=${weeklyYear}`);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const lendingSheet = ss.getSheetByName(DATA_MASTER_SHEET_NAME);
  const karyawanSheet = ss.getSheetByName(MASTER_KARYAWAN_SHEET_NAME);

  // Memastikan sheet yang dibutuhkan ada
  if (!lendingSheet || !karyawanSheet) {
    Logger.log("Spreadsheet 'MASTER LENDING' or 'MASTER KARYAWAN' not found. Returning default data.");
    return getDefaultDashboardData();
  }

  // Mengambil semua data dari sheet
  const lendingRows = lendingSheet.getDataRange().getDisplayValues();
  const karyawanRows = karyawanSheet.getDataRange().getDisplayValues();

  // Memastikan sheet tidak kosong (minimal ada header dan 1 baris data)
  if (lendingRows.length <= 1 || karyawanRows.length <= 1) {
    Logger.log('One or both sheets are empty or only contain headers. Returning default data.');
    return getDefaultDashboardData();
  }

  const lendingHeader = lendingRows[0];
  const karyawanHeader = karyawanRows[0];

  // Membuat peta header untuk mempermudah akses kolom berdasarkan nama
  const lendingColMap = createHeaderMap(lendingHeader, DATA_MASTER_SHEET_NAME, [
    "TANGGAL PENCAIRAN", "MARKETING", "STATUS", "PINJAMAN DITERIMA",
    "FEE MARKETING", "PLAFON DIAJUKAN", "CHANNEL INCOME", "RECRUITER", "FEE RECRUITER"
  ]);
  const karyawanColMap = createHeaderMap(karyawanHeader, MASTER_KARYAWAN_SHEET_NAME, [
    "NAMA", "TARGET", "PENAWARAN GAJI", "STATUS"
  ]);

  // Menginisialisasi struktur data untuk performa marketing berdasarkan data karyawan
  const initialMarketingPerformance = initializeMarketingPerformance(karyawanRows, karyawanColMap);
  // Menghitung jumlah marketing aktif (tidak termasuk 'All' jika ada)
  let activeMarketersCount = Object.keys(initialMarketingPerformance).length - 1;
  if (activeMarketersCount < 0) activeMarketersCount = 0;

  // --- 1. Hitung Periode Waktu Saat Ini ---
  // Menentukan tanggal mulai dan akhir untuk periode yang dipilih (bulanan atau mingguan)
  let currentStartDate, currentEndDate;
  if (periodType === 'monthly') {
    currentStartDate = new Date(year, month - 1, 1);
    currentEndDate = new Date(year, month, 0); // Hari terakhir bulan sekarang
  } else { // weekly
    const date = new Date(weeklyYear, 0, 1 + (weekNumber - 1) * 7);
    const day = date.getDay(); // 0: Minggu, 1: Senin, dst.
    const diff = date.getDate() - day + (day === 0 ? -6 : 1); // Mendapatkan tanggal Senin dari minggu tersebut
    currentStartDate = new Date(date.setDate(diff));
    currentEndDate = new Date(currentStartDate);
    currentEndDate.setDate(currentStartDate.getDate() + 6); // Mendapatkan tanggal Minggu dari minggu tersebut
  }
  currentStartDate.setHours(0, 0, 0, 0); // Atur waktu ke awal hari
  currentEndDate.setHours(23, 59, 59, 999); // Atur waktu ke akhir hari
  Logger.log(`Current Period: ${currentStartDate.toDateString()} - ${currentEndDate.toDateString()}`);

  // --- 2. Hitung Periode Waktu Sebelumnya ---
  // Menentukan tanggal mulai dan akhir untuk periode sebelumnya (untuk perhitungan growth)
  let prevMonth, prevYear, prevWeekNumber, prevWeeklyYear;
  let prevStartDate, prevEndDate;

  if (periodType === 'monthly') {
    prevMonth = month - 1;
    prevYear = year;
    if (prevMonth === 0) { // Jika bulan saat ini Januari, periode sebelumnya adalah Desember tahun lalu
      prevMonth = 12;
      prevYear = year - 1;
    }
    prevStartDate = new Date(prevYear, prevMonth - 1, 1);
    prevEndDate = new Date(prevYear, prevMonth, 0);
  } else { // weekly
    // Untuk minggu sebelumnya, cukup kurangi 7 hari dari startDate periode saat ini
    prevStartDate = new Date(currentStartDate);
    prevStartDate.setDate(currentStartDate.getDate() - 7);
    prevEndDate = new Date(prevStartDate);
    prevEndDate.setDate(prevStartDate.getDate() + 6);

    // Update weekNumber dan weeklyYear untuk logging/debugging
    prevWeekNumber = Utilities.formatDate(prevStartDate, Session.getScriptTimeZone(), 'w');
    prevWeeklyYear = prevStartDate.getFullYear();
  }
  prevStartDate.setHours(0, 0, 0, 0);
  prevEndDate.setHours(23, 59, 59, 999);
  Logger.log(`Previous Period: ${prevStartDate.toDateString()} - ${prevEndDate.toDateString()}`);

  // --- 3. Proses Data untuk Periode Saat Ini (current data) ---
  // Filter data pinjaman berdasarkan rentang tanggal saat ini dan parameter marketing/periode
  const { filteredLendingRows, totalProspectCount, totalClosingCount, marketingSpecificCounts } = filterLendingData(
    lendingRows, lendingColMap, month, year, selectedMarketingParam, periodType, weekNumber, weeklyYear, currentStartDate, currentEndDate
  );

  // Inisialisasi objek untuk mengakumulasi performa marketing dan recruiter untuk periode saat ini
  let marketingPerformance = JSON.parse(JSON.stringify(initialMarketingPerformance)); // Clone agar tidak memodifikasi initial
  let recruiterPerformance = {}; // BARU: Inisialisasi objek untuk menyimpan data komisi recruiter

  // Agregasi data pinjaman yang sudah difilter ke dalam objek performa marketing dan recruiter
  aggregateLendingData(filteredLendingRows, lendingColMap, marketingPerformance, recruiterPerformance);

  // Memperbarui total prospek per marketing dari hasil filter
  for (const marketerName in marketingSpecificCounts) {
    if (marketingPerformance[marketerName]) {
      marketingPerformance[marketerName].totalProspect = marketingSpecificCounts[marketerName].totalProspect || 0;
    }
  }

  // Menghitung total inflow, outflow, dan calculated income untuk periode saat ini
  const currentPeriodInflow = getSumOfInflow(marketingPerformance);
  const currentPeriodOutflow = getSumOfOutflow(marketingPerformance); // Mungkin tidak digunakan di summary, tapi relevan untuk Gross Profit
  const currentPeriodNewCalculatedIncome = getSumOfNewCalculatedIncome(marketingPerformance);

  // --- 4. Proses Data untuk Periode Sebelumnya (previous data) ---
  // Filter data pinjaman untuk periode sebelumnya untuk menghitung pertumbuhan (growth)
  const { filteredLendingRows: prevFilteredLendingRows } = filterLendingData(
    lendingRows, lendingColMap, prevMonth, prevYear, selectedMarketingParam, periodType, prevWeekNumber, prevWeeklyYear, prevStartDate, prevEndDate
  );

  let previousPeriodInflow = 0;
  // Asumsi inflow adalah PINJAMAN DITERIMA untuk status 'approve' atau 'claimed'
  const validClosingStatuses = ['simulasi', 'approve', 'claimed'];
  const pinjamanDiterimaColIdx = lendingColMap['PLAFON DIAJUKAN'];
  // const pinjamanDiterimaColIdx = lendingColMap['PINJAMAN DITERIMA'];
  const statusColIdx = lendingColMap['STATUS'];

  for (const row of prevFilteredLendingRows) {
    const pinjaman = parseNumber(row[pinjamanDiterimaColIdx]);
    const status = (row[statusColIdx] || '').toString().trim().toLowerCase();
    if (validClosingStatuses.includes(status)) {
      previousPeriodInflow += pinjaman;
    }
  }
  Logger.log(`Previous Period Inflow: ${previousPeriodInflow}`);

  // --- Hitung Total Overall Target ---
  // Menjumlahkan target dari semua marketing aktif (tidak termasuk 'All')
  let totalOverallTarget = 0;
  for (const marketerName in initialMarketingPerformance) {
    if (marketerName !== 'All') {
      totalOverallTarget += initialMarketingPerformance[marketerName].target;
    }
  }

  // Menghitung gaji dan komisi untuk marketing berdasarkan performa periode saat ini
  const { performanceData, gajiKomisiData, totalGajiKeseluruhan, totalKomisiKeseluruhan } = calculateGajiKomisi(
    marketingPerformance
  );

  // --- PANGGIL FUNGSI BARU UNTUK POTENSI PENDAPATAN DI SINI ---
  const potensiPendapatanData = getPotensiPendapatanData(
    lendingRows,
    lendingColMap,
    initialMarketingPerformance,
    currentStartDate, // Gunakan currentStartDate dan currentEndDate
    currentEndDate,
    selectedMarketingParam // Pastikan filter marketing juga diterapkan
  );

  // Menyiapkan data untuk grafik kinerja marketing
  const chartData = prepareChartData(performanceData, marketingPerformance);

  // --- Bagian untuk Absensi Data ---
  // Mengambil data absensi untuk periode saat ini
  const absensiData = getAbsensiData(currentStartDate, currentEndDate); // Data summary/aggregated
  const absensiDailyData = getAbsensiDailyData(currentStartDate, currentEndDate, 'All'); // Data harian, 'All' berarti semua marketing

  // --- Perhitungan Summary Akhir ---
  // Menghitung closing rate, gross profit, net profit, dan total growth
  let closingRate = 0;
  if (totalProspectCount > 0) {
    closingRate = (totalClosingCount / totalProspectCount) * 100;
  }

  let totalKomisiRecruiter = 0;
  for (const recruiterName in recruiterPerformance) {
    if (recruiterPerformance.hasOwnProperty(recruiterName)) {
      totalKomisiRecruiter += recruiterPerformance[recruiterName].totalKomisi || 0;
    }
  }

  const grossProfit = currentPeriodNewCalculatedIncome;
  const netProfit = grossProfit - totalGajiKeseluruhan - totalKomisiRecruiter;
  // const netProfit = grossProfit - totalGajiKeseluruhan - totalKomisiKeseluruhan - totalKomisiRecruiter;
  const totalGrowth = currentPeriodInflow - previousPeriodInflow; // Perbandingan inflow saat ini dengan sebelumnya
  Logger.log(`Current Inflow: ${currentPeriodInflow}, Previous Inflow: ${previousPeriodInflow}, Total Growth: ${totalGrowth}`);


  // Objek ringkasan yang akan ditampilkan di kartu dashboard
  const summary = {
    totalGrowth: totalGrowth,
    overallTargetPercentage: totalOverallTarget > 0 ? (currentPeriodInflow / totalOverallTarget) * 100 : 0,
    totalInflow: currentPeriodInflow,
    totalGaji: totalGajiKeseluruhan,
    totalKomisi: totalKomisiKeseluruhan + totalKomisiRecruiter,
    activeMarketers: activeMarketersCount,
    totalProspect: totalProspectCount,
    totalClosing: totalClosingCount,
    totalTarget: totalOverallTarget,
    closingRate: closingRate,
    grossProfit: grossProfit,
    netProfit: netProfit,
    newCalculatedIncome: currentPeriodNewCalculatedIncome
  };

  // Mendapatkan data detail pinjaman untuk tampilan tabel di frontend
  const detailPinjamanData = getDetailPinjamanData(month, year, selectedMarketingParam, periodType, weekNumber, weeklyYear);

  // Objek hasil akhir yang akan dikembalikan ke frontend
  const result = {
    summary: summary,
    performanceData: performanceData,
    gajiKomisiData: gajiKomisiData,
    recruiterPerformance: recruiterPerformance,
    top5Kinerja: chartData.top5Kinerja,
    bottom5Kinerja: chartData.bottom5Kinerja,
    top5Growth: chartData.top5Growth,
    negativeGrowth: chartData.negativeGrowth,
    growthVsTarget: chartData.growthVsTarget,
    absensiData: absensiData,
    absensiDailyData: absensiDailyData,
    detailPinjamanData: detailPinjamanData,
    reportStartDate: currentStartDate.getTime(),
    reportEndDate: currentEndDate.getTime(),
    potensiPendapatanData: potensiPendapatanData
  };

  // Logging hasil untuk debugging
  Logger.log("[END] getDashboardData. Returning object (excerpt summary): " + JSON.stringify(result.summary));
  Logger.log("[END] getDashboardData. Returning object (excerpt performanceData): " + JSON.stringify(result.performanceData.slice(0, 5)));
  Logger.log("[END] getDashboardData. Returning object (excerpt gajiKomisiData): " + JSON.stringify(result.gajiKomisiData.slice(0, 5)));
  Logger.log("[END] getDashboardData. Returning object (excerpt recruiterPerformance): " + JSON.stringify(result.recruiterPerformance)); // Log data recruiter
  Logger.log("[END] getDashboardData. Returning object (excerpt absensiData): " + JSON.stringify(result.absensiData.slice(0, 5)));
  Logger.log("[END] getDashboardData. Returning object (excerpt absensiDailyData): " + JSON.stringify(result.absensiDailyData.slice(0, 5)));
  Logger.log("[END] getDashboardData. Returning object (excerpt detailPinjamanData): " + JSON.stringify(result.detailPinjamanData.slice(0, 5)));

  return result;
}

// =============================================================================
// FUNGSI PEMBANTU
// =============================================================================

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
        totalProspect: 0,
        totalClosing: 0,
        target: target,
        gajiKontrakAwal: gajiKontrak,
        totalGaji: 0,
        totalKomisi: 0,
        newCalculatedIncome: 0
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

  const colLendingTanggalPencairan = lendingColMap['TANGGAL PENGAJUAN'];
  const colLendingMarketingName = lendingColMap['MARKETING'];
  const colLendingStatus = lendingColMap['STATUS'];

  const validClosingStatuses = ['simulasi', 'approve', 'claimed'];
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
 * Mengagregasi data pinjaman yang sudah difilter ke objek performa marketing dan recruiter.
 * Ini menghitung inflow, outflow, fee marketing, income baru, dan komisi recruiter.
 * @param {Array<Array>} filteredLendingRows - Data pinjaman yang sudah difilter.
 * @param {Object} lendingColMap - Peta indeks kolom untuk MASTER LENDING.
 * @param {Object} marketingPerformance - Objek untuk mengakumulasi data marketing (dimodifikasi secara referensi).
 * @param {Object} recruiterPerformance - Objek untuk mengakumulasi data recruiter (dimodifikasi secara referensi).
 */
function aggregateLendingData(filteredLendingRows, lendingColMap, marketingPerformance, recruiterPerformance) {
  const colLendingMarketingName = lendingColMap['MARKETING'];
  const colLendingStatus = lendingColMap['STATUS'];
  const colLendingPinjamanDiterima = lendingColMap['PLAFON DIAJUKAN'];
  // const colLendingPinjamanDiterima = lendingColMap['PINJAMAN DITERIMA'];
  const colLendingFeeMarketing = lendingColMap['FEE MARKETING'];
  const colLendingPlafonDiajukan = lendingColMap['PLAFON DIAJUKAN'];
  const colLendingIncome = lendingColMap['CHANNEL INCOME'];
  const colLendingRecruiterName = lendingColMap['RECRUITER'];
  const colLendingFeeRecruiter = lendingColMap['FEE RECRUITER'];

  const validClosingStatuses = ['simulasi', 'approve', 'claimed'];

  Logger.log(`Starting aggregateLendingData with ${filteredLendingRows.length} filtered rows.`);

  for (const row of filteredLendingRows) {
    const marketingNameInRow = (row[colLendingMarketingName] || '').toString().trim();
    const statusInRow = (row[colLendingStatus] || '').toString().trim().toLowerCase();
    const pinjamanDiterima = parseNumber(row[colLendingPinjamanDiterima]);
    const plafonDiajukan = parseNumber(row[colLendingPlafonDiajukan]);

    // Perhitungan New Calculated Income (berdasarkan Kolom Z: INCOME)
    const incomePercentageRaw = row[colLendingIncome];
    let incomePercentage = 0;

    if (typeof incomePercentageRaw === 'number') {
      // Kalau sudah desimal, misal 0.031 → langsung pakai
      incomePercentage = incomePercentageRaw < 1 ? incomePercentageRaw : incomePercentageRaw / 100;
    } else if (typeof incomePercentageRaw === 'string') {
      const cleaned = incomePercentageRaw.replace('%', '').trim();
      incomePercentage = parseFloat(cleaned) / 100;
    }


    const newCalculatedIncome = pinjamanDiterima * incomePercentage;

    // Perbarui marketingPerformance untuk marketingNameInRow yang ditemukan
    if (marketingPerformance[marketingNameInRow]) {
      // Akumulasi newCalculatedIncome ke marketing individual
      marketingPerformance[marketingNameInRow].newCalculatedIncome += newCalculatedIncome;
      Logger.log(`Aggregating newCalculatedIncome for ${marketingNameInRow}: +${newCalculatedIncome}. Current total: ${marketingPerformance[marketingNameInRow].newCalculatedIncome}`);

      // Hanya akumulasi inflow, feeMarketingCair, dan totalClosing untuk status yang valid
      if (validClosingStatuses.includes(statusInRow)) {
        const feeMarketingRaw = row[colLendingFeeMarketing];
        const feePercentage = parseNumber(feeMarketingRaw, true); // Pastikan parseNumber bisa handle true untuk persentase jika perlu

        const feeMarketingAmount = pinjamanDiterima * (feePercentage / 100);

        marketingPerformance[marketingNameInRow].inflow += pinjamanDiterima;
        marketingPerformance[marketingNameInRow].feeMarketingCair += feeMarketingAmount;
        marketingPerformance[marketingNameInRow].totalClosing++;
        Logger.log(`Aggregating for ${marketingNameInRow} (Closing): Inflow=${pinjamanDiterima}, Fee=${feeMarketingAmount}. Current Inflow for ${marketingNameInRow}: ${marketingPerformance[marketingNameInRow].inflow}`);
      }

    } else {
      Logger.log(`Marketing name "${marketingNameInRow}" from lending data not found in MASTER KARYAWAN. Skipping performance aggregation for row: ${JSON.stringify(row)}`);
    }

    // Akumulasi juga ke entri 'All' di marketingPerformance (jika sudah diinisialisasi)
    if (marketingPerformance['All']) {
      marketingPerformance['All'].newCalculatedIncome += newCalculatedIncome;
      if (validClosingStatuses.includes(statusInRow)) {
        marketingPerformance['All'].inflow += pinjamanDiterima;
        // marketingPerformance['All'].feeMarketingCair += feeMarketingAmount; // Jika Anda melacak total fee marketing cair
        marketingPerformance['All'].totalClosing++;
      }
    }

    // Akumulasi data untuk 'All' marketing
    if (marketingPerformance['All']) {
      marketingPerformance['All'].newCalculatedIncome += newCalculatedIncome;
      if (validClosingStatuses.includes(statusInRow)) {
        marketingPerformance['All'].inflow += pinjamanDiterima;
        marketingPerformance['All'].totalClosing++;
        // marketingPerformance['All'].outflow += plafonDiajukan; // Jika ingin mengakumulasi total outflow untuk 'All'
      }
    }

    // --- Logika Perhitungan Komisi Recruiter ---
    const recruiterNameInRow = (row[colLendingRecruiterName] || '').toString().trim();
    const feeRecruiterRaw = row[colLendingFeeRecruiter];

    let feeRecruiterPercentage = 0;
    if (typeof feeRecruiterRaw === 'number') {
      feeRecruiterPercentage = feeRecruiterRaw;
    } else {
      if (feeRecruiterPercentage > 1 && feeRecruiterPercentage <= 100) {
        feeRecruiterPercentage = feeRecruiterPercentage / 100;
      } else if (feeRecruiterRaw.toString().includes('%')) {
        // feeRecruiterPercentage = convertPercentToDecimal(feeRecruiterRaw.toString());
        feeRecruiterPercentage = parseNumber(feeRecruiterRaw.toString().replace('%', '')) / 10000;
      }
    }

    // Komisi Recruiter hanya dihitung jika ada nama recruiter DAN status pinjaman valid closing
    if (recruiterNameInRow && validClosingStatuses.includes(statusInRow)) {
      const recruiterCommission = pinjamanDiterima * feeRecruiterPercentage;

      // Inisialisasi jika recruiter belum ada di objek recruiterPerformance
      if (!recruiterPerformance[recruiterNameInRow]) {
        recruiterPerformance[recruiterNameInRow] = { totalKomisi: 0 };
      }
      recruiterPerformance[recruiterNameInRow].totalKomisi += recruiterCommission;
      // Logger.log(`Aggregating recruiter commission for ${recruiterNameInRow}: +${recruiterCommission}. Current total: ${recruiterPerformance[recruiterNameInRow].totalKomisi}`);
    }
  }
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
    const pencapaianInflow = data.inflow; // Ini adalah 'Present'
    const kinerja = target > 0 ? (pencapaianInflow / target) : 0; // Ini akan tetap sebagai desimal (0.85, 1.0, 1.25)

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
    // data.feeMarketingCair sudah dihitung di aggregateLendingData
    let komisiFeeKreditCair = data.feeMarketingCair;
    Logger.log(`Marketing ${name}: Komisi Fee Kredit Cair = ${komisiFeeKreditCair.toFixed(2)}`);

    // Rule 2: Reward Tambahan 0.25% (diambil dari FEE MARKETING) jika achievement > 100%
    // Reward berlaku hanya untuk closingan setelah target
    let rewardTambahan = 0;
    if (kinerja > 1.0) {
      const excessInflow = pencapaianInflow - target; // Inflow yang melebihi target
      if (excessInflow > 0) {
        // Asumsi "diambil dari kolom FEE MARKETING" berarti menggunakan persentase yang sama
        // 0.25% dari kolom FEE MARKETING, jadi 0.25% dari excessInflow
        // Perhatikan bahwa di aggregateLendingData, feePercentage sudah dibagi 100 (misal 0.0025)
        // Jadi di sini kita menggunakan 0.0025 secara langsung (0.25%)
        rewardTambahan = excessInflow * 0.0025; // 0.25% dari sisa inflow setelah target
        Logger.log(`Marketing ${name}: Kinerja ${kinerja.toFixed(2)} (>1.0). Excess Inflow=${excessInflow.toFixed(2)}. Reward tambahan = ${rewardTambahan.toFixed(2)} (0.25% dari excess inflow)`);
      }
    } else {
      Logger.log(`Marketing ${name}: Kinerja ${kinerja.toFixed(2)} (<=1.0). Tidak ada reward tambahan.`);
    }

    let totalKomisiPerMarketing = komisiFeeKreditCair + rewardTambahan;
    data.totalKomisi = totalKomisiPerMarketing;
    totalKomisiKeseluruhan += totalKomisiPerMarketing;

    performanceData.push({
      name: name,
      performance: kinerja * 100, // Tetap sebagai desimal (0.85, 1.0, 1.25) lalu dikali10
      commission: totalKomisiPerMarketing,
      inflow: pencapaianInflow,
      target: target,
      totalProspect: data.totalProspect || 0,
      totalClosing: data.totalClosing || 0,
    });

    gajiKomisiData.push({
      name: name,
      gajiKontrakAwal: data.gajiKontrakAwal,
      kinerja: kinerja, // Tetap sebagai desimal
      gajiFinal: gajiFinal,
      feeMarketingCair: komisiFeeKreditCair,
      rewardTambahan: rewardTambahan,
      totalKomisi: totalKomisiPerMarketing,
      totalDibayarkan: gajiFinal + totalKomisiPerMarketing
    });
  }
  Logger.log(`[END] calculateGajiKomisi. Final totalGajiKeseluruhan: ${totalGajiKeseluruhan}, totalKomisiKeseluruhan: ${totalKomisiKeseluruhan}`);
  Logger.log(performanceData);
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
    growthVsTarget: growthVsTargetData,
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

function convertPercentToDecimal(input) {
  if (typeof input === 'string' && input.includes('%')) {
    const numberOnly = parseFloat(input.replace('%', '').trim());
    if (!isNaN(numberOnly)) {
      return numberOnly / 100;
    }
  }
  return null; // Jika input tidak valid
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
    growthVsTarget: { labels: [], growth: [], target: [] },
    absensiData: [],
    absensiDailyData: []
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
    sum += parseNumber(marketingPerformance[name].inflow);
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
    sum += parseNumber(marketingPerformance[name].outflow);
  }
  return sum;
}

/**
 * @param {number} currentMonth Bulan saat ini (1-12).
 * @param {number} currentYear Tahun saat ini.
 * @param {string} periodType Tipe periode ('monthly' atau 'weekly').
 * @param {number} currentWeekNumber Nomor minggu saat ini (1-52/53).
 * @param {number} currentWeeklyYear Tahun minggu saat ini.
 * @returns {number} Total inflow dari periode sebelumnya.
 */

function getPreviousPeriodInflow(currentMonth, currentYear, periodType, currentWeekNumber, currentWeeklyYear) {
  let prevStartDate, prevEndDate;

  if (periodType === 'monthly') {
    let prevMonth = currentMonth - 1;
    let prevYear = currentYear;
    if (prevMonth === 0) { // Jika bulan saat ini Januari (1), periode sebelumnya adalah Desember tahun lalu
      prevMonth = 12;
      prevYear--;
    }
    prevStartDate = new Date(prevYear, prevMonth - 1, 1);
    prevEndDate = new Date(prevYear, prevMonth, 0); // Hari terakhir bulan sebelumnya
    prevEndDate.setHours(23, 59, 59, 999);
  } else { // weekly
    // Hitung tanggal mulai minggu sebelumnya
    const currentWeekStart = new Date(currentWeeklyYear, 0, 1 + (currentWeekNumber - 1) * 7);
    const day = currentWeekStart.getDay();
    const diff = currentWeekStart.getDate() - day + (day === 0 ? -6 : 1); // Senin minggu ini
    currentWeekStart.setDate(diff); // Set ke Senin minggu ini

    prevStartDate = new Date(currentWeekStart);
    prevStartDate.setDate(currentWeekStart.getDate() - 7); // Kurangi 7 hari untuk Senin minggu lalu
    prevStartDate.setHours(0, 0, 0, 0);

    prevEndDate = new Date(prevStartDate);
    prevEndDate.setDate(prevStartDate.getDate() + 6); // Minggu lalu berakhir 6 hari setelah prevStartDate
    prevEndDate.setHours(23, 59, 59, 999);
  }

  Logger.log(`Calculating Previous Period Inflow for: ${prevStartDate.toLocaleDateString()} to ${prevEndDate.toLocaleDateString()}`);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const lendingSheet = ss.getSheetByName(DATA_MASTER_SHEET_NAME);

  if (!lendingSheet) {
    Logger.log(`Sheet '${DATA_MASTER_SHEET_NAME}' not found.`);
    return 0;
  }

  const lendingRows = lendingSheet.getDataRange().getDisplayValues();
  if (lendingRows.length <= 1) {
    return 0;
  }

  const lendingHeader = lendingRows[0];
  const lendingColMap = createHeaderMap(lendingHeader, DATA_MASTER_SHEET_NAME, ["TANGGAL PENCAIRAN", "PLAFON DIAJUKAN"]);
  // const lendingColMap = createHeaderMap(lendingHeader, DATA_MASTER_SHEET_NAME, ["TANGGAL PENCAIRAN", "PINJAMAN DITERIMA"]);

  const tanggalPencairanCol = lendingColMap["TANGGAL PENCAIRAN"];
  const pinjamanDiterimaCol = lendingColMap["PLAFON DIAJUKAN"];
  // const pinjamanDiterimaCol = lendingColMap["PINJAMAN DITERIMA"];

  if (tanggalPencairanCol === -1 || pinjamanDiterimaCol === -1) {
    Logger.log("Kolom 'TANGGAL PENCAIRAN' atau 'PINJAMAN DITERIMA' tidak ditemukan di sheet lending.");
    return 0;
  }

  let previousInflow = 0;

  for (let i = 1; i < lendingRows.length; i++) {
    const row = lendingRows[i];
    const tanggalPencairan = new Date(row[tanggalPencairanCol]);
    const pinjamanDiterima = parseFloat(String(row[pinjamanDiterimaCol]).replace(/\./g, '').replace(/,/g, '.')); // Handle IDR format

    if (!isNaN(tanggalPencairan.getTime()) && tanggalPencairan >= prevStartDate && tanggalPencairan <= prevEndDate) {
      if (!isNaN(pinjamanDiterima)) {
        previousInflow += pinjamanDiterima;
      }
    }
  }
  Logger.log(`Previous Period Inflow: ${previousInflow}`);
  return previousInflow;
}

function getAbsensiData(startDate, endDate) {
  Logger.log(`Processing Absensi Data for ${startDate.toLocaleDateString()} to ${endDate.toLocaleDateString()}`);

  const ssAbsensi = SpreadsheetApp.openById(SPREADSHEET_ABSENSI_ID);
  const sheetAbsensi = ssAbsensi.getSheetByName(ABSENSI_SHEET_NAME);

  if (!sheetAbsensi) {
    Logger.log(`Sheet '${ABSENSI_SHEET_NAME}' tidak ditemukan di Spreadsheet Absensi.`);
    return [];
  }

  const dataRange = sheetAbsensi.getDataRange();
  const values = dataRange.getValues();

  if (values.length < 2) {
    Logger.log("Tidak ada data absensi yang ditemukan.");
    return [];
  }

  const headers = values[0];
  const timestampCol = headers.indexOf('Timestamp');
  const namaMarketingCol = headers.indexOf('NAMA MARKETING');

  if (timestampCol === -1 || namaMarketingCol === -1) {
    Logger.log("Kolom 'Timestamp' atau 'NAMA MARKETING' tidak ditemukan.");
    return [];
  }

  const absensiByMarketing = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const timestamp = row[timestampCol];
    const marketingName = row[namaMarketingCol];

    // Pastikan timestamp adalah objek Date dan dalam rentang tanggal
    if (timestamp instanceof Date && timestamp >= startDate && timestamp <= endDate) {
      // --- PERUBAHAN DI SINI: HAPUS KONDISI FILTER selectedMarketing ---
      // Data absensi akan dihitung untuk setiap marketing yang ditemukan dalam rentang tanggal.
      if (!absensiByMarketing[marketingName]) {
        absensiByMarketing[marketingName] = {
          totalKunjungan: 0,
          hariKerja: new Set()
        };
      }

      absensiByMarketing[marketingName].totalKunjungan++;

      const dateKey = Utilities.formatDate(timestamp, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
      absensiByMarketing[marketingName].hariKerja.add(dateKey);
    }
  }

  const result = [];
  for (const name in absensiByMarketing) {
    const data = absensiByMarketing[name];
    const jumlahHariKerja = data.hariKerja.size;
    const totalKunjungan = data.totalKunjungan;
    const rataRataKunjungan = jumlahHariKerja > 0 ? (totalKunjungan / jumlahHariKerja).toFixed(2) : 0;
    const presentase = (jumlahHariKerja / 25) * 100;

    result.push({
      name: name,
      jumlahHariKerja: jumlahHariKerja,
      totalKunjungan: totalKunjungan,
      rataRataKunjungan: parseFloat(rataRataKunjungan),
      presentase: presentase,
    });
  }

  result.sort((a, b) => a.name.localeCompare(b.name));

  Logger.log("Absensi Data Result: " + JSON.stringify(result));
  return result;
}

/**
 * Mengambil dan memproses data absensi harian secara detail.
 * Mengembalikan semua entri dalam rentang tanggal, tidak mengelompokkan per marketing.
 *
 * @param {Date} startDate Tanggal awal periode.
 * @param {Date} endDate Tanggal akhir periode.
 * @returns {Array<Object>} Array objek data absensi harian yang terperinci.
 */
function getAbsensiDailyData(startDate, endDate, selectedMarketing) {
  Logger.log(`Processing Absensi Daily Data from ${startDate.toLocaleDateString()} to ${endDate.toLocaleDateString()}`);

  const ssAbsensi = SpreadsheetApp.openById(SPREADSHEET_ABSENSI_ID);
  const sheetAbsensi = ssAbsensi.getSheetByName(ABSENSI_SHEET_NAME);

  if (!sheetAbsensi) {
    Logger.log(`Sheet '${ABSENSI_SHEET_NAME}' tidak ditemukan di Spreadsheet Absensi.`);
    return [];
  }

  const dataRange = sheetAbsensi.getDataRange();
  const values = dataRange.getValues();

  if (values.length < 2) { // Baris header + minimal 1 baris data
    Logger.log("Tidak ada data absensi harian yang ditemukan.");
    return [];
  }

  const headers = values[0];
  const timestampCol = headers.indexOf('Timestamp');
  const namaMarketingCol = headers.indexOf('NAMA MARKETING');
  const namaAnggotaCalonAnggotaCol = headers.indexOf('NAMA ANGGOTA / CALON ANGGOTA'); // Pastikan ini PERSIS sama
  const alamatCol = headers.indexOf('ALAMAT');
  const fuViaCol = headers.indexOf('FOLLOW UP VIA');
  const keteranganCol = headers.indexOf('KETERANGAN');
  const hasilKunjunganCol = headers.indexOf('HASIL KUNJUNGAN'); // Pastikan ini PERSIS sama

  if (timestampCol === -1 || namaMarketingCol === -1 ||
    namaAnggotaCalonAnggotaCol === -1 || alamatCol === -1 ||
    fuViaCol === -1 || keteranganCol === -1 || hasilKunjunganCol === -1) {
    Logger.log("Satu atau lebih kolom yang diperlukan untuk laporan absensi harian tidak ditemukan.");
    Logger.log(`Headers found: ${headers.join(', ')}`);
    Logger.log(`Missing checks: Timestamp=${timestampCol === -1}, NAMA MARKETING=${namaMarketingCol === -1}, NAMA ANGGOTA / CALON ANGGOTA=${namaAnggotaCalonAnggotaCol === -1}, ALAMAT=${alamatCol === -1}, FOLLOW UP VIA=${fuViaCol === -1}, KETERANGAN=${keteranganCol === -1}, HASIL KUNJUNGAN=${hasilKunjunganCol === -1}`);
    return [];
  }

  const result = []; // Inisialisasi array untuk menyimpan semua baris data

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const timestamp = row[timestampCol];
    const marketingName = row[namaMarketingCol];

    // Pastikan timestamp adalah objek Date dan dalam rentang tanggal
    if (timestamp instanceof Date && timestamp >= startDate && timestamp <= endDate) {
      // Filter berdasarkan marketing yang dipilih jika bukan 'All'
      if (selectedMarketing === 'All' || marketingName === selectedMarketing) {
        result.push({
          name: marketingName,
          tanggal: Utilities.formatDate(timestamp, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'dd-MM-yyyy HH:mm'),
          namaAnggotaCalonAnggota: row[namaAnggotaCalonAnggotaCol] || '',
          alamat: row[alamatCol] || '',
          fuVia: row[fuViaCol] || '',
          keterangan: row[keteranganCol] || '',
          hasilKunjungan: row[hasilKunjunganCol] || ''
        });
      }
    }
  }

  // Mengurutkan berdasarkan tanggal, dari yang paling awal ke paling akhir
  // Jika ada data dengan tanggal yang sama, bisa diurutkan berdasarkan nama marketing
  result.sort((a, b) => {
    // Konversi string tanggal ke objek Date agar bisa diurutkan
    // Format 'dd-MM-yyyy HH:mm' perlu diubah ke 'YYYY-MM-DDTHH:mm:ss' agar Date() bisa parse dengan benar
    const dateA = new Date(a.tanggal.replace(/(\d{2})-(\d{2})-(\d{4}) (\d{2}):(\d{2})/, '$3-$2-$1T$4:$5:00'));
    const dateB = new Date(b.tanggal.replace(/(\d{2})-(\d{2})-(\d{4}) (\d{2}):(\d{2})/, '$3-$2-$1T$4:$5:00'));

    if (dateA.getTime() !== dateB.getTime()) {
      return dateB.getTime() - dateA.getTime(); // PERUBAHAN DI SINI: Paling baru ke paling lama
    }
    // Jika tanggal sama, urutkan berdasarkan nama marketing secara alfabetis
    return a.name.localeCompare(b.name);
  });

  Logger.log("Absensi Daily Data Result (first 5): " + JSON.stringify(result.slice(0, 5)));
  Logger.log("Absensi Daily Data Result (total items): " + result.length);
  return result;
}

/**
 * Mengambil dan memfilter data detail pinjaman dari MASTER LENDING.
 * @param {number} month Bulan yang dipilih (1-12).
 * @param {number} year Tahun yang dipilih.
 * @param {string} selectedMarketingParam Nama marketing yang dipilih, atau "All".
 * @param {string} periodType Tipe periode ('monthly' atau 'weekly').
 * @param {number} weekNumber Nomor minggu (jika periodType 'weekly').
 * @param {number} weeklyYear Tahun untuk nomor minggu (jika periodType 'weekly').
 * @returns {Array<Object>} Array of pinjaman data objects.
 */
function getDetailPinjamanData(month, year, selectedMarketingParam, periodType, weekNumber, weeklyYear) {
  Logger.log(`[START] getDetailPinjamanData with params: month=${month}, year=${year}, selectedMarketingParam=${selectedMarketingParam}, periodType=${periodType}, weekNumber=${weekNumber}, weeklyYear=${weeklyYear}`);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const lendingSheet = ss.getSheetByName(DATA_MASTER_SHEET_NAME);

  if (!lendingSheet) {
    Logger.log(`Sheet '${DATA_MASTER_SHEET_NAME}' not found for getDetailPinjamanData. Returning empty array.`);
    return [];
  }

  const lendingRows = lendingSheet.getDataRange().getDisplayValues();
  if (lendingRows.length <= 1) {
    Logger.log('MASTER LENDING sheet is empty or only contains headers for getDetailPinjamanData. Returning empty array.');
    return [];
  }

  const lendingHeader = lendingRows[0];
  const requiredCols = [
    "DEBITUR", "TELEPON", "NIK", "NOPEN", "PEKERJAAN", "ALAMAT", "PLAFON DIAJUKAN",
    "PINJAMAN DITERIMA", "TANGGAL PENGAJUAN", "CHANNEL", "MARKETING", "STATUS"
  ];
  const lendingColMap = createHeaderMap(lendingHeader, DATA_MASTER_SHEET_NAME, requiredCols);

  const colDebitur = lendingColMap['DEBITUR'];
  const colTelepon = lendingColMap['TELEPON'];
  const colNIK = lendingColMap['NIK'];
  const colNopen = lendingColMap['NOPEN'];
  const colPekerjaan = lendingColMap['PEKERJAAN'];
  const colAlamat = lendingColMap['ALAMAT'];
  const colPinjamanDiterima = lendingColMap['PLAFON DIAJUKAN'];
  // const colPinjamanDiterima = lendingColMap['PINJAMAN DITERIMA'];
  const colTanggalPengajuan = lendingColMap['TANGGAL PENGAJUAN'];
  const colChannel = lendingColMap['CHANNEL'];
  const colMarketing = lendingColMap['MARKETING'];
  const colStatus = lendingColMap['STATUS'];

  const detailPinjaman = [];

  for (let i = 1; i < lendingRows.length; i++) {
    const row = lendingRows[i];
    const dateCell = row[colTanggalPengajuan];
    const date = (dateCell instanceof Date) ? dateCell : new Date(dateCell);

    if (isNaN(date.getTime())) {
      continue;
    }

    const marketingNameInRow = (row[colMarketing] || '').toString().trim();

    // Check date match
    let isDateMatch = false;
    if (periodType === 'monthly' && month !== null && year !== null) {
      isDateMatch = date.getMonth() + 1 === parseInt(month) && date.getFullYear() === parseInt(year);
    } else if (periodType === 'weekly' && weekNumber !== null && weeklyYear !== null) {
      const rowWeekNumber = getWeekNumber(date);
      // Asumsi weekNumberCalculated adalah typo dan seharusnya rowWeekNumber
      isDateMatch = rowWeekNumber === parseInt(weekNumber) && date.getFullYear() === parseInt(weeklyYear);
    } else {
      isDateMatch = true; // Jika tidak ada filter waktu spesifik, anggap cocok
    }

    // Check marketing match
    const isMarketingMatch = (selectedMarketingParam === 'All' || !selectedMarketingParam || marketingNameInRow === selectedMarketingParam);

    if (isDateMatch && isMarketingMatch) {
      detailPinjaman.push({
        debitur: row[colDebitur] || '',
        telepon: row[colTelepon] || '',
        nik: row[colNIK] || '',
        nopen: row[colNopen] || '',
        pekerjaan: row[colPekerjaan] || '',
        alamat: row[colAlamat] || '',
        pinjamanDiterima: parseNumber(row[colPinjamanDiterima]), // Pastikan diformat sebagai angka
        tanggalPencairan: row[colTanggalPengajuan] || '',
        channel: row[colChannel] || '',
        marketing: marketingNameInRow,
        status: row[colStatus] || ''
      });
    }
  }

  // Sort the data by tanggalPencairan from newest to oldest
  detailPinjaman.sort((a, b) => {
    const dateA = new Date(a.tanggalPencairan.replace(/(\d{2})-(\d{2})-(\d{4}) (\d{2}):(\d{2})/, '$3-$2-$1T$4:$5:00'));
    const dateB = new Date(b.tanggalPencairan.replace(/(\d{2})-(\d{2})-(\d{4}) (\d{2}):(\d{2})/, '$3-$2-$1T$4:$5:00'));
    return dateB.getTime() - dateA.getTime(); // Sort descending (newest first)
  });

  Logger.log(`[END] getDetailPinjamanData. Found ${detailPinjaman.length} items.`);
  return detailPinjaman;
}

function getSumOfNewCalculatedIncome(marketingPerformance) {
  let sum = 0;
  for (const name in marketingPerformance) {
    sum += parseNumber(marketingPerformance[name].newCalculatedIncome);
  }
  return sum;
}

function getSumOfOutflow(marketingPerformance) {
  let sum = 0;
  for (const name in marketingPerformance) {
    sum += parseNumber(marketingPerformance[name].outflow);
  }
  return sum;
}

/**
 * Menghitung potensi pendapatan per marketing berdasarkan TANGGAL PENGAJUAN.
 * Mengembalikan data dalam format yang mirip dengan performanceData.
 * @param {Array<Array>} lendingRows Semua data dari sheet MASTER LENDING.
 * @param {Object} lendingColMap Peta kolom untuk MASTER LENDING.
 * @param {Object} initialMarketingPerformance Data awal marketing (untuk target).
 * @param {Date} startDate Tanggal mulai periode.
 * @param {Date} endDate Tanggal akhir periode.
 * @param {string} selectedMarketingParam Marketing yang dipilih ('All' atau nama spesifik).
 * @returns {Array<Object>} Array objek potensi pendapatan per marketing.
 */
function getPotensiPendapatanData(lendingRows, lendingColMap, initialMarketingPerformance, startDate, endDate, selectedMarketingParam) {
  Logger.log(`[START] getPotensiPendapatanData for period: ${startDate.toDateString()} - ${endDate.toDateString()}`);

  // Clone initial performance object untuk menghindari modifikasi objek asli
  let potensiMarketingPerformance = JSON.parse(JSON.stringify(initialMarketingPerformance));

  const colTanggalPengajuan = lendingColMap['TANGGAL PENGAJUAN'];
  const colPlafonDiajukan = lendingColMap['PLAFON DIAJUKAN'];
  const colMarketingName = lendingColMap['MARKETING'];
  const colStatus = lendingColMap['STATUS'];

  // Status yang dianggap sebagai "potensi" (belum ditolak/dibatalkan secara final)
  const potentialStatuses = ['simulasi', 'verifikasi', 'approve', 'claimed'];

  for (let i = 1; i < lendingRows.length; i++) { // Mulai dari baris 1 (setelah header)
    const row = lendingRows[i];
    const tanggalPengajuan = parseDate(row[colTanggalPengajuan]);
    const marketingNameInRow = (row[colMarketingName] || '').toString().trim();
    const plafonDiajukan = parseNumber(row[colPlafonDiajukan]);
    const statusInRow = (row[colStatus] || '').toString().trim().toLowerCase();

    // Filter berdasarkan tanggal pengajuan
    if (tanggalPengajuan && tanggalPengajuan >= startDate && tanggalPengajuan <= endDate) {
      // Filter berdasarkan marketing yang dipilih (jika bukan 'All')
      const marketingMatch = (selectedMarketingParam === "All" || marketingNameInRow === selectedMarketingParam);

      if (marketingMatch && potentialStatuses.includes(statusInRow)) {
        if (potensiMarketingPerformance[marketingNameInRow]) {
          // Akumulasi plafon diajukan sebagai "inflow" potensial
          potensiMarketingPerformance[marketingNameInRow].inflow += plafonDiajukan;
        } else {
          // Jika marketing tidak ada di initialMarketingPerformance (misal: baru direkrut/data belum update)
          // Anda bisa memilih untuk mengabaikannya atau menambahkannya dengan target 0
          Logger.log(`Marketing "${marketingNameInRow}" not found in initialMarketingPerformance. Skipping potential aggregation for this entry.`);
        }
      }
    }
  }

  // Format data ke dalam array yang mirip dengan performanceData
  const potensiPendapatanData = [];
  for (const name in potensiMarketingPerformance) {
    const data = potensiMarketingPerformance[name];
    const target = data.target;
    const potensiInflow = data.inflow;
    const kinerja = target > 0 ? (potensiInflow / target) : 0;

    potensiPendapatanData.push({
      name: name,
      performance: kinerja * 100, // Konversi ke persentase
      commission: 0, // Tidak ada komisi untuk potensi
      inflow: potensiInflow, // Ini adalah plafon diajukan
      target: target,
      totalProspect: 0, // Tidak relevan di sini, ini adalah potensi closing
      totalClosing: 0 // Tidak relevan di sini
    });
  }

  potensiPendapatanData.sort((a, b) => b.performance - a.performance);
  Logger.log(`[END] getPotensiPendapatanData. Result (first 5): ${JSON.stringify(potensiPendapatanData.slice(0, 5))}`);
  return potensiPendapatanData;
}

/**
 * Helper untuk menjumlahkan 'inflow' (potensi pendapatan) dari array potensiPendapatanData.
 * @param {Array<Object>} potensiPendapatanData
 * @returns {number}
 */
function getSumOfPotensiPendapatan(potensiPendapatanData) {
  let sum = 0;
  for (const item of potensiPendapatanData) {
    sum += item.inflow;
  }
  return sum;
}

/**
 * Mengubah nilai dari spreadsheet (teks, angka, atau objek Date) menjadi objek Date yang valid.
 * @param {*} value Nilai yang akan diparse.
 * @returns {Date|null} Objek Date jika berhasil, null jika gagal.
 */
function parseDate(value) {
  if (value instanceof Date) {
    return value;
  }
  if (typeof value === 'string' || typeof value === 'number') {
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date;
    }
  }
  return null;
}
