/**
 * Fungsi Utama: Mengendalikan permintaan HTTP GET.
 * Menyokong dwi-fungsi:
 * 1. Jika ada parameter ?action=getMasterData -> Pulangkan JSON untuk Netlify.
 * 2. Jika diakses biasa melalui pelayar web -> Paparkan antaramuka Index.html.
 */
function doGet(e) {
  // Semak sama ada panggilan datang daripada Netlify (Fetch JSON)
  var action = (e && e.parameter) ? e.parameter.action : "";
  
  if (action === "getMasterData") {
    try {
      var masterData = getMasterDataFromSheet();
      return ContentService.createTextOutput(JSON.stringify(masterData))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: err.toString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Jika dibuka sebagai Web App biasa di pelayar Google
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Kalkulator Pentauliahan Pusat Bertauliah JPK')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Membaca data daripada Tab 'DaftarNOSS', 'Jajar', dan 'WA'
 * ID Spreadsheet rasmi disimpan secara berpusat.
 */
function getMasterDataFromSheet() {
  var ssId = "1uGiBnoDFnvdqe-9tQ38VvKwCEBGSaUp-kChzCvS6BTI";
  var ss = SpreadsheetApp.openById(ssId);
  
  // ==========================================
  // 1. Baca Tab DaftarNOSS
  // ==========================================
  var nossSheet = ss.getSheetByName("DaftarNOSS");
  var nossValues = nossSheet ? nossSheet.getDataRange().getValues() : [];
  var nossList = [];
  
  for (var i = 1; i < nossValues.length; i++) {
    var code = nossValues[i][0];        // Kolum A: KOD PROGRAM
    var name = nossValues[i][1];        // Kolum B: NAMA PROGRAM
    var statusTahap = nossValues[i][3]; // Kolum D: Status Tahap
    
    if (code) {
      nossList.push({
        code: String(code).trim(),
        name: String(name).trim(),
        statusTahap: String(statusTahap).trim()
      });
    }
  }

  // ==========================================
  // 2. Baca Tab Jajar (CU)
  // ==========================================
  var jajarSheet = ss.getSheetByName("Jajar");
  var jajarValues = jajarSheet ? jajarSheet.getDataRange().getValues() : [];
  var jajarMap = {};
  
  for (var j = 1; j < jajarValues.length; j++) {
    var programCode = String(jajarValues[j][0] || '').trim(); // Kolum A: Kod Program
    var cuCode = String(jajarValues[j][2] || '').trim();      // Kolum C: Kod CU
    var cuName = String(jajarValues[j][3] || '').trim();      // Kolum D: Nama CU
    var masaRaw = jajarValues[j][4];                          // Kolum E: Masa
    
    var masa = 0;
    var isNA = false;
    
    if (masaRaw === "#N/A" || masaRaw === "" || masaRaw === null || isNaN(masaRaw)) {
      isNA = true;
      masa = 0;
    } else {
      masa = parseFloat(masaRaw);
    }
    
    if (programCode && cuCode) {
      if (!jajarMap[programCode]) {
        jajarMap[programCode] = [];
      }
      jajarMap[programCode].push({
        code: cuCode,
        name: cuName,
        asal: isNA ? 0 : masa,
        isNA: isNA
      });
    }
  }

  // ==========================================
  // 3. Baca Tab WA mengikut Lajur Sebenar & Logik Fallback Pintar
  // ==========================================
  var waSheet = ss.getSheetByName("WA");
  var waValues = waSheet ? waSheet.getDataRange().getValues() : [];
  var waMap = {};
  
  if (waValues.length > 1) {
    var headers = waValues[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
    
    // Cari kedudukan indeks lajur secara dinamik
    var idxProg      = headers.findIndex(function(h) { return h.includes("program"); });
    var idxCu        = headers.findIndex(function(h) { return h.includes("kod cu") || h === "cu"; });
    var idxNoWa      = headers.findIndex(function(h) { return h.includes("cu-wa") || h.includes("kod wa") || h.includes("no"); });
    var idxNamaWa    = headers.findIndex(function(h) { return h === "wa" || h.includes("tajuk") || h.includes("aktiviti"); });
    var idxMasa      = headers.findIndex(function(h) { return h.includes("masa") || h.includes("hours") || h.includes("jam"); });
    var idxWeightage = headers.findIndex(function(h) { return h.includes("weightage") || h.includes("peratus") || h.includes("%"); });

    // Tetapan kedudukan lalai jika tajuk lajur tiada
    if (idxProg === -1) idxProg = 0;     // Kolum A (Kod Program)
    if (idxCu === -1) idxCu = 3;         // Kolum D (Kod CU)
    if (idxNoWa === -1) idxNoWa = 5;     // Kolum F (No WA / CU-WA)
    if (idxNamaWa === -1) idxNamaWa = 7; // Kolum H (Tajuk WA)
    if (idxMasa === -1) idxMasa = 8;     // Kolum I (Masa WA)

    var tempCuGroups = {};

    for (var k = 1; k < waValues.length; k++) {
      var row = waValues[k];
      var pCode = String(row[idxProg] || "").trim();
      var cCode = String(row[idxCu] || "").trim();
      var noWa  = String(row[idxNoWa] || "").trim();
      var wName = String(row[idxNamaWa] || "").trim();
      
      var wMasa = 0;
      if (idxMasa !== -1 && row[idxMasa] !== "" && row[idxMasa] !== null) {
        wMasa = parseFloat(row[idxMasa]) || 0;
      }

      var wWeight = 0;
      if (idxWeightage !== -1 && row[idxWeightage] !== "" && row[idxWeightage] !== null) {
        var cleanWeight = String(row[idxWeightage]).replace("%", "").trim();
        wWeight = parseFloat(cleanWeight) || 0;
      }

      if (cCode && wName) {
        if (!tempCuGroups[cCode]) {
          tempCuGroups[cCode] = { pCode: pCode, items: [] };
        }
        tempCuGroups[cCode].items.push({
          noWa: noWa,
          kodWa: noWa,
          namaWa: wName,
          masa: wMasa,
          weightage: wWeight
        });
      }
    }

    var dummyPattern = [4, 4, 12, 6, 8, 4, 4];

    for (var cuKey in tempCuGroups) {
      var group = tempCuGroups[cuKey];
      var items = group.items;
      var hasValidMasa = items.some(function(it) { return it.masa > 0; });
      var hasValidWeightage = items.some(function(it) { return it.weightage > 0; });

      var processedItems = items.map(function(item, idx) {
        var finalMasa = item.masa;
        var finalWeightage = item.weightage;

        if (hasValidMasa) {
          finalMasa = item.masa;
        } else if (hasValidWeightage) {
          finalWeightage = item.weightage;
          finalMasa = 0;
        } else {
          finalMasa = dummyPattern[idx % dummyPattern.length];
          finalWeightage = 100 / items.length;
        }

        return {
          noWa: item.noWa,
          kodWa: item.kodWa,
          namaWa: item.namaWa,
          masa: finalMasa,
          weightage: finalWeightage
        };
      });

      waMap[cuKey] = processedItems;

      if (group.pCode) {
        var compositeKey = group.pCode + "___" + cuKey;
        waMap[compositeKey] = processedItems;
      }
    }
  }
  
  return {
    noss: nossList,
    jajar: jajarMap,
    waData: waMap
  };
}