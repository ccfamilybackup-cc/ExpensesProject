var PIN_RAHASIA = "140997"; 
var SHEET_ID = "1pdMumDgPhwWKiFxRw8Fd_Na1ep4dU4ABA_fzuPSGgO8"; // ID Murni

// ==========================================
// 1. PINTU UTAMA UNTUK AMBIL DATA & LOGIN (GET)
// ==========================================
function doGet(e) {
  // Validasi Parameter PIN
  var pin = e && e.parameter ? e.parameter.pin : null; 
  if (pin !== PIN_RAHASIA) { 
    return ContentService.createTextOutput(JSON.stringify({ error: "Unauthorized" }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  
  try {
    // Buka Spreadsheet
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheets()[0]; 
    
    // Ambil Budget Total (G2)
    var budgetTotal = sheet.getRange("G2").getValue() || 0; 
    
    // Ambil Semua Data di Sheet
    var data = sheet.getDataRange().getValues();
    var history = []; 
    
    // Loop mulai dari baris ke-2 (index 1) untuk melewati header
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Validasi: Hanya ambil jika kolom C (Nama Barang - index 2) tidak kosong
      if (row[2] && row[2].toString().trim() !== "") { 
        var dateValue = row[0];
        var dateStr = "";
        
        // Proteksi konversi format Tanggal (Kolom A)
        if (dateValue instanceof Date) {
          // Jika format di sheet berupa Objek Date asli, konversi ke DD/MM/YYYY agar sesuai dengan parseTransactionDate
          dateStr = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else if (dateValue) {
          var strValue = dateValue.toString().trim();
          if (strValue !== "") {
            // Jika formatnya sudah DD/MM/YYYY, gunakan sebagaimana adanya
            if (strValue.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
              dateStr = strValue;
            }
            // Jika formatnya YYYY-MM-DD, ubah ke DD/MM/YYYY
            else if (strValue.match(/^\d{4}-\d{2}-\d{2}$/)) {
              var parts = strValue.split("-");
              dateStr = parts[2] + "/" + parts[1] + "/" + parts[0];
            } else {
              // Untuk format lain, coba parsing otomatis
              dateStr = strValue;
            }
          }
        } else if (typeof dateValue === "number") {
          // Jika tanggal disimpan sebagai nomor serial Excel
          var jsDate = new Date((dateValue - 25569) * 86400 * 1000);
          dateStr = Utilities.formatDate(jsDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
        }
        
        // Fallback: jika tanggal kosong, gunakan hari ini
        if (!dateStr || dateStr === "") {
          dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
        }
        
        // Ambil data Jam dari Kolom B (index 1)
        var jamValue = row[1] ? row[1].toString().trim() : "";
        
        // Masukkan data ke array history
        history.push({ 
          date: dateStr,
          jam: jamValue, 
          barang: row[2] ? row[2].toString().trim() : "", 
          kategori: row[3] ? row[3].toString().trim() : "", 
          jumlah: row[4] ? Number(row[4]) : 0
        });
      }
    }
    
    // Susun Hasil Akhir - JANGAN di-reverse agar urutan sesuai (terbaru di atas akan di-handle di HTML)
    var result = { 
      budget: Number(budgetTotal), 
      history: history
    }; 
    
    Logger.log("Data yang dikirim ke HTML: " + JSON.stringify(result));
    
    return ContentService.createTextOutput(JSON.stringify(result))
                         .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("Error di doGet: " + err.message);
    return ContentService.createTextOutput(JSON.stringify({ error: "Internal Error", message: err.message }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 2. PINTU UTAMA UNTUK MENERIMA / SIMPAN DATA (POST)
// ==========================================
function doPost(e) {
  try {
    // Ambil data parameter dari form body HTML
    var p = e.parameter;
    var pin = p.pin;
    var jam = p.jam; 
    var barang = p.barang;
    var kategori = p.kategori;
    var jumlah = Number(p.jumlah) || 0;

    // Validasi PIN sebelum menyimpan data
    if (pin !== PIN_RAHASIA) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Unauthorized" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // Buka Spreadsheet
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheets()[0]; 

    // MENGUBAH FORMAT TANGGAL SIMPAN KE DD/MM/YYYY (Contoh: 11/07/2026)
    var tanggalHariIni = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    // Jika parameter "jam" dari HTML kosong, otomatis isi dengan jam sistem saat ini (Format: HH:mm)
    if (!jam || jam.trim() === "") {
      jam = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm");
    }

    // Masukkan data ke baris paling bawah secara berurutan:
    // Kolom A: Tanggal (DD/MM/YYYY), Kolom B: Jam, Kolom C: Barang, Kolom D: Kategori, Kolom E: Jumlah
    sheet.appendRow([tanggalHariIni, jam, barang, kategori, jumlah]);

    Logger.log("Data berhasil disimpan: " + barang + " - Rp " + jumlah + " pada " + tanggalHariIni + " " + jam);

    // Kirim respon sukses ke HTML
    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data saved successfully" }))
                         .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("Error di doPost: " + err.message);
    return ContentService.createTextOutput(JSON.stringify({ error: "Post Failed", message: err.message }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
