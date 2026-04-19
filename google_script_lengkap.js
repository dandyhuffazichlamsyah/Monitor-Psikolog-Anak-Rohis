function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. MAPPING ARRAY (Merubah Angka Jadi Bahasa Manusia)
    var scalesMap = {
      1: "Nggak Pernah", 
      2: "Sesekali", 
      3: "Sering", 
      4: "Sering Banget"
    };
    
    // 2. MAPPING PERTANYAAN (Dipakai untuk Header & Auto-Diagnosis)
    // Label di ujung string "(Positif)" dipakai backend untuk baca logika reverse skor
    var P = {
       "S1": "Sholat sbg tempat istirahat batin (Positif)", 
       "S2": "Bisa sejuk baca Qur'an (Positif)", 
       "S3": "Buru-buru ibadah buat ngecek chat dia", 
       "S4": "Ikut kajian cuma buat ketemu sosok doi", 
       "S5": "Kesel sama takdir ortu & males ibadah", 
       "S6": "Pura-pura alim biar ga dijudge sirkel", 
       "S7": "Yakin Allah nyiapin pelangi di ujung jalan (Positif)",
       
       "A1": "Fokus diterangin materi kelas (Positif)",
       "A2": "Ambyar nugas krn inget rumah lg tegang",
       "A3": "Konsentrasi pecah mikirin gebetan",
       "A4": "Keteteran akademis krn beban proker Rohis",
       "A5": "Capek tertekan tuntutan nilai orang tua",
       "A6": "Nongkrong di luar krn ga mau pulang ke rumah",
       "A7": "Bisa rapi membagi waktu belajar asik (Positif)",

       "K1": "Merasa sangat aman cerita ke keluarga (Positif)",
       "K2": "Merasa berharga HANYA kalo habis dipuji lawan jenis",
       "K3": "Dengar nada tinggi/silent treatment di rumah",
       "K4": "Tersisih krn sirkel eksklusif pengurus Rohis",
       "K5": "Takut ngasih kritik syuro krn bakal disindir",
       "K6": "Curhat ke lawan jenis krn dirumah tak ada yg denger",
       "K7": "Harus bersih-bersih chat karena ortu asal nge-judge",

       "E1": "Sering bangun tidur seger dan plong (Positif)",
       "E2": "Nangis malam menderita krn kelakuan orang serumah",
       "E3": "Mood seharian ancur krn chat telat dibalas doi",
       "E4": "Pengen pamit dr Rohis tutup usia kepanitiaan krn drama",
       "E5": "Cemburu buta mikirin dia ngobrol sama orang lain",
       "E6": "Lelah nahan uneg-uneg nyimpen aib proker/senior",
       "E7": "Bisa narik napas saat rencana tiba-tiba ancur (Positif)"
    };

    var ts = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd-MMM-yyyy HH:mm");

    // ==========================================
    // SISTEM AUTO-NARASI (DIAGNOSIS PER ANAK)
    // ==========================================
    function genDiag(ids) { 
        var causes = [];
        for(var i=0; i<ids.length; i++) {
           var val = data[ids[i]];
           var qText = P[ids[i]];
           
           if (!val) continue; // Skip if somehow undefined

           // Jika ini soal Negatif (yg merusak mental) -> Bahaya di skala 3 (Sering) atau 4 (Sering Banget)
           if(qText.indexOf("(Positif)") === -1) { 
               if(val >= 3) {
                   causes.push("• " + qText + " => ngaku: " + scalesMap[val]);
               }
           } else {
               // Jika ini soal Positif (pondasi ketahanan mental) -> Bahaya jika dia milih 1 (Nggak) atau 2 (Jarang)
               if(val <= 2) {
                   causes.push("• " + qText.replace(" (Positif)","") + " => ngaku: cuma " + scalesMap[val]);
               }
           }
        }
        if(causes.length === 0) return "✅ Aman. Anak ini relatif stabil dan jernih pada kompartemen ini.";
        return "\n⚠️ DITEMUKAN MASALAH:\n" + causes.join("\n");
    }

    var diagS = genDiag(["S1","S2","S3","S4","S5","S6","S7"]);
    var diagA = genDiag(["A1","A2","A3","A4","A5","A6","A7"]);
    var diagK = genDiag(["K1","K2","K3","K4","K5","K6","K7"]);
    var diagE = genDiag(["E1","E2","E3","E4","E5","E6","E7"]);

    // ===============================================
    // SHEET 1: DASHBOARD STATISTIK (MENGGUNAKAN RUMUS)
    // ===============================================
    // Diciptakan paling awal tapi diisi paling depan tab-nya
    setupDashboard(ss);

    // ===============================================
    // SHEET 2: 📝 REKAP PENILAIAN SINGKAT (EVALUASI)
    // ===============================================
    var shRingkas = getOrCreateSheet(ss, "📝 Lembar Evaluasi Utama", [
        "Waktu", "Nama Siswa", "Kelas", 
        "Skor Spi", "Skor Akd", "Skor Sos", "Skor Ems", 
        "Isu Asmara(X)", "Isu Keluarga(Y)", "Isu Rohis(Z)",
        "Skor Akhir", "Level", "Status Deteksi Kritis"
    ]);

    shRingkas.appendRow([
        ts, data.nama, data.kelas, 
        data.skorSpiritual, data.skorAkademik, data.skorSosial, data.skorEmosi,
        data.indikatorX, data.indikatorY, data.indikatorZ,
        data.skorTotal, data.level, data.kondisi
    ]);

    // Beri peringatan merah apabila menyentuh poin Muntaber/Kritis
    var rLastRow = shRingkas.getLastRow();
    if(data.indikatorX >= 70) shRingkas.getRange(rLastRow, 8).setBackground("#FEE2E2").setFontColor("#991B1B").setFontWeight("bold");
    if(data.indikatorY >= 70) shRingkas.getRange(rLastRow, 9).setBackground("#FEE2E2").setFontColor("#991B1B").setFontWeight("bold");
    if(data.indikatorZ >= 70) shRingkas.getRange(rLastRow, 10).setBackground("#FFFBEB").setFontColor("#92400E").setFontWeight("bold");

    // ===============================================
    // SHEET 3: 📖 RAPOR DIAGNOSIS NARATIF 
    // ===============================================
    var shDetail = getOrCreateSheet(ss, "📖 Rapor Detail Diagnosis", [
        "Waktu", "Nama Siswa", "Kelas", "Peringatan Keseluruhan",
        "1. PENYEBAB KERUSAKAN SPIRITUAL", 
        "2. PENYEBAB DROP AKADEMIK & FOKUS", 
        "3. PENYEBAB KEMUNDURAN SOSIAL", 
        "4. PENYEBAB OVERTHINKING/EMOSI DANGKAL"
    ]);
    
    // Set column wrap agar teks bisa memajang rapi seperti rapor
    shDetail.getRange("E:H").setWrap(true);
    shDetail.setColumnWidth(5, 300); shDetail.setColumnWidth(6, 300);
    shDetail.setColumnWidth(7, 300); shDetail.setColumnWidth(8, 300);

    shDetail.appendRow([
        ts, data.nama, data.kelas, data.kondisi,
        diagS, diagA, diagK, diagE
    ]);

    var dLastRow = shDetail.getLastRow();
    shDetail.setRowHeight(dLastRow, 120); // Lebarkan baris vertikal agar enak dibaca

    // ===============================================
    // SHEET 4: 🗂️ JAWABAN MENTAH (BAHASA REAL)
    // ===============================================
    var shRaw = getOrCreateSheet(ss, "🗂️ Jawaban Mentah per Soal", [
        "Waktu", "Nama", "Kelas", 
        P["S1"], P["S2"], P["S3"], P["S4"], P["S5"], P["S6"], P["S7"],
        P["A1"], P["A2"], P["A3"], P["A4"], P["A5"], P["A6"], P["A7"],
        P["K1"], P["K2"], P["K3"], P["K4"], P["K5"], P["K6"], P["K7"],
        P["E1"], P["E2"], P["E3"], P["E4"], P["E5"], P["E6"], P["E7"]
    ]);
    
    var rawArray = [ts, data.nama, data.kelas];
    var qs = ["S1","S2","S3","S4","S5","S6","S7", "A1","A2","A3","A4","A5","A6","A7", "K1","K2","K3","K4","K5","K6","K7", "E1","E2","E3","E4","E5","E6","E7"];
    for(var j=0; j<qs.length; j++) {
        rawArray.push(scalesMap[data[qs[j]]] || "-");
    }
    shRaw.appendRow(rawArray);


    return ContentService.createTextOutput(JSON.stringify({"result": "success"})).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": error.message, "stack": error.stack})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ====================================================================
// FUNGSI BANTUAN GENERATOR SHEET
// ====================================================================

function getOrCreateSheet(ss, name, headers) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
        sheet = ss.insertSheet(name);
        sheet.appendRow(headers);
        var headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setFontWeight("bold").setBackground("#0F172A").setFontColor("#FFFFFF").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
        sheet.setFrozenRows(1);
        sheet.setRowHeight(1, 40);
    }
    return sheet;
}

function setupDashboard(ss) {
    var dsName = "📊 Dashboard Statistik";
    var sheet = ss.getSheetByName(dsName);
    
    if (!sheet) {
        // Buat sheet baru dan taruh di posisi paling kiri (Indeks 0)
        sheet = ss.insertSheet(dsName, 0); 
    } else {
        // Jika sudah ada, jangan buat ulang agar performanya ngebut
        return; 
    }

    // Set Tampilan Kanvas Dasar
    sheet.getRange("A1:K40").setBackground("#F8FAFC").setFontFamily("Arial");
    sheet.setHiddenGridlines(true);
    sheet.setColumnWidth(2, 280);
    sheet.setColumnWidth(3, 100);

    // Header Biru Gelap
    sheet.getRange("B2:D2").merge().setValue("DASHBOARD SURVEILANS MENTAL ROHIS")
        .setFontWeight("bold").setFontSize(14).setBackground("#0F172A").setFontColor("white")
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.setRowHeight(2, 35);
    
    // RUMUS 1: Total Pengisi Data
    sheet.getRange("B4").setValue("Data Responden Masuk Saat Ini").setFontWeight("bold").setVerticalAlignment("middle");
    sheet.getRange("C4").setFormula("=COUNTA('📝 Lembar Evaluasi Utama'!A2:A)").setFontWeight("bold").setFontSize(16).setHorizontalAlignment("center").setBackground("#E2E8F0").setVerticalAlignment("middle");
    sheet.setRowHeight(4, 30);
    
    // PEMBAGIAN ZONA ANCAMAN
    sheet.getRange("B6:C6").merge().setValue("🔥 KORBAN ZONA KRITIS (Perlu Evakuasi)").setBackground("#EF4444").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.setRowHeight(6, 25);
    
    sheet.getRange("B7").setValue("Siswa dengan Kasus Asmara/Bucin Berat").setFontColor("#DC2626").setVerticalAlignment("middle");
    sheet.getRange("C7").setFormula("=COUNTIF('📝 Lembar Evaluasi Utama'!H2:H, \">=70\")").setFontWeight("bold").setHorizontalAlignment("center").setFontColor("#DC2626").setVerticalAlignment("middle");
    
    sheet.getRange("B8").setValue("Siswa Stres/Depresi Krn Orang Tua/Rumah").setFontColor("#DC2626").setVerticalAlignment("middle");
    sheet.getRange("C8").setFormula("=COUNTIF('📝 Lembar Evaluasi Utama'!I2:I, \">=70\")").setFontWeight("bold").setHorizontalAlignment("center").setFontColor("#DC2626").setVerticalAlignment("middle");

    sheet.getRange("B9").setValue("Siswa Burnout/Toxic/Capek di Organisasi (Rawan Muntaber)").setFontColor("#B45309").setVerticalAlignment("middle");
    sheet.getRange("C9").setFormula("=COUNTIF('📝 Lembar Evaluasi Utama'!J2:J, \">=70\")").setFontWeight("bold").setHorizontalAlignment("center").setFontColor("#B45309").setVerticalAlignment("middle");

    // Boarder Pelindung
    sheet.getRange("B4:C9").setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);

    // KOTAK PENJELASAN UNTUK GURU EKSKUL/PEMBINA
    sheet.getRange("B12:E12").merge().setValue("BAGAIMANA CARA MENGUNAKAN SPREADSHEETS INI?").setFontWeight("bold").setBackground("#3B82F6").setFontColor("white").setVerticalAlignment("middle");
    
    sheet.getRange("B13:E13").merge().setValue("1. Tab 'Lembar Evaluasi Utama' berisi perhitungan kasar. Jika Anda ingin melihat nilai angka mutlak spiritual, dan persentase indikatornya, buka sheet tersebut.")
         .setWrap(true).setVerticalAlignment("middle").setFontSize(10).setBackground("white");
    sheet.getRange("B14:E14").merge().setValue("2. Tab 'Rapor Detail Diagnosis' ADALAH JANTUNG DARI PLATFORM INI. Sistem telah membacakan isi kepala anak tersebut dan menceritakan secara naratif kepada Anda apa AKAR MASALAH KENAPA DIA MALAS ROHIS/BELAJAR. Anda cukup membaca tanpa harus menebak.")
         .setWrap(true).setVerticalAlignment("middle").setFontSize(10).setBackground("white");
    sheet.getRange("B15:E15").merge().setValue("3. Tab 'Jawaban Mentah' berisi bahasa gaul sejujur-jujurnya dari kuesioner jika Anda butuh investigasi detail setiap butir yang anak tekan.")
         .setWrap(true).setVerticalAlignment("middle").setFontSize(10).setBackground("white");

    sheet.setRowHeight(12, 25); sheet.setRowHeight(13, 40); sheet.setRowHeight(14, 55); sheet.setRowHeight(15, 40);
    sheet.getRange("B12:E15").setBorder(true, true, true, true, true, true, "#3B82F6", SpreadsheetApp.BorderStyle.SOLID);
}
