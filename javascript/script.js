function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Asesmen Nasional - Simulasi TKA SMP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fungsi memeriksa ID dan Password Siswa di Sheet
function cekLogin(noPeserta, password) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Siswa");
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() == noPeserta && data[i][1].toString() == password) {
      return { status: "valid", nama: data[i][2] };
    }
  }
  return { status: "salah" };
}

// Fungsi menyimpan hasil ujian ke Sheet Hasil
function simpanHasil(dataUjian) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hasil");
  sheet.appendRow([
    new Date(),
    dataUjian.noPeserta,
    dataUjian.nama,
    dataUjian.skor,
    dataUjian.detailBenar
  ]);
  return { status: "sukses" };
}