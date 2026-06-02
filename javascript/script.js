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

    // ==========================================
    // ⚙️ PENGATURAN DURASI WAKTU UJIAN (EDIT DI SINI)
    // ==========================================
    const DURASI_MENIT = 30; // Ganti angka ini untuk mengatur durasi menit ujian
    
    // ==========================================
    // 🛠️ DATA SOAL UJIAN (EDIT TEKS / GAMBAR DI SINI)
    // ==========================================
    const examData = [
        {
            id: 1,
            tipe: "PG", 
            question: "Manakah organel sel berikut yang berfungsi sebagai tempat respirasi seluler untuk menghasilkan energi?",
            image: "", 
            options: ["Ribosom", "Mitokondria", "Lisosom", "Badan Golgi"],
            correct: 1 
        },
        {
            id: 2,
            tipe: "PGK", 
            question: "Perhatikan karakteristik zat cair! Manakah pernyataan di bawah ini yang BENAR mengenai sifat air pada tekanan normal? (Pilih semua jawaban yang benar)",
            image: "https://via.placeholder.com/400x180?text=Simulasi+Gambar+Sains+SMP", 
            options: [
                "Membeku pada suhu 0°C",
                "Mendidih pada suhu 80°C",
                "Memiliki masa jenis terbesar pada suhu 4°C",
                "Bentuknya selalu tetap tidak mengikuti wadah"
            ],
            correct: [0, 2] 
        },
        {
            id: 3,
            tipe: "BS", 
            question: "Pernyataan: Lagu 'Indonesia Raya' diciptakan oleh Ibu Sud.",
            image: "",
            options: ["BENAR", "SALAH"],
            correct: 1 
        }
    ];

    let currentIndex = 0;
    let userAnswers = new Array(examData.length).fill(null).map((_, i) => examData[i].tipe === "PGK" ? [] : null);
    let sisaWaktu = DURASI_MENIT * 60; 
    let infoSiswa = { noPeserta: "", nama: "" };
    let timerInterval;

    function prosesLogin() {
        const id = document.getElementById('login-id').value.trim();
        const pass = document.getElementById('login-pass').value.trim();
        
        if(!id || !pass) return alert("Harap isi ID Peserta dan Password!");
        
        document.getElementById('btn-login').innerText = "Memvalidasi Sistem...";
        document.getElementById('btn-login').disabled = true;

        google.script.run.withSuccessHandler(function(hasil) {
            if(hasil.status === "valid") {
                infoSiswa.noPeserta = id;
                infoSiswa.nama = hasil.nama;
                
                document.getElementById('nama-siswa-display').innerText = hasil.nama + " (" + id + ")";
                document.getElementById('layer-login').style.display = 'none';
                document.getElementById('layer-ujian').style.display = 'block';
                
                initUjian();
            } else {
                alert("ID Peserta atau Password tidak ditemukan / salah!");
                document.getElementById('btn-login').innerText = "Masuk Jendela Ujian";
                document.getElementById('btn-login').disabled = false;
            }
        }).cekLogin(id, pass);
    }

    function initUjian() {
        tampilkanSoal();
        buatNavigasiNomor();
        startTimer();
    }

    function tampilkanSoal() {
        const soal = examData[currentIndex];
        document.getElementById('nomor-soal-display').innerText = "Soal Nomor " + (currentIndex + 1);
        document.getElementById('teks-soal-display').innerText = soal.question;

        const badge = document.getElementById('jenis-soal-badge');
        if(soal.tipe === "PG") badge.innerText = "Pilihan Ganda";
        else if(soal.tipe === "PGK") badge.innerText = "Pilihan Ganda Kompleks (Jawaban Jamak)";
        else if(soal.tipe === "BS") badge.innerText = "Pernyataan Benar / Salah";

        const imgNode = document.getElementById('soal-gambar');
        if(soal.image && soal.image.trim() !== "") {
            imgNode.src = soal.image;
            imgNode.style.display = "block";
        } else {
            imgNode.style.display = "none";
        }

        const boxJawaban = document.getElementById('box-pilihan-jawaban');
        boxJawaban.innerHTML = "";
        const opsiHuruf = ['A', 'B', 'C', 'D', 'E'];

        soal.options.forEach((opsi, idx) => {
            const node = document.createElement('div');
            node.className = "option-node";
            
            let inputType = (soal.tipe === "PGK") ? "checkbox" : "radio";
            let isChecked = "";

            if(soal.tipe === "PGK") {
                if(userAnswers[currentIndex].includes(idx)) isChecked = "checked";
                node.setAttribute('onclick', `pilihJawabanPGK(${idx})`);
            } else {
                if(userAnswers[currentIndex] === idx) isChecked = "checked";
                node.setAttribute('onclick', `pilihJawabanTunggal(${idx})`);
            }

            let prefiks = soal.tipe === "BS" ? "" : `<strong>${opsiHuruf[idx]}.</strong> `;

            node.innerHTML = `
                <input type="${inputType}" name="opsi-jawaban" id="opt-${idx}" ${isChecked} onclick="event.stopPropagation(); ${soal.tipe==='PGK'?'pilihJawabanPGK':'pilihJawabanTunggal'}(${idx})">
                <label>${prefiks}${opsi}</label>
            `;
            boxJawaban.appendChild(node);
        });

        // Pengkondisian Tombol Navigasi
        document.getElementById('btn-prev').disabled = (currentIndex === 0);
        
        // 🌟 LOGIKA TOMBOL SELESAI OTOMATIS DI NOMOR TERAKHIR
        const rightNavContainer = document.getElementById('nav-right-container');
        if(currentIndex === examData.length - 1) {
            rightNavContainer.innerHTML = `<button class="btn btn-nav-finish" onclick="klikSelesai()">Selesai Ujian ✔</button>`;
        } else {
            rightNavContainer.innerHTML = `<button class="btn btn-nav-next" id="btn-next" onclick="geserSoal(1)">Berikutnya ▶</button>`;
        }

        refreshNavigasiWarna();
    }

    function pilihJawabanTunggal(idx) {
        userAnswers[currentIndex] = idx;
        document.getElementById(`opt-${idx}`).checked = true;
        refreshNavigasiWarna();
    }

    function pilihJawabanPGK(idx) {
        let arr = userAnswers[currentIndex];
        if(arr.includes(idx)) {
            userAnswers[currentIndex] = arr.filter(item => item !== idx);
            document.getElementById(`opt-${idx}`).checked = false;
        } else {
            userAnswers[currentIndex].push(idx);
            document.getElementById(`opt-${idx}`).checked = true;
        }
        refreshNavigasiWarna();
    }

    function geserSoal(arah) {
        currentIndex += arah;
        tampilkanSoal();
    }

    function lompatKeSoal(idx) {
        currentIndex = idx;
        tampilkanSoal();
    }

    function buatNavigasiNomor() {
        const grid = document.getElementById('grid-nomor-soal');
        grid.innerHTML = "";
        examData.forEach((_, idx) => {
            const numBox = document.createElement('div');
            numBox.id = "num-box-" + idx;
            numBox.className = "num-box";
            numBox.innerText = idx + 1;
            numBox.setAttribute('onclick', `lompatKeSoal(${idx})`);
            grid.appendChild(numBox);
        });
    }

    function refreshNavigasiWarna() {
        examData.forEach((soal, idx) => {
            const el = document.getElementById("num-box-" + idx);
            if(!el) return;
            
            el.className = "num-box";
            let sudahJawab = false;
            if(soal.tipe === "PGK") {
                if(userAnswers[idx] && userAnswers[idx].length > 0) sudahJawab = true;
            } else {
                if(userAnswers[idx] !== null) sudahJawab = true;
            }

            if(sudahJawab) el.classList.add('answered');
            if(idx === currentIndex) el.classList.add('active');
        });
    }

    // ⏳ HANDLER TIMER MUNDUR
    function startTimer() {
        timerInterval = setInterval(() => {
            if(sisaWaktu <= 0) {
                clearInterval(timerInterval);
                alert("Waktu habis! Jawaban Anda akan otomatis dikirim.");
                kirimNilaiKeSheet();
            } else {
                sisaWaktu--;
                let m = Math.floor(sisaWaktu / 60);
                let s = sisaWaktu % 60;
                document.getElementById('timer').innerText = `Sisa Waktu: ${m.toString().padStart(2,'0')}:${s.toString().padStart(2,'0')}`;
            }
        }, 1000);
    }

    function klikSelesai() {
        let kosong = 0;
        examData.forEach((soal, idx) => {
            if(soal.tipe === "PGK" && userAnswers[idx].length === 0) kosong++;
            if(soal.tipe !== "PGK" && userAnswers[idx] === null) kosong++;
        });

        let pesan = "Apakah Anda yakin ingin menyelesaikan ujian dan mengirim jawaban?";
        if(kosong > 0) pesan = `⚠️ PERINGATAN: Masih ada ${kosong} soal yang BELUM dijawab!\n\n` + pesan;

        if(confirm(pesan)) {
            clearInterval(timerInterval); // Matikan timer saat selesai diklik
            kirimNilaiKeSheet();
        }
    }

    function kirimNilaiKeSheet() {
        let benar = 0;
        
        examData.forEach((soal, idx) => {
            if(soal.tipe === "PGK") {
                let jawabanUser = [...userAnswers[idx]].sort().toString();
                let jawabanBenar = [...soal.correct].sort().toString();
                if(jawabanUser === jawabanBenar) benar++;
            } else {
                if(userAnswers[idx] === soal.correct) benar++;
            }
        });

        let totalSkor = Math.round((benar / examData.length) * 100);

        document.getElementById('layer-ujian').style.display = 'none';
        document.getElementById('layer-selesai').style.display = 'block';
        document.getElementById('skor-akhir').innerText = totalSkor;
        document.getElementById('rekap-benar').innerText = `Benar ${benar} dari ${examData.length} Soal`;

        // Mengirimkan data via Google Apps Script ke Tab "Hasil"
        google.script.run.simpanHasil({
            noPeserta: infoSiswa.noPeserta,
            nama: infoSiswa.nama,
            skor: totalSkor,
            detailBenar: `${benar} dari ${examData.length}`
        });
    }