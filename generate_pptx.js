const PptxGenJS = require('pptxgenjs');

(async () => {
  const pptx = new PptxGenJS();
  
  // Set layout
  pptx.layout = 'LAYOUT_16x9';

  // Colors
  const COLORS = {
    darkBg: '0F172A',     // Slate 900
    lightBg: 'F8FAFC',    // Slate 50
    primary: '0D9488',    // Teal 600
    primaryDark: '0F766E',// Teal 700
    textDark: '0F172A',   // Slate 900
    textLight: 'F8FAFC',  // Slate 50
    textMuted: '475569',  // Slate 600
    textMutedDark: '94A3B8', // Slate 400
    accent: '06B6D4',     // Cyan 500
    border: 'E2E8F0'      // Slate 200
  };

  // -------------------------------------------------------------
  // Slide 1: Cover (Dark Theme)
  // -------------------------------------------------------------
  const slide1 = pptx.addSlide();
  slide1.background = { color: COLORS.darkBg };
  
  // Decorative lines/shapes
  slide1.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.5, w: 9, h: 0.1, fill: { color: COLORS.primary } });
  
  slide1.addText('NADI SIMRS v2.0', {
    x: 0.5, y: 1.5, w: 9,
    fontSize: 48, bold: true, color: COLORS.primary,
    fontFace: 'Arial'
  });
  
  slide1.addText('Sistem Informasi Manajemen Rumah Sakit & Klinik Terpadu', {
    x: 0.5, y: 2.5, w: 9,
    fontSize: 22, color: COLORS.textLight,
    fontFace: 'Arial'
  });
  
  slide1.addText('Solusi IT Fasyankes yang Modern, Modular, dan Terintegrasi SatuSehat', {
    x: 0.5, y: 3.2, w: 9,
    fontSize: 16, color: COLORS.textMutedDark,
    fontFace: 'Arial'
  });

  slide1.addText([
    { text: 'Target Audiens: ' },
    { text: 'Klinik Kecil & Pratama ', bold: true, color: COLORS.accent },
    { text: 'hingga ' },
    { text: 'RSUD Skala Enterprise', bold: true, color: COLORS.accent }
  ], {
    x: 0.5, y: 4.5, w: 9,
    fontSize: 14, color: COLORS.textLight,
    fontFace: 'Arial'
  });

  // -------------------------------------------------------------
  // Slide 2: Latar Belakang & Regulasi (Light Theme)
  // -------------------------------------------------------------
  const slide2 = pptx.addSlide();
  slide2.background = { color: COLORS.lightBg };
  
  slide2.addText('Tantangan & Regulasi Fasyankes', {
    x: 0.5, y: 0.4, w: 9,
    fontSize: 24, bold: true, color: COLORS.textDark,
    fontFace: 'Arial'
  });

  // Left Column - Tantangan
  slide2.addText('Tantangan Operasional:', { x: 0.5, y: 1.0, w: 4.2, fontSize: 16, bold: true, color: COLORS.primaryDark });
  slide2.addText([
    { text: '• Antrean Menumpuk: Pendaftaran pasien lambat & rawan kesalahan input.\n\n' },
    { text: '• Kebocoran Inventori: Pelacakan obat & BHP di apotek tidak real-time.\n\n' },
    { text: '• Presensi Tidak Akurat: Staf medis melakukan absensi manual/fingerprint yang tidak contactless.\n\n' },
    { text: '• Pemantauan Kinerja Lemah: Manajemen kesulitan melihat produktivitas dokter.' }
  ], {
    x: 0.5, y: 1.5, w: 4.2, h: 3.5,
    fontSize: 13, color: COLORS.textMuted,
    lineSpacing: 18,
    fontFace: 'Arial'
  });

  // Right Column - Regulasi
  slide2.addText('Tuntutan Regulasi & Standar:', { x: 5.2, y: 1.0, w: 4.2, fontSize: 16, bold: true, color: COLORS.primaryDark });
  slide2.addText([
    { text: '• Peraturan Kemenkes (PMK 24/2022):\n', bold: true, color: COLORS.textDark },
    { text: 'Seluruh Fasyankes wajib terintegrasi dengan platform SatuSehat.\n\n' },
    { text: '• Standar Interoperabilitas Nasional:\n', bold: true, color: COLORS.textDark },
    { text: 'Mengadopsi format HL7 FHIR R4, ICD-10, SNOMED-CT, dan LOINC.\n\n' },
    { text: '• Standar BPJS Kesehatan:\n', bold: true, color: COLORS.textDark },
    { text: 'Wajib mendukung bridging V-Claim untuk pembuatan Surat Eligibilitas Peserta (SEP).' }
  ], {
    x: 5.2, y: 1.5, w: 4.2, h: 3.5,
    fontSize: 13, color: COLORS.textMuted,
    lineSpacing: 18,
    fontFace: 'Arial'
  });

  // -------------------------------------------------------------
  // Slide 3: Solusi Modular Nadi SIMRS (Light Theme)
  // -------------------------------------------------------------
  const slide3 = pptx.addSlide();
  slide3.background = { color: COLORS.lightBg };

  slide3.addText('Solusi Kami: Sistem SIMRS Modular', {
    x: 0.5, y: 0.4, w: 9,
    fontSize: 24, bold: true, color: COLORS.textDark,
    fontFace: 'Arial'
  });

  slide3.addText('Nadi SIMRS dirancang modular. Anda hanya membayar modul yang Anda butuhkan!', {
    x: 0.5, y: 0.9, w: 9,
    fontSize: 14, color: COLORS.textMuted,
    fontFace: 'Arial'
  });

  // Box 1: Klinik Kecil (Core Only)
  slide3.addShape(pptx.ShapeType.rect, { x: 0.5, y: 1.5, w: 4.2, h: 3.5, fill: { color: 'FFFFFF' }, line: { color: COLORS.border, width: 1 } });
  slide3.addText('KLINIK KECIL & PRATAMA', { x: 0.7, y: 1.7, w: 3.8, fontSize: 16, bold: true, color: COLORS.primary, fontFace: 'Arial' });
  slide3.addText([
    { text: 'Fokus pada sistem inti (Core) dengan efisiensi biaya maksimal:\n\n' },
    { text: '✓ Pendaftaran mandiri via KTP OCR\n' },
    { text: '✓ Rekam Medis (EMR) SOAP Basic\n' },
    { text: '✓ Modul Billing & Kasir Terintegrasi\n' },
    { text: '✓ Notifikasi WhatsApp Gateway\n' },
    { text: '✓ Tanpa biaya lisensi user (Unlimited Users)\n\n' },
    { text: 'Biaya ringan, operasional teratur!' }
  ], {
    x: 0.7, y: 2.1, w: 3.8, h: 2.7,
    fontSize: 12, color: COLORS.textMuted,
    fontFace: 'Arial'
  });

  // Box 2: RSUD / RS Besar (Enterprise)
  slide3.addShape(pptx.ShapeType.rect, { x: 5.2, y: 1.5, w: 4.2, h: 3.5, fill: { color: 'FFFFFF' }, line: { color: COLORS.primary, width: 1.5 } });
  slide3.addText('RSUD & RUMAH SAKIT BESAR', { x: 5.4, y: 1.7, w: 3.8, fontSize: 16, bold: true, color: COLORS.primaryDark, fontFace: 'Arial' });
  slide3.addText([
    { text: 'Paket Enterprise lengkap dengan skalabilitas tinggi:\n\n' },
    { text: '✓ Kepatuhan penuh SatuSehat FHIR R4\n' },
    { text: '✓ Integrasi BPJS Kesehatan (V-Claim & SEP)\n' },
    { text: '✓ Presensi AI Biometrik (Face-Recognition)\n' },
    { text: '✓ Dispatcher Ambulans GPS (Google Maps)\n' },
    { text: '✓ Modul Khusus: IGD Triase, NICU, VK & Kamar Bedah\n' },
    { text: '✓ Dukungan Multi-Cabang & HIE' }
  ], {
    x: 5.4, y: 2.1, w: 3.8, h: 2.7,
    fontSize: 12, color: COLORS.textMuted,
    fontFace: 'Arial'
  });

  // -------------------------------------------------------------
  // Slide 4: Fitur Utama Klinik (Light Theme)
  // -------------------------------------------------------------
  const slide4 = pptx.addSlide();
  slide4.background = { color: COLORS.lightBg };

  slide4.addText('Fitur Unggulan: Paket Klinik Kecil', {
    x: 0.5, y: 0.4, w: 9,
    fontSize: 24, bold: true, color: COLORS.textDark,
    fontFace: 'Arial'
  });

  const clinicFeatures = [
    { text: 'Pendaftaran Cepat via OCR KTP: Cukup foto KTP pasien, NIK terisi otomatis secara offline (Zero Data Leak).' },
    { text: 'Portal Pasien & Family Portal: Pasien bisa booking jadwal dari rumah. Satu akun bisa mengelola riwayat medis anak dan lansia.' },
    { text: 'EMR SOAP Terstruktur: Catatan medis standar, riwayat alergi, tindakan, dan resep tersimpan aman dan cepat diakses.' },
    { text: 'Auto-Decrement Dispensing: Pembukuan kasir terintegrasi apotek, mengurangi stok obat otomatis saat resep ditebus.' },
    { text: 'WhatsApp Gateway terintegrasi: Mengirim pengingat kontrol, antrean, dan OTP tanpa biaya pulsa SMS.' }
  ];

  slide4.addText(clinicFeatures, {
    x: 0.5, y: 1.2, w: 9, h: 3.8,
    fontSize: 14, color: COLORS.textMuted,
    bullet: true,
    lineSpacing: 18,
    fontFace: 'Arial'
  });

  // -------------------------------------------------------------
  // Slide 5: Fitur Utama RSUD / Enterprise (Light Theme)
  // -------------------------------------------------------------
  const slide5 = pptx.addSlide();
  slide5.background = { color: COLORS.lightBg };

  slide5.addText('Fitur Unggulan: Paket RSUD Enterprise', {
    x: 0.5, y: 0.4, w: 9,
    fontSize: 24, bold: true, color: COLORS.textDark,
    fontFace: 'Arial'
  });

  const enterpriseFeatures = [
    { text: 'Sinkronisasi SatuSehat Kemenkes: Otomatis sinkronisasi data rekam medis ke SatuSehat secara asinkron menggunakan BullMQ (HL7 FHIR R4).' },
    { text: 'Bridging BPJS Kesehatan: Integrasi V-Claim untuk cek kepesertaan dan pembuatan Surat Eligibilitas Peserta (SEP) tanpa ganti tab browser.' },
    { text: 'Presensi Biometrik AI: Presensi staf medis tanpa mesin fingerprint mahal menggunakan kamera tablet + Face-Recognition.' },
    { text: 'GIS Ambulans & Evakuasi Cepat: Pelacakan unit ambulans secara real-time dan pengiriman rute tercepat via Google Maps Navigator ke driver.' },
    { text: 'Modul ICU/NICU, VK & Kamar Bedah: Pencatatan rekam medis terperinci untuk tindakan CITO dan triage IGD berbasis kode warna.' }
  ];

  slide5.addText(enterpriseFeatures, {
    x: 0.5, y: 1.2, w: 9, h: 3.8,
    fontSize: 14, color: COLORS.textMuted,
    bullet: true,
    lineSpacing: 18,
    fontFace: 'Arial'
  });

  // -------------------------------------------------------------
  // Slide 6: Arsitektur & Performa (Dark Theme)
  // -------------------------------------------------------------
  const slide6 = pptx.addSlide();
  slide6.background = { color: COLORS.darkBg };

  slide6.addText('Arsitektur Teknologi & Performa', {
    x: 0.5, y: 0.4, w: 9,
    fontSize: 24, bold: true, color: COLORS.textLight,
    fontFace: 'Arial'
  });

  slide6.addText('Dirancang untuk keandalan ekstrem, kecepatan tinggi, dan bebas biaya lisensi tersembunyi.', {
    x: 0.5, y: 0.9, w: 9,
    fontSize: 14, color: COLORS.textMutedDark,
    fontFace: 'Arial'
  });

  const techSpecs = [
    { text: 'Siap 100k+ Kunjungan/Hari: Dibangun dengan Next.js 15 Server Actions dan PM2 Cluster Mode untuk load balancing optimal.' },
    { text: 'Type-Safe dengan Drizzle ORM: Koneksi PostgreSQL 16 ultra-cepat, didukung GIN Trigram index untuk pencarian nama pasien <0.1 detik.' },
    { text: 'Keamanan Data HIPAA & PMK 24: Semua pemrosesan data sensitif (OCR KTP & Face Recognition) dikerjakan di RAM client (Zero server leak).' },
    { text: 'Penyimpanan Efisien (Object Storage): Mendukung upload foto radiologi dan USG dengan kompresi otomatis ke Object Storage.' }
  ];

  slide6.addText(techSpecs, {
    x: 0.5, y: 1.5, w: 9, h: 3.5,
    fontSize: 14, color: COLORS.textLight,
    bullet: true,
    lineSpacing: 22,
    fontFace: 'Arial'
  });

  // -------------------------------------------------------------
  // Slide 7: Komparasi SIMRS (Light Theme)
  // -------------------------------------------------------------
  const slide7 = pptx.addSlide();
  slide7.background = { color: COLORS.lightBg };

  slide7.addText('Nadi SIMRS vs SIMRS Konvensional', {
    x: 0.5, y: 0.4, w: 9,
    fontSize: 24, bold: true, color: COLORS.textDark,
    fontFace: 'Arial'
  });

  // Table
  const tableRows = [
    [
      { text: 'Aspek Komparasi', options: { bold: true, fill: { color: COLORS.primaryDark }, color: 'FFFFFF', align: 'center' } },
      { text: 'SIMRS Konvensional (Masa Lalu)', options: { bold: true, fill: { color: COLORS.primaryDark }, color: 'FFFFFF', align: 'center' } },
      { text: 'Nadi SIMRS (Masa Depan)', options: { bold: true, fill: { color: COLORS.primary }, color: 'FFFFFF', align: 'center' } }
    ],
    [
      { text: 'Model Lisensi', options: { bold: true } },
      'Vendor Lock-in, biaya lisensi mahal per user/dokter',
      'Open-source core, lisensi unlimited user'
    ],
    [
      { text: 'Pendaftaran Pasien', options: { bold: true } },
      'Lobby antrean panjang, ketik manual nama & alamat',
      'Self-service OCR KTP & pendaftaran online via Family Portal'
    ],
    [
      { text: 'SatuSehat & BPJS', options: { bold: true } },
      'Modul bridging terpisah, biaya integrasi tambahan',
      'Sudah native SatuSehat FHIR R4 & V-Claim BPJS'
    ],
    [
      { text: 'Presensi Karyawan', options: { bold: true } },
      'Fingerprint manual (rawan penularan virus & titip absen)',
      'Face Recognition AI contactless lewat peramban'
    ]
  ];

  slide7.addTable(tableRows, {
    x: 0.5, y: 1.2, w: 9,
    fontSize: 12,
    border: { type: 'solid', color: COLORS.border, size: 1 },
    colW: [2.0, 3.5, 3.5]
  });

  // -------------------------------------------------------------
  // Slide 8: Skema Implementasi & Hubungi Kami (Light Theme)
  // -------------------------------------------------------------
  const slide8 = pptx.addSlide();
  slide8.background = { color: COLORS.lightBg };

  slide8.addText('Langkah Implementasi & Hubungi Kami', {
    x: 0.5, y: 0.4, w: 9,
    fontSize: 24, bold: true, color: COLORS.textDark,
    fontFace: 'Arial'
  });

  slide8.addText('Proses implementasi cepat dan didampingi penuh hingga go-live.', {
    x: 0.5, y: 0.9, w: 9,
    fontSize: 14, color: COLORS.textMuted,
    fontFace: 'Arial'
  });

  // Left Column - Langkah
  slide8.addText('Langkah Kerja:', { x: 0.5, y: 1.4, w: 4.2, fontSize: 16, bold: true, color: COLORS.primaryDark });
  slide8.addText([
    { text: '1. Assessment Kebutuhan: ', bold: true, color: COLORS.textDark },
    { text: 'Pemetaan modul (Core vs Enterprise).\n\n' },
    { text: '2. Setup Infrastruktur: ', bold: true, color: COLORS.textDark },
    { text: 'Instalasi server on-premise RS atau cloud klinik.\n\n' },
    { text: '3. Migrasi & Training: ', bold: true, color: COLORS.textDark },
    { text: 'Migrasi database pasien lama dan pelatihan staf.\n\n' },
    { text: '4. Go-Live & Sertifikasi: ', bold: true, color: COLORS.textDark },
    { text: 'Pendampingan sinkronisasi SatuSehat & BPJS.' }
  ], {
    x: 0.5, y: 1.9, w: 4.2, h: 3.2,
    fontSize: 12, color: COLORS.textMuted,
    lineSpacing: 14,
    fontFace: 'Arial'
  });

  // Right Column - Kontak
  slide8.addShape(pptx.ShapeType.rect, { x: 5.2, y: 1.4, w: 4.2, h: 3.5, fill: { color: COLORS.darkBg } });
  
  slide8.addText('NADI SIMRS ENTERPRISE', {
    x: 5.4, y: 1.7, w: 3.8,
    fontSize: 18, bold: true, color: COLORS.primary,
    fontFace: 'Arial'
  });
  
  slide8.addText('Hubungi kami untuk konsultasi modul, skema harga, dan demonstrasi produk gratis.', {
    x: 5.4, y: 2.2, w: 3.8,
    fontSize: 13, color: COLORS.textMutedDark,
    fontFace: 'Arial'
  });

  slide8.addText([
    { text: '✉ Email: ', bold: true, color: COLORS.textLight },
    { text: 'simrs@nadi.local\n\n' },
    { text: '🌐 Website: ', bold: true, color: COLORS.textLight },
    { text: 'nadi.local / docs.nadi.local\n\n' },
    { text: '📞 WhatsApp: ', bold: true, color: COLORS.textLight },
    { text: '+62 812-3456-7890' }
  ], {
    x: 5.4, y: 2.8, w: 3.8,
    fontSize: 13, color: COLORS.textLight,
    fontFace: 'Arial'
  });

  // Write file
  await pptx.writeFile({ fileName: 'NADI_Presentasi.pptx' });
  console.log('PPTX created successfully!');
})();

