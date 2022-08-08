VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIUtama 
   BackColor       =   &H8000000C&
   Caption         =   "Medifirst2000 - Kasir Sentral  (Cashier & Bill Payment)"
   ClientHeight    =   8130
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   11280
   Icon            =   "MDIFrm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIFrm1.frx":0CCA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CDPrinter 
      Left            =   0
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7875
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
            MinWidth        =   7832
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   9596
            MinWidth        =   9596
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "08/01/2020"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "12:35"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MEditKomponen 
      Caption         =   "Edit Komponen DI PAKE FRM TAGIHAN PASIEN"
      Visible         =   0   'False
      Begin VB.Menu MEditTanggunganPenjamin 
         Caption         =   "Edit Tanggungan Penjamin"
      End
   End
   Begin VB.Menu MEditKomponenVerifikasi 
      Caption         =   "Edit Komponen DI PAKE FRM TAGIHAN PASIEN Verifikasi"
      Visible         =   0   'False
      Begin VB.Menu MEditTanggunganPenjaminVerifikasi 
         Caption         =   "Edit Tanggungan Penjamin Untuk Verifikasi"
      End
   End
   Begin VB.Menu mnuEditKomponenTarifdiTambahPelayanan 
      Caption         =   "Edit Komponen Tarif di Tambah Pelayanan"
      Visible         =   0   'False
      Begin VB.Menu mnuEditKomponenTarif 
         Caption         =   "EditKomponenTarif"
      End
   End
   Begin VB.Menu mnberkas 
      Caption         =   "&Berkas"
      Begin VB.Menu mnudata 
         Caption         =   "Data"
         Begin VB.Menu mnuCariDataPasien 
            Caption         =   "Cari Data Pasien"
            Begin VB.Menu mnuDaftarPasienPulang 
               Caption         =   "Daftar Pasien Pulang"
               Shortcut        =   {F3}
            End
            Begin VB.Menu mnuPelayananDepositPasien2 
               Caption         =   "Daftar Pasien Aktif "
               Shortcut        =   {F2}
            End
            Begin VB.Menu frmDaftarPenjualanApotik 
               Caption         =   "Daftar Penjualan Apotik"
               Shortcut        =   ^{F5}
            End
            Begin VB.Menu mnuPembayaranOtomatis 
               Caption         =   "Daftar Pasien Asuransi"
            End
            Begin VB.Menu mnugariss 
               Caption         =   "-"
            End
            Begin VB.Menu mnuDaftarPasienSudahBayar 
               Caption         =   "Daftar Pembayaran Pasien"
               Shortcut        =   {F4}
            End
            Begin VB.Menu mnupembayaranapotik 
               Caption         =   "Daftar Pembayaran Apotik"
               Shortcut        =   ^{F4}
            End
            Begin VB.Menu mnuDaftarPasienDeposit 
               Caption         =   "Daftar Pembayaran Deposit"
               Shortcut        =   ^{F2}
            End
            Begin VB.Menu mnugaris2 
               Caption         =   "-"
            End
            Begin VB.Menu mnuDaftarPasienReturPembayaran 
               Caption         =   "Daftar Retur Pelayanan Pasien"
            End
            Begin VB.Menu mnuReturPelayananApotik 
               Caption         =   "Daftar Retur Pelayanan Apotik"
            End
            Begin VB.Menu mnugaris3 
               Caption         =   "-"
            End
            Begin VB.Menu mnuReturPembayaran 
               Caption         =   "Daftar Pengeluaran Retur Pasien"
            End
            Begin VB.Menu mnuPengeluaranReturApotik 
               Caption         =   "Daftar Pengeluaran Retur Apotik"
            End
            Begin VB.Menu mnugaris4 
               Caption         =   "-"
            End
            Begin VB.Menu mnuDaftarPasienKredit 
               Caption         =   "Daftar Piutang Kredit Pasien"
               Shortcut        =   {F8}
            End
            Begin VB.Menu mnuDaftarPasienKlaim 
               Caption         =   "Daftar Piutang Penjamin Pasien"
            End
            Begin VB.Menu mnuSisaTagihan 
               Caption         =   "Daftar Piutang Kredit Penjamin Pasien"
            End
            Begin VB.Menu mnugaris5 
               Caption         =   "-"
            End
            Begin VB.Menu mnuPembayaranPenjaminPasien 
               Caption         =   "Daftar Pembayaran Penjamin Pasien"
            End
            Begin VB.Menu mnuInfoDaftarBatalKuitansi 
               Caption         =   "Daftar Pembatalan Kuitansi"
            End
         End
         Begin VB.Menu batascaridata 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPenerimaan 
            Caption         =   "Penerimaan Umum"
            Begin VB.Menu mnuPelayananPasien 
               Caption         =   "Pelayanan Pasien"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuPelayananUmum 
               Caption         =   "Penerimaan Kas Umum"
               Shortcut        =   ^{F6}
            End
            Begin VB.Menu mnuPenerimaanUmum 
               Caption         =   "Daftar Penerimaan Umum"
            End
            Begin VB.Menu mnuPelayananDepositPasien 
               Caption         =   "Daftar Pasien Aktif"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuPelayananKlaimPenjamin 
               Caption         =   "Pelayanan Klaim Penjamin"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuTRS 
               Caption         =   "Pelayanan Pasien Klaim TRS"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuPelayananCicilanPasien 
               Caption         =   "Pelayanan Cicilan Pasien"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuPengeluaran 
            Caption         =   "Pengeluaran Umum"
            Begin VB.Menu mnuReturTagihanPerPelayanan 
               Caption         =   "Retur Tagihan Per Pelayanan"
               Shortcut        =   ^{F7}
               Visible         =   0   'False
            End
            Begin VB.Menu lnretur 
               Caption         =   "-"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuPengeluaranUmum 
               Caption         =   "Pengeluaran Kas Umum"
            End
            Begin VB.Menu mnuDafPengeluaranUmum 
               Caption         =   "Daftar Pengeluaran Umum"
            End
            Begin VB.Menu mnuDaftarTagihanSupplier 
               Caption         =   "Daftar Tagihan Supplier"
            End
            Begin VB.Menu mnuDafPembayaranSupplier 
               Caption         =   "Daftar Pembayaran Supplier"
            End
            Begin VB.Menu mnuPengeluaranSupplier 
               Caption         =   "Pengeluaran Supplier"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuVerifikasi 
            Caption         =   "Verifikasi"
            Visible         =   0   'False
            Begin VB.Menu mnuTagihanPasien 
               Caption         =   "Tagihan Pasien"
               Visible         =   0   'False
            End
            Begin VB.Menu deline 
               Caption         =   "-"
            End
            Begin VB.Menu mnuDaftarPasienygBelumDanSudahPernahBayar 
               Caption         =   "Daftar Pasien yg Belum dan Sudah Pernah Bayar"
               Visible         =   0   'False
            End
            Begin VB.Menu MDaftarDataPenerimaanKasirSudahDiPosting 
               Caption         =   "Daftar Data Penerimaan Kasir Sudah di Posting"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuDaftarPasienMshAktif 
               Caption         =   "Daftar Pasien Masih Aktif Belum Bayar"
               Visible         =   0   'False
            End
            Begin VB.Menu deline2 
               Caption         =   "-"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuValidasiUlangYangGagalBayar 
               Caption         =   "Validasi Ulang Yang Gagal Bayar"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuVerPemakaianAsuransi 
               Caption         =   "Pemakaian Asuransi Pasien"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuVerifikasiBayar 
               Caption         =   "Pembayaran"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuLine 
               Caption         =   "-"
               Visible         =   0   'False
            End
            Begin VB.Menu MPostingDataKasirPenerimaan 
               Caption         =   "Posting Data Kasir Penerimaan"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuPositngDataKasirPendapatan 
               Caption         =   "Posting Data Kasir Pendapatan"
               Visible         =   0   'False
            End
            Begin VB.Menu linePosint 
               Caption         =   "-"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuPostingDataRemunerasi 
               Caption         =   "Posting Data Remunerasi"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuPostingDataPendapatanRemunerasi 
               Caption         =   "Posting Data Pendapatan Remunerasi"
               Visible         =   0   'False
            End
            Begin VB.Menu linePositn1 
               Caption         =   "-"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu line 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuInformasiTarifPelayanan 
            Caption         =   "Informasi Tarif Pelayanan"
            Visible         =   0   'False
         End
         Begin VB.Menu mnutp 
            Caption         =   "Tagihan Pasien"
            Visible         =   0   'False
         End
         Begin VB.Menu MEditTagihanPasien 
            Caption         =   "Edit Tagihan Pasien"
            Visible         =   0   'False
         End
         Begin VB.Menu MKasirUmum 
            Caption         =   "Kasir Umum"
            Visible         =   0   'False
         End
         Begin VB.Menu MPembayaranDepositBiayaPerawatan 
            Caption         =   "Pembayaran Deposit Biaya Perawatan"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuBar1 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDReturStruk 
            Caption         =   "Retur Struk"
            Visible         =   0   'False
         End
         Begin VB.Menu MDaftarReturStruk 
            Caption         =   "Daftar Retur Struk"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuBDBar1 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudpsb 
            Caption         =   "Daftar Pasien Sudah Bayar"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuIDaftarPasienNgutang 
            Caption         =   "Daftar Pasien yang Berhutang"
            Visible         =   0   'False
         End
         Begin VB.Menu MDaftarPembayaranDepositBiayaPerawatan 
            Caption         =   "Daftar Pembayaran Deposit Biaya Perawatan"
            Visible         =   0   'False
         End
         Begin VB.Menu mnusepTPA 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu MTagihanPasienApotik 
            Caption         =   "Tagihan Pasien Apotik"
            Visible         =   0   'False
         End
         Begin VB.Menu MTagihanSupplier 
            Caption         =   "Tagihan Supplier"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuBDBar2 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudaftarKasMasuk 
            Caption         =   "Daftar Kas Masuk"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuKasKeluar 
            Caption         =   "Daftar Kas Keluar"
            Visible         =   0   'False
         End
         Begin VB.Menu LInformasiTarifPelayanan 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnukul 
            Caption         =   "Komponen Unit Laporan"
            Visible         =   0   'False
         End
         Begin VB.Menu MKonversiKomponenUnitKePelayanan 
            Caption         =   "Konversi Unit Laporan Ke Pelayanan"
         End
         Begin VB.Menu sepmnukul 
            Caption         =   "-"
         End
         Begin VB.Menu MVerifikasi 
            Caption         =   "Verifikasi"
            Begin VB.Menu MDaftarPasienUntukVerifikasi 
               Caption         =   "Daftar Pasien Untuk Verifikasi"
               Visible         =   0   'False
            End
            Begin VB.Menu MPembayaranOtomatis 
               Caption         =   "Pembayaran Otomatis"
            End
            Begin VB.Menu MDaftarKasirUntukClaim 
               Caption         =   "Daftar Kasir Untuk Claim"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu MInformasiTarifPelayanan 
            Caption         =   "Informasi Tarif Pelayanan"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnusepTP 
         Caption         =   "-"
      End
      Begin VB.Menu mSettingPrinter 
         Caption         =   "Setting Printer"
         Shortcut        =   ^P
      End
      Begin VB.Menu mGantiKataKunci 
         Caption         =   "Ganti Kata Kunci"
         Shortcut        =   ^G
      End
      Begin VB.Menu mspace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnlogout 
         Caption         =   "Log Off"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnSelesai 
         Caption         =   "Keluar"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "Inf&ormasi"
      Begin VB.Menu mnuInfoJenisPelayanan 
         Caption         =   "Jenis - Pelayanan"
      End
      Begin VB.Menu mnuInfoKelasPel 
         Caption         =   "Kelas Pelayanan"
      End
      Begin VB.Menu mnuInfoTarifPel 
         Caption         =   "Tarif Pelayanan"
      End
      Begin VB.Menu batasinfopenjamin 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfoKelompokPasien 
         Caption         =   "Kelompok - Penjamin Pasien"
         Begin VB.Menu mnuProfilPenjamin 
            Caption         =   "Profil Penjamin"
         End
         Begin VB.Menu mnuKelompokPenjamin 
            Caption         =   "Kelompok Penjamin"
         End
      End
      Begin VB.Menu mnuInfoPaketPelayananPenjamin 
         Caption         =   "Paket Pelayanan Penjamin"
      End
      Begin VB.Menu mnuInforTarifPaketPenjamin 
         Caption         =   "Tarif Paket Pelayanan Penjamin"
      End
      Begin VB.Menu mnuInfoNonPaket 
         Caption         =   "Tarif Non Paket Pelayanan Penjamin"
      End
      Begin VB.Menu mnuInfoPerTindakan 
         Caption         =   "Tanggungan Paket Pelayanan Penjamin Per Tindakan"
      End
      Begin VB.Menu mnuIndoTindaknonTanggungan 
         Caption         =   "Pelayanan Tindakan Non Tanggungan Penjamin"
      End
      Begin VB.Menu mnuInfoKonversi 
         Caption         =   "Konversi Paket Penjamin ke Pelayanan"
      End
   End
   Begin VB.Menu mnInventory 
      Caption         =   "&Inventory"
      Begin VB.Menu mnPemesananBarang 
         Caption         =   "Pemesanan Barang"
      End
      Begin VB.Menu mnuPemakaianBahandanAlat 
         Caption         =   "Pemakaian Bahan dan Alat"
         Visible         =   0   'False
      End
      Begin VB.Menu grs1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mBarangMedis 
         Caption         =   "Barang Medis"
         Visible         =   0   'False
         Begin VB.Menu mStokBarang 
            Caption         =   "Stok Barang"
         End
         Begin VB.Menu mClosingStok 
            Caption         =   "Closing Stok"
            Begin VB.Menu mCetakLembarInput 
               Caption         =   "Cetak Lembar Input"
            End
            Begin VB.Menu mInputStokOpn 
               Caption         =   "Input Stok Opname"
            End
            Begin VB.Menu mNilaiPersediaan 
               Caption         =   "Nilai Persediaan"
            End
         End
         Begin VB.Menu ln1 
            Caption         =   "-"
         End
         Begin VB.Menu MInformasiPemesananBarang 
            Caption         =   "Informasi Pemesanan && Penerimaan Barang"
         End
         Begin VB.Menu POAKaryawan 
            Caption         =   "Informasi Pemakaian Barang"
         End
         Begin VB.Menu mLapSaldoBarang 
            Caption         =   "Laporan Saldo Barang"
         End
      End
      Begin VB.Menu mnBarangNM 
         Caption         =   "Barang Non Medis"
         Begin VB.Menu mnStokBarangNM 
            Caption         =   "Stok Barang"
         End
         Begin VB.Menu mnKondisiBarangNM 
            Caption         =   "Kondisi Barang"
         End
         Begin VB.Menu mnMutasiBarangNM 
            Caption         =   "Mutasi Barang"
         End
         Begin VB.Menu mnClosingStokNM 
            Caption         =   "Closing Stok"
            Begin VB.Menu mnCetakLembarInputNM 
               Caption         =   "Cetak Lembar Input"
            End
            Begin VB.Menu mnInputStokOpname 
               Caption         =   "Input Stok Opname"
            End
            Begin VB.Menu mnNilaiPersediaan 
               Caption         =   "Nilai Persediaan"
            End
         End
         Begin VB.Menu mnStokNM 
            Caption         =   "-"
         End
         Begin VB.Menu mnInfoPesanBarangNM 
            Caption         =   "Informasi Pemesanan && Penerimaan Barang"
         End
         Begin VB.Menu mnLapSaldoBarangNM 
            Caption         =   "Laporan Saldo Barang"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuI 
      Caption         =   "In&formasi"
      Visible         =   0   'False
      Begin VB.Menu mnuCariDataTagihanSupplier 
         Caption         =   "Cari Data Tagihan Supplier"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDataTagihanPasienApotik 
         Caption         =   "Cari Data Tagihan Pasien Apotik"
         Visible         =   0   'False
      End
      Begin VB.Menu sepdpsb 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuLap 
      Caption         =   "&Laporan"
      Begin VB.Menu mnuPenerimaanKasir 
         Caption         =   "Penerimaan Kasir"
         Begin VB.Menu mnuLapPelayananPasien 
            Caption         =   "Pelayanan Pasien"
         End
         Begin VB.Menu mnuPelayananRuangan 
            Caption         =   "Pelayanan Ruangan"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuLapPelayananUmum 
            Caption         =   "Pelayanan Umum"
         End
         Begin VB.Menu mnuLapPelayananDepositPasien 
            Caption         =   "Pelayanan Deposit Pasien"
         End
         Begin VB.Menu mnuLapPelayananKlaimPenjamin 
            Caption         =   "Pelayanan Klaim Penjamin"
         End
         Begin VB.Menu mnuKlaimDetail 
            Caption         =   "Pelayanan Klaim Penjamin Detail"
         End
         Begin VB.Menu mnuLapPelayananCicilanPasien 
            Caption         =   "Pelayanan Cicilan Pasien"
         End
      End
      Begin VB.Menu mnuPengeluaranKasir 
         Caption         =   "Pengeluaran Kasir"
         Begin VB.Menu mnuLapReturPembayaran 
            Caption         =   "Retur Pembayaran"
         End
         Begin VB.Menu mnuLapPengeluaranUmum 
            Caption         =   "Pengeluaran Umum"
         End
         Begin VB.Menu mnuLapPelayananSupplier 
            Caption         =   "Pelayanan Supplier"
         End
      End
      Begin VB.Menu mnuRekapitulasiPendapatan 
         Caption         =   "Rekapitulasi Pendapatan"
      End
      Begin VB.Menu mnuRekapitulasiRemunerasi 
         Caption         =   "Rekapitulasi Remunerasi"
      End
      Begin VB.Menu MLaporanPendapatanPerDokterUnit 
         Caption         =   "Pendapatan Dokter --> Unit"
         Visible         =   0   'False
      End
      Begin VB.Menu MDaftarPenerimaan 
         Caption         =   "Penerimaan Kasir"
         Visible         =   0   'False
      End
      Begin VB.Menu MPenerimaanKasirUmum 
         Caption         =   "Penerimaan Kasir Umum"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLPKPP 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLPKPPD 
         Caption         =   "Penerimaan Kasir Per Pasien Detail"
         Visible         =   0   'False
      End
      Begin VB.Menu mnDetailRincianBiayaPelayananPasien 
         Caption         =   "Detail Rincian Biaya Pelayanan Pasien"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuseppkppd 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnupkpk2 
         Caption         =   "Penerimaan Kasir"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuknp 
         Caption         =   "Detail Penerimaan Kasir"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusepdpk 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnupkpd2 
         Caption         =   "Penerimaan Kasir Per Dokter"
         Visible         =   0   'False
      End
      Begin VB.Menu mnupkpk 
         Caption         =   "Penerimaan Kasir Per Komponen"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuseppkpp 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnudpybh 
         Caption         =   "Pasien yang Berhutang"
         Visible         =   0   'False
      End
      Begin VB.Menu mnupsk 
         Caption         =   "Pembatalan Struk Kuitansi"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusepPSK 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPenerimaanKasApotik 
         Caption         =   "Penerimaan Kas Apotik"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPengeluaranKasApotik 
         Caption         =   "Pengeluaran Kas Apotik"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusepKKA 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPenerimaanKasRS 
         Caption         =   "Penerimaan Kas RS"
         Visible         =   0   'False
      End
      Begin VB.Menu MPendapatanPerUnit 
         Caption         =   "Pendapatan Unit"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPendapatanPerUnit 
         Caption         =   "Detail Pendapatan Unit"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPengeluaranKasRS 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPendapatanRuangan 
         Caption         =   "Pendapatan Ruangan"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuseplpr 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MLaporanPendapatanPerRuanganUnit 
         Caption         =   "Pendapatan Ruangan --> Unit"
         Visible         =   0   'False
      End
      Begin VB.Menu MLaporanPendapatanPerUnitRuangan 
         Caption         =   "Pendapatan Unit --> Ruangan"
         Visible         =   0   'False
      End
      Begin VB.Menu MLaporanPerSMFDokter 
         Caption         =   "Pendapatan SMF --> Dokter"
         Visible         =   0   'False
      End
      Begin VB.Menu MPendapatanKelasUnit 
         Caption         =   "Pendapatan Kelas-->Unit"
         Visible         =   0   'False
      End
      Begin VB.Menu MLaporanPertindakanRuangan 
         Caption         =   "Pendapatan Tindakan --> Ruangan"
         Visible         =   0   'False
      End
      Begin VB.Menu MLaporanPerOARuangan 
         Caption         =   "Pendapatan Obat Alkes --> Ruangan"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRekapKlaim 
         Caption         =   "Rekap Klaim"
      End
   End
   Begin VB.Menu MWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu MCascade 
         Caption         =   "&Cascade"
      End
   End
   Begin VB.Menu mbantuan 
      Caption         =   "Ban&tuan"
      Begin VB.Menu mTentang 
         Caption         =   "Tentang Medifirst2000"
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "MDIUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim sepuh As Boolean

Private Sub frmDaftarPenjualanApotik_Click()
   frmDaftarPenjualanTanpaBayar.Show
End Sub

Private Sub MCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mCetakLembarInput_Click()
    mstrKdKelompokBarang = "02"
    frmDaftarCetakInputStokOpname.Show
End Sub

Private Sub MDaftarDataPenerimaanKasirSudahDiPosting_Click()
    frmDaftarDataPenerimaanKasirHavePostingOATMApotik.Show
End Sub

Private Sub MDaftarKasirUntukClaim_Click()
    frmDaftarPasienForClaim.Show
End Sub

Private Sub MDaftarPasienUntukVerifikasi_Click()
    frmVerifikasiData.Show
End Sub

Private Sub MDaftarPembayaranDepositBiayaPerawatan_Click()
    frmDaftarPembayaranDepositBiayaPerawatan.Show
End Sub

Private Sub MDaftarPenerimaan_Click()
    frmDaftarPenerimaanKasir.Show
End Sub

Private Sub MDaftarReturStruk_Click()
    frmDaftarReturStruk.Show
End Sub

Private Sub MDIForm_Load()
    blnForm = False
    strSQL = "SELECT * FROM DataPegawai WHERE IdPegawai = '" & strIDPegawaiAktif & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    strNmPegawai = rs.Fields("NamaLengkap").Value
    Set rs = Nothing
    StatusBar1.Panels(1).Text = "Nama User : " & strNmPegawai
    StatusBar1.Panels(2).Text = "Nama Ruangan : " & mstrNamaRuangan
    StatusBar1.Panels(5).Text = "Nama Komputer : " & strNamaHostLocal
    StatusBar1.Panels(6).Text = "Server : " & strServerName & " (" & strDatabaseName & ")"
    mnlogout.Caption = "Log Off..." & strNmPegawai

    strSQL = "SELECT TerminBayarFakturSupplier, PersentasePpn, PersentaseLimitDiscount, PersentaseJasaPenulisResep, BiayaAdministrasi " & _
    " From SettingDataPendukung" & _
    " WHERE (KdInstalasi = '" & mstrKdInstalasiLogin & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        typSettingDataPendukung.intTerminBayarFaktur = 0
        typSettingDataPendukung.realJasaPenulisResep = 0
        typSettingDataPendukung.realLimitDiscount = 0
        typSettingDataPendukung.realPPn = 0
        typSettingDataPendukung.curBiayaAdministrasi = 0
    Else
        typSettingDataPendukung.intTerminBayarFaktur = rs("TerminBayarFakturSupplier").Value
        typSettingDataPendukung.realJasaPenulisResep = rs("PersentaseJasaPenulisResep").Value
        typSettingDataPendukung.realLimitDiscount = rs("PersentaseLimitDiscount").Value
        typSettingDataPendukung.realPPn = rs("PersentasePpn").Value
        typSettingDataPendukung.curBiayaAdministrasi = rs("BiayaAdministrasi").Value
    End If

    strSQL = "SELECT JmlPembulatanHarga, JumlahBAdminOAPerBaris, JumlahBAdminTMPerHari FROM MasterDataPendukung"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        typSettingDataPendukung.intJmlPembulatanHarga = 0
        typSettingDataPendukung.intJumlahBAdminOAPerBaris = 0
        typSettingDataPendukung.intJumlahBAdminTMPerHari = 0
    Else
        typSettingDataPendukung.intJmlPembulatanHarga = dbRst(0)
        typSettingDataPendukung.intJumlahBAdminOAPerBaris = dbRst(1)
        typSettingDataPendukung.intJumlahBAdminTMPerHari = dbRst(2)
    End If
    
    strSQL = "Select MetodeStokBarang From SuratKeputusanRuleRS where statusenabled=1"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        bolStatusFIFO = False
    Else
        If dbRst("MetodeStokBarang") = 0 Then
            bolStatusFIFO = False
        Else
            bolStatusFIFO = True
        End If
    End If
    
    strSQL = "Select MetodeHargaBarang From SuratKeputusanRuleRS where statusenabled=1"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        bolStatusHarga = False
    Else
        If dbRst("MetodeHargaBarang") = 0 Then
            bolStatusHarga = False
        Else
            bolStatusHarga = True
        End If
    End If
End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Exit Sub
    PopupMenu mnudata
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim q As String
    If sepuh = True Then
        q = MsgBox("Log Off user " & strNmPegawai & " ", vbQuestion + vbOKCancel, "Konfirmasi")
        If q = 2 Then
            Unload frmLogin
            Cancel = 1
        Else
            Cancel = 0
            frmLogin.Show
        End If
        sepuh = False
    Else
        q = MsgBox("Tutup aplikasi ", vbQuestion + vbOKCancel, "Konfirmasi")
        If q = 2 Then

            Unload frmLogin
            Cancel = 1
        Else
            dTglLogout = Now
            Call subSp_HistoryLoginAplikasi("U")
            Cancel = 0
        End If
    End If
End Sub

Private Sub MEditTagihanPasien_Click()
    frmTagihanPasienEdit.Show
End Sub

Private Sub MEditTanggunganPenjamin_Click()
    Call frmTagihanPasien.subLoadDataKomponenPel
End Sub

Private Sub MEditTanggunganPenjaminVerifikasi_Click()
    Call frmTagihanPasienVerifikasi.subLoadDataKomponenPel
End Sub

Private Sub mGantiKataKunci_Click()
    frmLoginEditAccount.Show
End Sub

Private Sub MInformasiPemesananBarang_Click()
    mstrKdKelompokBarang = "02"
    frmInfoPesanBarang.Show
End Sub

Private Sub MInformasiTarifPelayanan_Click()
    frmInformasiTarifPelayanan.Show
End Sub

Private Sub mInputStokOpn_Click()
    mstrKdKelompokBarang = "02"
    frmStokOpname.Show
End Sub

Private Sub MKasirUmum_Click()
    frmBayarUmum.Show
End Sub

Private Sub MKonversiKomponenUnitKePelayanan_Click()
'    frmKonversiKomponenUnitKePelayanan.Show
    frmKonversiKomponenLaporanToPelayanan.Show
End Sub

Private Sub MLaporanPendapatanPerDokterUnit_Click()
    frmDaftarPendapatanPerDokter_Header.Show
End Sub

Private Sub MLaporanPendapatanPerRuanganUnit_Click()
    frmDaftarPendapatanPerRuangan_Header.Show
End Sub

Private Sub MLaporanPendapatanPerUnitRuangan_Click()
    frmDaftarPendapatanPerUnit_Header.Show
End Sub

Private Sub MLaporanPerOARuangan_Click()
    frmDaftarPendapatanPerObatAlkes_Header.Show
End Sub

Private Sub MLaporanPerSMFDokter_Click()
    frmDaftarPendapatanPerSMF_Dokter.Show
End Sub

Private Sub MLaporanPertindakanRuangan_Click()
    frmDaftarPendapatanPerTindakanKhusus_Header.Show
End Sub

Private Sub mLapSaldoBarang_Click()
    frmLaporanSaldoBarangMedis_v3.Show
End Sub

Private Sub mnCetakLembarInputNM_Click()
    mstrKdKelompokBarang = "01" 'non medis
    frmDaftarCetakInputStokOpnameNM.Show
End Sub

Private Sub mnDetailRincianBiayaPelayananPasien_Click()
    frmCetakDetailRincianBiayaPelayanan.Show
End Sub

Private Sub mNilaiPersediaan_Click()
    mstrKdKelompokBarang = "02"
    frmNilaiPersediaan.Show
End Sub

Private Sub mnInfoPesanBarangNM_Click()
    mstrKdKelompokBarang = "01" 'non medis
    frmInfoPesanBarangNM.Show
End Sub

Private Sub mnInputStokOpname_Click()
    mstrKdKelompokBarang = "01" 'non medis
    frmStokOpnameNM.Show
End Sub

Private Sub mnKondisiBarangNM_Click()
    frmKondisiBarangNM.Show
End Sub

Private Sub mnLapSaldoBarangNM_Click()
    frmLaporanSaldoBarangNM_v3.Show
End Sub

Private Sub mnlogout_Click()
    Dim adoCommand As New ADODB.Command
    openConnection
    sepuh = True
    strQuery = "UPDATE Login SET Status = '0' " & _
    "WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
    adoCommand.ActiveConnection = dbConn
    adoCommand.CommandText = strQuery
    adoCommand.CommandType = adCmdText
    adoCommand.Execute

    dTglLogout = Now
    Call subSp_HistoryLoginAplikasi("U")
    Unload Me
End Sub

Private Sub mnMutasiBarangNM_Click()
    frmMutasiBarangNM.Show
End Sub

Private Sub mnNilaiPersediaan_Click()
    mstrKdKelompokBarang = "01"
    frmNilaiPersediaanNM.Show
End Sub

Private Sub mnPemesananBarang_Click()
    frmPemesananBarang.Show
End Sub

Private Sub mnRekapTransBrgNM_Click()
    mstrKdKelompokBarang = "01" 'non medis
    frmDataTransaksiBarangNM.Show
End Sub

Private Sub mnSelesai_Click()
    Dim pesan As VbMsgBoxResult
    Dim adoCommand As New ADODB.Command
    pesan = MsgBox("Tutup aplikasi ", vbQuestion + vbYesNo, "Konfirmasi")
    If pesan = vbYes Then

        openConnection
        strQuery = "UPDATE Login SET Status = '0' " & _
        "WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
        adoCommand.ActiveConnection = dbConn
        adoCommand.CommandText = strQuery
        adoCommand.CommandType = adCmdText
        adoCommand.Execute

        dTglLogout = Now
        Call subSp_HistoryLoginAplikasi("U")
        End
    Else
    End If
End Sub

Private Sub mnStokBarangNM_Click()
    frmStokBarangNonMedis.Show
End Sub

Private Sub mnuDafPembayaranSupplier_Click()
    frmDaftarPembayaranTagihanSupplier.Show
End Sub

Private Sub mnuDafPengeluaranUmum_Click()
 frmDaftarPengeluaranKasirUmum.Show
End Sub

Private Sub mnuDaftarPasienDeposit_Click()
    frmDaftarPembayaranDepositBiayaPerawatan.Show
End Sub

Private Sub mnuDaftarPasienKlaim_Click()
    frmDaftarPasienForClaim.Show
End Sub

Private Sub mnuDaftarPasienKredit_Click()
    frmDaftarPasienBayarKredit.Show
End Sub

Private Sub mnuDaftarPasienMshAktif_Click()
    frmBiayaDaftarPasienBelumPulang.Show
End Sub

Private Sub mnuDaftarPasienPulang_Click()
    frmCariPasien.Show
End Sub

Private Sub mnuDaftarPasienReturPembayaran_Click()
    frmDaftarReturStruk.Show
End Sub

Private Sub mnuDaftarPasienSudahBayar_Click()
    frmDaftarPasienSudahBayar.Show
End Sub

Private Sub mnuDaftarPasienygBelumDanSudahPernahBayar_Click()
    frmDaftarPasienYgBelumDanSudahBayar.Show
End Sub

Private Sub mnuDaftarTagihanSupplier_Click()
    frmDaftarTagihanSupplier.Show
End Sub

Private Sub mnuEditKomponenTarif_Click()
    frmTindakan.fraUpdateKomponenTarif.Visible = True
End Sub

Private Sub mnuIndoTindaknonTanggungan_Click()
    frmDaftarTMNonTanggungan.Show
End Sub

Private Sub mnuInfoDaftarBatalKuitansi_Click()
    frmDaftarReturStrukInfo.Show
End Sub

Private Sub mnuInfoJenisPelayanan_Click()
    frmDafPelMed.Show
End Sub

Private Sub mnuInfoKelasPel_Click()
    frmKelPelMed.Show
End Sub

Private Sub mnuInfoKonversi_Click()
    frmKonversiPaskeAsuransi.Show
End Sub

Private Sub mnuInfoNonPaket_Click()
    FrmTanggunganAsuransiNonPaket.Show
End Sub

Private Sub mnuInfoPaketPelayananPenjamin_Click()
    frmdatapendukungpenjaminpaket.Show
End Sub

Private Sub mnuInfoPerTindakan_Click()
    FrmTanggunganPaketAsuransiPerTindakan.Show
End Sub

Private Sub mnuInformasiTarifPelayanan_Click()
    frmInformasiTarifPelayanan.Show
End Sub

Private Sub mnuInforTarifPaketPenjamin_Click()
    FrmTanggunganPaketAsuransi.Show
End Sub

Private Sub mnuInfoTarifPel_Click()
    frmTarifPelMedikBaru.Show
End Sub

Private Sub mnuKelompokPenjamin_Click()
    frmdatapendukungpenjamin.Show
End Sub

Private Sub mnuKlaimDetail_Click()
    frmDaftarPenerimaanKasirPelayananKlaimPenjaminDetail.Show
End Sub

Private Sub mnuLapPelayananCicilanPasien_Click()
    frmDaftarPenerimaanKasirPelayananKreditPasien.Show
End Sub

Private Sub mnuLapPelayananDepositPasien_Click()
    frmDaftarPenerimaanKasirDeposit.Show
End Sub

Private Sub mnuLapPelayananKlaimPenjamin_Click()
    frmDaftarPenerimaanKasirPelayananKlaimPenjamin.Show
End Sub

Private Sub mnuLapPelayananPasien_Click()
    frmDaftarPenerimaanKasirPelayanan.Show
End Sub

Private Sub mnuLapPelayananSupplier_Click()
    frmDaftarPengeluaranKasirPelayananSupplier.Show
End Sub

Private Sub mnuLapPelayananUmum_Click()
    frmDaftarPenerimaanKasirUmum.Show
End Sub

Private Sub mnuLapPengeluaranUmum_Click()
    frmDaftarPengeluaranKasirUmum.Show
End Sub

Private Sub mnuLapReturPembayaran_Click()
    frmLapPengeluaranKasirPelayananReturPembayaran.Show
End Sub

Private Sub mnuPelayananCicilanPasien_Click()
    frBayarHutangPasien.Show
End Sub

Private Sub mnuPelayananDepositPasien2_Click()
  frmDaftarPasienAktif.Show
End Sub

Private Sub mnuPelayananKlaimPenjamin_Click()
    frmDaftarPasienForClaim.Show
End Sub

Private Sub mnuPelayananPasien_Click()
    blnForm = True
    frmTagihanPasien.Show
End Sub

Private Sub mnuPelayananRuangan_Click()
    frmDaftarPenerimaanKasirPelayananNew.Show
End Sub

Private Sub mnuPelayananUmum_Click()
    frmBayarUmum.Show
End Sub

Private Sub mnuPemakaianBahandanAlat_Click()
    frmPemakaianBahanAlat.Show
End Sub

Private Sub mnupembayaranapotik_Click()
 frmDaftarPenjualan.Show
End Sub

Private Sub mnuPembayaranOtomatis_Click()
    frmPembayaranOtomatis.Show
End Sub

Private Sub mnuPembayaranPenjaminPasien_Click()
    frmDaftarPembayaranPenjamin.Show
End Sub

Private Sub mnuPendapatanPerUnit_Click()
    frmLapDetailPendapatanPerUnit.Show
End Sub

Private Sub mnuPendapatanRuangan_Click()
    frmLapDetailPendapatanRuangan.Show
End Sub

Private Sub mnuPenerimaanUmum_Click()
  frmDaftarPenerimaanKasirUmum.Show
End Sub

Private Sub mnuPengeluaranReturApotik_Click()
frmReturPenjualanApotik.Show
End Sub

Private Sub mnuPengeluaranSupplier_Click()
    frmBayarTagihanSupplier_Aja.Show
End Sub

Private Sub mnuPengeluaranUmum_Click()
    frmPengeluaranStrukKas.Show
End Sub

Private Sub mnuPositngDataKasirPendapatan_Click()
    frmPostingDataKasirPendapatan.Show
End Sub

Private Sub mnuPostingDataPendapatanRemunerasi_Click()
    frmPostingDataKasirRemunerasiPendapatan.Show
End Sub

Private Sub mnuPostingDataRemunerasi_Click()
    frmPostingDataKasirRemunerasi.Show
End Sub

Private Sub mnuProfilPenjamin_Click()
    frmpenjamin.Show
End Sub

Private Sub mnuRekapitulasiPendapatan_Click()
'    strCetak = "Pendapatan"
'    frmLapRekapitulasiPendapatanRemunerasi.Show
    frmPendapatanRuangan.Show
End Sub

Private Sub mnuRekapitulasiRemunerasi_Click()
    strCetak = "Remunerasi"
    frmLapRekapitulasiPendapatanRemunerasi.Show
End Sub

Private Sub mnuReturApotik_Click()
    frmDaftarReturStrukApotik.Show
End Sub

Private Sub mnuRekapKlaim_Click()
    frmRekapClaim.Show
End Sub

Private Sub mnuReturPelayananApotik_Click()
    frmDaftarReturStrukApotik.Show
End Sub

Private Sub mnuReturPembayaran_Click()
    frmReturStrukPelayananPasien.Show
End Sub

Private Sub mnuReturTagihanPerPelayanan_Click()
    frmReturTagihanPasien.Show
End Sub

Private Sub mnuSisaTagihan_Click()
    frmDaftarPasienForST2.Show
End Sub

Private Sub mnuTagihanPasien_Click()
    frmTagihanPasienEdit.Show
End Sub

Private Sub mnuTerimaBarangLangsung_Click()
    frmTerimaBarangLangsung.Show
End Sub

Private Sub mnuTRS_Click()
    frmDaftarPasienForClaimTRS.Show
End Sub

Private Sub mnuValidasiUlangYangGagalBayar_Click()
    frmValidasiUlangGagalBayar.Show
End Sub

Private Sub mnuVerifikasiBayar_Click()
    frmValidasiDataBayar.Show
End Sub

Private Sub mnuVerPemakaianAsuransi_Click()
    frmVerifikasiDataAsuransiPasien.Show
End Sub

Private Sub MPembayaranDepositBiayaPerawatan_Click()
    frmBayarDeposit.Show
End Sub

Private Sub MPembayaranOtomatis_Click()
    frmPembayaranOtomatis.Show
End Sub

Private Sub MPendapatanKelasUnit_Click()
    frmDaftarPendapatanPerKelas_Header.Show
End Sub

Private Sub MPendapatanPerUnit_Click()
    frmDaftarPendapatanPerUnit.Show
End Sub

Private Sub MPenerimaanKasirUmum_Click()
    frmDaftarPenerimaanKasirUmum.Show
End Sub

Private Sub MPostingDataKasirPenerimaan_Click()
    frmPostingDataKasirPenerimaan.Show
End Sub

Private Sub mStokBarang_Click()
    frmStokBrg.Show
End Sub

Private Sub MTagihanPasienApotik_Click()
    frmStrukBuktiKasMasuk.Show
End Sub

Private Sub mnucdp_Click()
    frmCariPasien.Show
End Sub

Private Sub mnudaftarKasMasuk_Click()
    frmDaftarKasMasuk.Show
End Sub

Private Sub mnudpsb_Click()
    frmDaftarPasienSudahBayar.Show
End Sub

Private Sub mnuDReturStruk_Click()
    frmReturStrukPelayananPasien.Show
End Sub

Private Sub mnuDataTagihanPasienApotik_Click()
    frmDaftarTagihanPasienApotik.Show
End Sub

Private Sub mnuCariDataTagihanSupplier_Click()
    frmDaftarTagihanSupplier.Show
End Sub

Private Sub mnuIDaftarPasienNgutang_Click()
    frmDaftarPasienBayarKredit.Show
End Sub

Private Sub mnuKasKeluar_Click()
    frmDaftarKasKeluar.Show
End Sub

Private Sub mnuPengeluaranKas_Click()
    mstrLaporan = "LaporanKasKeluar"
    frmRekapLaporanKas.Show
End Sub

Private Sub mnuknp_Click()
    mstrLaporan = "PenerimaanDetail"
    frmRekapLaporan.Show
    frmRekapLaporan.Caption = "Medifirst2000 - Detail Penerimaan Kasir"
End Sub

Private Sub mnuLPKPP_Click()
    vLaporan = ""
    mstrCetak = "PenerimaanKasirPerPasienTDetail"
    frmPenerimaanKasir.Show
End Sub

Private Sub mnuLPKPPD_Click()
    vLaporan = ""
    mstrCetak = "PenerimaanKasirPerPasienDetail"
    frmPenerimaanKasir.Show
End Sub

Private Sub mnuPenerimaanKas_Click()
    mstrLaporan = "LaporanKasMasuk"
    frmRekapLaporanKas.Show
End Sub

Private Sub mnupkpd2_Click()
    mstrLaporan = "PenerimaanDokter"
    frmRekapLaporanKasirPerDokter.Show
End Sub

Private Sub mnupkpk_Click()
    frmRekapLaporanKomp.Show
End Sub

Private Sub mnupkpk2_Click()
    mstrLaporan = "PenerimaanPerda"
    frmRekapLaporan.Show
    frmRekapLaporan.Caption = "Medifirst2000 - Penerimaan Kasir"
End Sub

Private Sub mnupsk_Click()
    mstrLaporan = "PembatalanStruk"
    frmRekapLaporan.Show
    frmRekapLaporan.Caption = "Medifirst2000 - Pembatalan (Retur) Struk"
End Sub

Private Sub mnutp_Click()
    frmTagihanPasien.Show
End Sub

Private Sub mSettingPrinter_Click()
    frmSetupPrinter3.Show
End Sub

Private Sub mTentang_Click()
    frmAbout.Show
End Sub

Private Sub POAKaryawan_Click()
    frmDaftarPakaiAlkesKaryawan.Show
End Sub

