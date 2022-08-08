Attribute VB_Name = "modVar"
Public bolEditDiagnosa As Boolean
Public mstrKdDiagnosa As String
Public mstrKdJenisDiagnosaTindakan As String

Public KdInstalasiDiagnosa As String
Public KdRuanganDiagnosa As String
Public KdInstalasiDokter As String
Public bolinacbgs As Boolean

Public strsql503 As String
Public rs503 As ADODB.recordset
Public strsqlA As String
Public strsqlB As String
Public strsqlC As String
Public strsqlD As String
Public strsqlE As String

Public statusTagihan As String
Public strStsPasien As String
Public strNmPegawai As String
Public strKdSubInstalasi As String
Public strSQLIdentifikasi As String
Public strUserName As String
Public strPassword As String
Public strStatusPembayaran As String
Public intPembayaranPeritem As Integer
Public subFilterDataBuatLaporan As String

Public intTahun As Integer
Public intBulan As Integer
Public intTanggal As Integer
Public blnMeTglLahEdi As Boolean
Public blnTambah As Boolean
Public blndokter As Boolean
Public Periode As String
Public BlnAwal As String
Public BlnAkhir As String
Public ThnAwal As String
Public ThnAkhir As String
Public TglAwal As String
Public TglAkhir As String
Public strjudul As String
Public strKLaimPenjamin As String
Public mstrNoRetur As String
Public strNamaHostLocal As String
Public strKdAplikasi As String
Public dTglLogin As Date
Public dTglLogout As Date
Public blnValidasiInput As Boolean
Public vCetakLaporan As String

Public strKdRuanganBayar As String
Public bolDatavalid As Boolean
Public jData As Integer
Public jLopKredit As Integer
Public BoolKembali As Boolean
Public dNTglBerlaku As Date
Public strNStsCITO As String

Public strPathLogo As String
Public strNamaFileLogoRS As String
Public strNNamaRS As String
Public strNAlamatRS As String
Public strNTeleponRS As String
Public strNKotaRS As String
Public strNKodepos As String
Public strWebsite As String
Public strEmail As String
Public strkelasRS As String
Public strketkelasRS As String
Public strPropinsi As String
Public blnError As Boolean

Public strNKdJenisTarif As String
Public strNKdJenisTarif2 As String
Public mstrKdInstalasi As String
Public mstrNamaInstalasi As String
Public kdsubinstalasi As String
Public idpegawai As String
Public strNIdPejamin As String
Public blnStatusAsuransi As Boolean
Public strNKdRuangan As String
Public strNNamaRuangan As String

Public strCetak As String
Public strCetak2 As String
Public strCetak3 As String
Public mstrCetak2 As String

Public mstrKdSubInstalasi As String
Public mstrKdRuangan As String
Public mstrNamaRuangan As String
Public mstrKdRuanganPasien As String
Public mblnPsnMati As Boolean
Public mstrFilterDokter As String
Public mstrFilterSupir As String
Public mintJmlDokter As Integer
Public mintRowNow As Double
Public mdTglBerlaku As Date
Public mstrKdKelas As String
Public mdTglMasuk As Date
Public fStatusPiutang As String
Public fStatusBayarSemua As String
Public strDaftarRuangan As String
Public strDaftarPenjamin As String
Public sJenisNilai As String
Public sDetailJenisLaporan As String
Public sJenisTransksi As String
Public sJenisLaporan As String
Public sKdJenisPasien As String
Public sIdPenjamin As String
Public sKdKomponenLaporan As String
Public sIdDokterPemeriksa As String
Public sKdKomponen As String
Public sTempGrup  As String
Public sKdTempGrup  As String
Public tempPenjamin As String
Public sUnitLaporan As String
Public sDokterPemeriksa As String
Public sJenisPasien As String
Public sPenjaminPasien As String
Public skomponenTarif As String

Public mbolCetakJasaDokter As Boolean

Public mstrNamaBarang As String
Public strPasien As String
Public mblnForm As Boolean
Public cetak As String
Type typeBarang
    strkdbarang As String
    strNamaBarang As String
    strKdAsal As String
    intJmlTerkecil As Double
    intJmlTempTotal As Double
    intJmlJualTerkecil As Double
End Type

Public mblnFormDaftarPasienRI As Boolean

Public typBarang() As typeBarang
Public ctk As String

Public mblnAdmin As Boolean
Public mstrKdKategoryUser As String
Public mblnVerifikator As Boolean

Public mstrLaporan As String
Public mstrKdDokter As String
Public mstrNamaDokter As String
Public strKodePelayananRS As String
Public mstrKdJenisLaporan As String
Public mstrStatusBayar As String
Public mstrNamaKomponenTarif As String
Public mstrFilterData As String
Public mstrFilterData2 As String
Public mstrFilterData3 As String
Public mstrNoLabRad As String

Public mstrKdJenisPasien As String
Public mstrNamaJenisPasien As String

Public mstrKdPenjaminPasien As String
Public mstrNamaPenjaminPasien As String
Public mstrFilter As String
Public mstrKriteria As String
Public mstrNoOrder As String
Public mstrGroup As String
Public mstrCetak As String
Public mstrNama As String
Public mstrIsiGroup As String

Public mstrKdKelompokBarang As String
Public mstrKdInstalasiNonMedis As String

Public editpoli As Boolean
Public grafikkelompok As Boolean
Public grafikklinik As Boolean
Public nike As Boolean
Public noidpegawai As String
Public petugas As String
Public strsqlx2 As String
Public strsqlx3 As String
Public strsqlx4 As String

Public strsqlx5 As String
Public strsqlx6 As String

Public blnAdmin As Boolean
Public strUser As String
Public strPass As String
Public strPass2 As String
Public strPassEn As String
Public strStatus As String
Public intLenUser As Integer
Public strIDPegawai As String
Public strIDPegawaiAktif As String
Public tgl As String
Public varcounter As Boolean
Public darah As Boolean
Public alamat As Boolean
Public dadang As Boolean
Public dadang1 As Boolean
Public mstrNoCM2 As String
Public strSQL As String

Public strSQLx As String
Public strSQLRem As String
Public blnEditPoli As Boolean
Public strJK As String
Public strNoStruk() As String
Public VarNoStrukCek As String
Public VarNoPendaftaran As String
Public varNoCMCek As String
Public varNamaPasienCek As String
Public varAlamatCek As String
Public varJmlBayarCek As Currency
Public varTglBKMCek As Date
Public varRuangCek As String
Public varInstalasiCek As String
Public varTglMasukCek As Date
Public varTglKeluarCek As Date
Public varAnPasienCek As String
Public varKelompokPasienCek As String
Public varPenjaminCek As String
Public intJmlStruk As Integer

'Kode Aplikasi yang sedang dijalankan, ganti sesuai keperluan
'**************************************
Public Const intAplikasi = 1
'**************************************

'variabel global koneksi & recordset ke db

'Hilangkan bila telah dideklarasikan sebelumnya
'**************************************
'Public dbConn As New ADODB.Connection
'**************************************

Public rsPegawai As New ADODB.recordset
Public rsPegawaiCount As New ADODB.recordset
Public rslogin As New ADODB.recordset
Public rsLoginApp As New ADODB.recordset
Public rsAplikasi As New ADODB.recordset
Public rsAplikasiCount As New ADODB.recordset
Public rsLoginCompare As New ADODB.recordset
Public strQuery As String

Public rsF As New ADODB.recordset
Public rsG As New ADODB.recordset
Public rsH As New ADODB.recordset
Public rsI As New ADODB.recordset
Public rsJ As New ADODB.recordset
Public rsK As New ADODB.recordset
Public rsL As New ADODB.recordset
Public rsM As New ADODB.recordset
Public rsE As New ADODB.recordset
Public rsD As New ADODB.recordset
Public rsB As New ADODB.recordset
Public rsC As New ADODB.recordset
Public rs As New ADODB.recordset
Public rsCek As New ADODB.recordset
Public dbrs As New ADODB.recordset
Public adors As New ADODB.recordset
Public dbRst As New ADODB.recordset
Public dbConn As New ADODB.connection 'dipakai hampir disemua modul
Public dbcmd As New ADODB.Command
Public dbCmdSubReport As New ADODB.Command
Public adoComm As New ADODB.Command
Public dmParam As New ADODB.Parameter
Public dbcomm As New ADODB.Command

'Public querystring As String
Public crxReport As CRAXDDRT.Report         'dipakai untuk modul crystal reports
'
'variabel servername, databasename, namarumahsakit dibuat global,
'karena nilai dari variabel2 ini akan dipakai untuk fungsi getsetting.
'
Public mstrNoPen As String
Public mstrNoCM As String
Public mstrTglKeluar As Date
Public strNamaRuangan As String
Public instalasi As Integer
Public KdInstalasi As String
Public SubInstalasi As Integer
Public ServerName As String
Public DatabaseName As String
Public UserID As String
Public UserName As String
Public NamaRumahSakit As String

Public NoCM As String
Public NamaPasien As String
Public isFindFirst As Boolean
Public enableEdit As Boolean
Public message As String
Public Umur As udt_umur
Public prmTgl As String
Public continue As Boolean
Public queryString As String
Public JenisKelamin As String
Public KodeInstalasi As String
Public KodeSubInstalasi As String
Public DataString As String

Public TtlHari As Integer

Type udt_umur
    tahun As Integer
    bulan As Integer
    hari As Integer
End Type

Public mdTglAwal As Date
Public mdTglAkhir As Date
Public mblnGrafik As Boolean
Public strDatabaseName As String
Public strServerName As String
Public YOC_intYear As Integer
Public YOC_intMonth As Integer
Public YOC_intDay As Integer

'setting printer
Public prn As Printer
Public sizepaper As CRPaperSize
Public duplexpaper As CRPrinterDuplexType
Public sPrinter As String, sDriver As String, sUkuranKertas As String
Public sPrinter2 As String, sDriver2 As String, sUkuranKertas2 As String
Public sPrinter3 As String, sDriver3 As String, sUkuranKertas3 As String
Public sPrinter4 As String, sDriver4 As String, sUkuranKertas4 As String
Public sPrinter5 As String, sDriver5 As String, sUkuranKertas5 As String
Public intTimerPrinter As Integer
Public sDuplex As String, sOrientasKertas As String
Public tmpOrien As String
'setting printer

Public mcurAll_TBP As Currency
Public mcurAll_TP As Currency
Public mcurAll_TRS As Currency
Public mcurAll_Pemb As Currency
Public mcurAll_HrsDibyr As Currency
Public mcurTM_TBP As Currency
Public mcurTM_TP As Currency
Public mcurTM_TRS As Currency
Public mcurTM_Pemb As Currency
Public mcurTM_Discount As Currency
Public mcurTM_HrsDibyr As Currency
Public mcurTM_HrsDibyrNow As Currency
Public mcurTM_JmlByr As Currency
Public mcurTM_ST As Currency
Public mcurOA_TBP As Currency
Public mcurOA_TP As Currency
Public mcurOA_TRS As Currency
Public mcurOA_Pemb As Currency
Public mcurOA_Discount As Currency
Public mcurOA_HrsDibyr As Currency
Public mcurOA_HrsDibyrNow As Currency
Public mcurOA_JmlByr As Currency
Public mcurOA_ST As Currency
Public mblnTM As Boolean
Public mblnOA As Boolean
Public mblnAP As Boolean

Public mcurAP_TBP As Currency
Public mcurAP_TP As Currency
Public mcurAP_TRS As Currency
Public mcurAP_Pemb As Currency
Public mcurAP_Discount As Currency
Public mcurAP_HrsDibyr As Currency
Public mcurAP_HrsDibyrNow As Currency
Public mcurAP_JmlByr As Currency
Public mcurAP_ST As Currency

Public mstrKdPenjamin As String
Public mcurBayar As Currency
Public mcurBayarTM As Currency
Public mcurBayarOA As Currency
Public mcurPembebasan As Currency
Public mstrNoStruk As String

Public strRegistrasi As String
Public blnCariPasien As Boolean
Public intJmlDokter  As Integer
Public mblnFormDaftarPasienIGD As Boolean
Public mblnFormDaftarAntrian As Boolean
Public vLaporan As String
Public mstrKdRuanganORS As String
Public mstrNoIBS As String
Public mblnTP As String
Public mstrKdJenisOperasi As String
Public mstrJenisOperasi  As String
Public subTanggalTerakhir As Integer
Public mblnOperator As Boolean
Public mstrKdInstalasiLogin As String
Public mstrNoHasilLab As String
Public mstrNoBKM As String
Public varNoBKMCek As String
Public mstrNamaRuanganPerujuk As String
Public mstrKdInstalasiPerujuk As String
Public mstrNamaKelas As String
Public mstrNoLab  As String
Public mblnKonsul As Boolean
Public dTglRujukan As Date
Public mstrNoBKK As String
Public mstrNoRad As String
Public mstrJenisPasien As String
Public mstrRuanganPerujuk As String
Public mstrNoHasilRad As String
Public mblnCariPasien As Boolean
Public strNHargaSatuan As Double
Public strNTotal As Double
Public mstrServerPrinterBarcode As String
Public miRowNow As Double
Public mstrNamaRuanganPasien As String
Public mstrKelas As String
Public blnFrmCariPasien As Boolean
Public mstrKdRuanganKasir As String
Public vJudul As String

Type PenjaminSisaTagihan
    strNamaLengkap As String
    dTglLahir As Date
    strJenisKelamin As String
    strNoIdentitas As String
    strHubungan As String
    strAlamat As String
    strTelepon As String
    strPropinsi As String
    strKota As String
    strKecamatan As String
    strKelurahan As String
    strRTRW As String
    strKodePos As String
    blnStatus As Boolean
End Type

Public typPenjaminSisaTagihan As PenjaminSisaTagihan
Public typPenjaminSisaTagihanApotik As PenjaminSisaTagihan
Public mblnAdaPlynTdkDibyr As Boolean
Public mcurDiscount As Currency
Public mblnTindakanKasir As Boolean
Public mLapPerParameter As String

Type typeSettingDataPendukung
    intTerminBayarFaktur As Integer
    realPPn As Double
    realLimitDiscount As Double
    realJasaPenulisResep As Double
    intJmlPembulatanHarga As Integer
    intJumlahBAdminOAPerBaris As Integer
    intJumlahBAdminTMPerHari As Integer
    curBiayaAdministrasi As Currency
End Type

Public typSettingDataPendukung As typeSettingDataPendukung
Public mstrValid As String
Public mstrNoTerima As String

Type Asuransi
    strIdPenjamin As String
    strIdAsuransi As String
    strIDPerusahaan As String
    strNoCm As String
    strNamaPeserta As String
    strIdPeserta As String
    strPerusahaanPenjamin As String
    strKdGolongan As String
    strKdKelasDitanggung As String
    dTglLahir As Date
    strAlamat As String
    strNoPendaftaran As String
    strHubungan As String

    strNoSJP As String
    dTglSJP As Date
    strNoBp As String
    intNoKunjungan As Integer

    strStatusNoSJP As String
    intAnakKe As Integer
    strUnitBagian As String
    strKdPaket As String

    strNoRujukan As String
    strKdRujukanAsal As String
    strDetailRujukanAsal As String
    strKdDetailRujukanAsal As String
    strNamaPerujuk As String

    dTglDirujuk As String
    strDiagnosaRujukan As String
    strKdDiagnosa As String

    blnSuksesAsuransi As Boolean
    strKdKelompokPasien As String
End Type

Public typAsuransi As Asuransi
Public mstrFormPengirim As String
Public mstrNoKirim As String
Public mdAwal As Date
Public mdAkhir As Date
Public mSup As String
Public mblnFormDaftarPasienRJ As Boolean
Public DbRec As New ADODB.recordset

Public strKdKelompokPasien As String
Public mstrKdKelompokPasien As String
Public mstrNoSJP As String

Public mintJmlBarisGrafik As Integer ' number of rows needed in Chart data
Public mintJmlKolomGrafik As Integer ' number of colomns needed in Chart data
Public arrGrafik() ' Chart data
Public JnsKriteria() ' criteria
Public mstrGrafik As String

Public mstrKdJenisTarif As String
Public mstrJenisTarif As String
Public intJmlPelayanan As Integer
Public intJmlPelayananAhkir As Integer

Type TypePembayaranNonPaket
    curHarga As Currency
    curTanggunganPjm As Currency
    curCostSharing As Currency
    curTanggunganRS As Currency
    curHrsDibyrPsn As Currency
End Type

Type TypePembayaranUmum
    curHrsDibyrPsn As Currency
End Type

Type TypePembayaranPaket
    curHarga As Currency
    curTanggunganPjm As Currency
    curCostSharing As Currency
    curTanggunganRS As Currency
    curHrsDibyrPsn As Currency
    strKdPaket As String
    strKdJnsPelayanan As String
End Type

Public mintJmlPktPlyn As Integer
Public TypByrPkt() As TypePembayaranPaket
Public TypByrPktTotal As TypePembayaranNonPaket
Public TypByrNonPkt As TypePembayaranNonPaket
Public TypByrUmum As TypePembayaranUmum
Public TypByrAll As TypePembayaranNonPaket
Public TypByrOA As TypePembayaranNonPaket
Public TypByrTM As TypePembayaranNonPaket

Public mstrKriteriaLaporan As String
Public intRowNow As Integer
Public mstrNoValidasi As String
Public VFilter As String

Public blnStatusFrmUtama As Boolean
Public blnForm As Boolean

Public mstrNoPenKlaim As String
Public mstrNoStrukKlaim As String
Public status As String
Public sisa As Currency
Public curSisaHutangFinal As Currency
Public curSisaHutangFinalTM As Currency
Public curSisaHutangFinalOA As Currency
Public curSisaHutangPenjamin As Currency
Public curSisaHutangPenjaminTM As Currency
Public curSisaHutangPenjaminOA As Currency
Public curTotalSisaHutangPenjamin As Currency
Public curJmlBayarTM As Currency
Public curJmlBayarOA As Currency

Public TglCetak As Date
Public mstrKdKelasx As String
Public mstrKdRuanganx As String
Public rsx As New recordset
Public rs3 As New recordset
Public rs4 As New recordset
Public rs5 As New recordset
Public rs6 As New recordset

Public mdtglclosing As Date

Public substrKdPegawai As String
Public substrNoOrder As String
Public strnonmedis As Boolean
Public strCetakKendaliLaporan As String
Public intCetakKe As Integer
Public NamaCR As String



Public bolStatusFIFO As Boolean
Public bolStatusHarga As Boolean


Public valEnd As Boolean

Public bolTampil As Boolean
Public rsSplakuk As New ADODB.recordset
Public mblnTemp As Boolean
Public bolDel As Boolean
Public bolEnter As Boolean
Public mcurAll_HrsDibyrMulti As Currency


Public mcurAll_TBP_M As Currency
Public mcurAll_TP_M As Currency
Public mcurAll_TRS_M As Currency
Public mcurAll_Pemb_M As Currency
Public mcurAll_HrsDibyr_M As Currency
Public mcurTM_TBP_M As Currency
Public mcurTM_TP_M As Currency
Public mcurTM_TRS_M As Currency
Public mcurTM_Pemb_M As Currency
Public mcurTM_Discount_M As Currency
Public mcurTM_HrsDibyr_M As Currency
Public mcurTM_HrsDibyrNow_M As Currency
Public mcurTM_JmlByr_M As Currency
Public mcurTM_ST_M As Currency
Public mcurOA_TBP_M As Currency
Public mcurOA_TP_M As Currency
Public mcurOA_TRS_M As Currency
Public mcurOA_Pemb_M As Currency
Public mcurOA_Discount_M As Currency
Public mcurOA_HrsDibyr_M As Currency
Public mcurOA_HrsDibyrNow_M As Currency
Public mcurOA_JmlByr_M As Currency
Public mcurOA_ST_M As Currency
Public mcurBayarTM_M As Currency
Public mcurBayarOA_M As Currency
Public mcurPembebasan_M As Currency
Public mStatMulti As Boolean
Public mstrNoBKMAsal As String
Public mcurBayarMulti As Currency
Public bolLapSup As Boolean
Public cMix As Currency
Public cTotMix As Currency
Public bolbatal As Boolean
Public bolRetur As Boolean
Public sNoRetur As String
Public KdInstalasiCetak As String
Public iJnsByr As Integer




Public mstrKdJenisPasienMultiPenjamin As String
Public mstrKdPenjaminPasienMultiPenjamin As String
Public curBiayaAdmin As Currency
Public curBiayaAdminRetur As Currency
Public bolMultiReturPerpelayanan As Boolean

Public sPrinterLabel1 As String
Public sPrinterLabel2 As String

Public strGlobalAdaKenaikanKelas As String
Public strGlobalKelasPerawatan As String
Public strGlobalNamaKelasTertinggiNaikKelas As String
Public strGlobalKdRuanganNaikKelas As String
Public strGlobalKdKelasNaikKelas As String
