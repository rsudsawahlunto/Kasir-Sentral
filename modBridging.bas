Attribute VB_Name = "modBridging"
Option Explicit

Public currGlobalTmpTotalDetailTarif As Currency

Public currGlobalTarifCBG As Currency
Public currGlobalTarifSubAcute As Currency
Public currGlobalTarifChronic As Currency
Public currGlobalBiayaTambahan As Currency
Public currGlobalTarifINACBGKelas1 As Currency
Public currGlobalTarifINACBGKelas2 As Currency
Public currGlobalTarifINACBGKelas3 As Currency
Public strGlobalSpecialCMGOption() As Currency
Public currGlobalTotalSpecialCMG As Currency


Private Type SpecialCMGOption
    Code As String
    Description As String
    Type As String
End Type

Public globalSpecialCMGOption() As SpecialCMGOption


'----Inisialisasi objec context
Public strGlobalUrlINACBG As String
Public strGlobalINACBGKeyEnkripDanDekripEklaim As String
Public strGlobalINACBGKodeTarifRs As String
Public strGlobalINACBGJenisPasienId As String
Public strGlobalINACBGJenisPasienNama
Public strGlobalINACBGNIKPegawai
Public strGlobalPathReferenceBridgingINACBG As String
Public strGlobalINACBGVersiEklaim As String
Public strGlobalINACBGUrlEklaim51 As String



'2+2=5
Public Function SyncronINACBGPerPasien( _
ParamNopendaftaran As String, _
ParamTotalBiaya As Currency, Optional ParamKoefisienTambahanBiayaKeVIP As Double) _
As String

On Error GoTo duaTambahDuaSamaDenganLima

Dim intLamaHariRawatNaikKelas As Integer
Dim intLamaHariRawatIntensif As Integer
Dim curBiayaTambahanPoliEksekutif As Currency
Dim strAdaKenaikanKelas As Integer
Dim strAdaPerawatanIntensif As String
Dim strNamaKelasTertinggiNaikKelas As String
Dim strKelasPerawatan As String
Dim strNamaPoli As String

Dim strTmpBalikanDariWebServiceINACBGS As String

Dim context As BridgingInaCbg.context

Dim strNamaDokter As String
Dim strCaraPulang As String
Dim strDerajatKelasTanggungan As String
Dim intCountHariYgSama As Integer
Dim dateTglMasukNaikKelasSebelumnya As Date
Dim strKdKelasIntensif As String
Dim dateTglMasukIntensifSebelumnya As Date
Dim NoCMTerm As String

Dim strDiagnosa As String
Dim strDiagnosaTindakan  As String
Dim strJenisRawat As String
Dim i As Integer


Dim boolNoSJPKosong As Boolean
                
    If (Dir(strGlobalPathReferenceBridgingINACBG) <> "") Then
        
        currGlobalTarifCBG = 0
        currGlobalTarifSubAcute = 0
        currGlobalTarifChronic = 0
        currGlobalBiayaTambahan = 0
        currGlobalTarifINACBGKelas1 = 0
        currGlobalTarifINACBGKelas2 = 0
        currGlobalTarifINACBGKelas3 = 0
        
        If (strGlobalUrlINACBG <> "") Then
            Set context = New BridgingInaCbg.context
            
            context.SetEndpoint strGlobalUrlINACBG
            context.SetKey strGlobalINACBGKeyEnkripDanDekripEklaim
            
            'Antisipasi jika masih pakai 5.1
            If strGlobalINACBGVersiEklaim = "5.1" Then
                If strGlobalINACBGUrlEklaim51 <> "" Then
                    context.SetEndpoint strGlobalINACBGUrlEklaim51
                End If
            End If
            
            
            'Diagnosa utama
            strSQL = "SELECT PeriksaDiagnosa.KdDiagnosa FROM PeriksaDiagnosa " _
                     & " INNER JOIN SettingGlobal ON PeriksaDiagnosa.KdJenisDiagnosa = SettingGlobal.Value " _
                     & " WHERE (PeriksaDiagnosa.NoPendaftaran = '" & ParamNopendaftaran & "') AND (SettingGlobal.Prefix = 'KdDiagnosaUtama' )"
                     
            Call msubRecFO(rsD, strSQL)
            
            strDiagnosa = ""
            
            For i = 1 To rsD.RecordCount
                If (strDiagnosa = "") Then
                    strDiagnosa = rsD(0).Value
                Else
                    strDiagnosa = rsD(0).Value & ";"
                End If
                rsD.MoveNext
            Next i
            
            'Diagnosa tambahan
            strSQL = "SELECT PeriksaDiagnosa.KdDiagnosa FROM PeriksaDiagnosa " _
                     & " INNER JOIN SettingGlobal ON PeriksaDiagnosa.KdJenisDiagnosa = SettingGlobal.Value " _
                     & " WHERE (PeriksaDiagnosa.NoPendaftaran = '" & ParamNopendaftaran & "') AND (SettingGlobal.Prefix = 'KdDiagnosaTambahan' )"
                     
            Call msubRecFO(rsD, strSQL)
            For i = 1 To rsD.RecordCount
                strDiagnosa = strDiagnosa & ";" & rsD(0).Value
                rsD.MoveNext
            Next i

            'Diangosa tindakan utama
            strSQL = "SELECT DetailPeriksaDiagnosa.KdDiagnosaTindakan FROM DetailPeriksaDiagnosa " _
                     & " INNER JOIN SettingGlobal ON DetailPeriksaDiagnosa.KdJenisDiagnosa = SettingGlobal.Value " _
                     & " WHERE (DetailPeriksaDiagnosa.NoPendaftaran = '" & ParamNopendaftaran & "') AND (SettingGlobal.Prefix = 'KdDiagnosaUtama' )"
                     
            Call msubRecFO(rsD, strSQL)
            
            strDiagnosaTindakan = ""
            For i = 1 To rsD.RecordCount
                strDiagnosaTindakan = strDiagnosaTindakan & rsD(0).Value + ";"
                rsD.MoveNext
            Next i

            'Diagnosa tindakan tambahan
            strSQL = "SELECT DetailPeriksaDiagnosa.KdDiagnosaTindakan FROM DetailPeriksaDiagnosa " _
                     & " INNER JOIN SettingGlobal ON DetailPeriksaDiagnosa.KdJenisDiagnosa = SettingGlobal.Value " _
                     & " WHERE (DetailPeriksaDiagnosa.NoPendaftaran = '" & ParamNopendaftaran & "') AND (SettingGlobal.Prefix = 'KdDiagnosaTambahan' )"
                     
            Call msubRecFO(rsD, strSQL)
            For i = 1 To rsD.RecordCount
                strDiagnosaTindakan = strDiagnosaTindakan & rsD(0).Value + ";"
                rsD.MoveNext
            Next i
            
            strSQL = "Select * From V_BridgingINACBGSNew Where NoPendaftaran='" & ParamNopendaftaran & "'"
            Call msubRecFO(rs, strSQL)
            If rs.RecordCount = 0 Then
                strSQL = "Select * From V_BridgingINACBGS Where NoPendaftaran='" & ParamNopendaftaran & "'"
                Call msubRecFO(rs, strSQL)
            End If
            
            
            If IsNull(rs("NoSJP")) = True Then
                boolNoSJPKosong = True
            End If
            
            If rs.RecordCount < 1 Then
                boolNoSJPKosong = True
            End If
            
            If boolNoSJPKosong = False Then
                If rs("NoSJP") = "" Then
                    boolNoSJPKosong = True
                End If
            End If
            
            If boolNoSJPKosong = True Then
                MsgBox "NoSJP kosong", vbCritical
                SyncronINACBGPerPasien = "NoSJP kosong"
                Exit Function
            End If
            
            'Jenis rawat
            strSQL = "select * From RegistrasiRI where NoPendaftaran='" & ParamNopendaftaran & "' "
            Call msubRecFO(rsE, strSQL)
            If rsE.RecordCount <> 0 Then
                strJenisRawat = "1"
            Else
                strJenisRawat = "2"
            End If
            
           
            'Dokter pemeriksa
            If strJenisRawat = "1" Then
                'Rawat Inap
                strSQL = "Select IdPegawai as IdDokter From DataPegawai Where IdPegawai='" & frmTagihanPasien.dcDokter.BoundText & "'"
                Call msubRecFO(rsE, strSQL)
                
            Else
                'IGD
                strSQL = "Select IdDokter From RegistrasiIGD Where NoPendaftaran='" & ParamNopendaftaran & "'"
                Call msubRecFO(rsE, strSQL)
                
                'Rawat Jalan
                If rsE.RecordCount = 0 Then
                    strSQL = "Select IdDokter From RegistrasiRJ Where NoPendaftaran='" & ParamNopendaftaran & "'"
                    Call msubRecFO(rsE, strSQL)
                End If
            End If
            
           
            If rsE.RecordCount <> 0 Then
                If IsNull(rsE("IdDokter")) = False Then
                    strSQL = "Select NamaLengkap From DataPegawai Where IdPegawai='" & rsE("IdDokter") & "'"
                    Call msubRecFO(rsF, strSQL)
                    strNamaDokter = rsF("NamaLengkap")
            
                Else
                    strNamaDokter = ""
                End If
            Else
                strNamaDokter = ""
            End If
                
            If strNamaDokter = "" Then
                MsgBox "Dokter penanggung jawab kosong", vbCritical
                SyncronINACBGPerPasien = "Dokter penanggung jawab kosong"
                Exit Function
            End If
            
            
            '2+2=5, cara pulang
            'Rawat Inap
            If strJenisRawat = 1 Then
                strSQL = "Select KdKondisiPulang From PasienPulang Where NoPendaftaran='" & ParamNopendaftaran & "'"
                Call msubRecFO(rsE, strSQL)
                If rsE.RecordCount <> 0 Then
                    strSQL = "Select KdCaraPulangINACBGS From MappingCaraPulangINACBGS Where KdKondisiPulangSIMRS='" & rsE("KdKondisiPulang") & "'"
                    Call msubRecFO(rsF, strSQL)
                    If IsNull(rsF("KdCaraPulangINACBGS")) = True Then
                        strCaraPulang = ""
                    Else
                        strCaraPulang = rsF("KdCaraPulangINACBGS")
                    End If
                Else
                    strCaraPulang = ""
                End If
                
            Else
                strSQL = "Select KdKondisiPulang From PasienIGDKeluar Where NoPendaftaran='" & ParamNopendaftaran & "'"
                Call msubRecFO(rsE, strSQL)
                
                'IGD
                If rsE.RecordCount <> 0 Then
                    If IsNull(rsE("KdKondisiPulang")) = True Then
                        strCaraPulang = ""
                    Else
                        strSQL = "Select KdCaraPulangINACBGS From MappingCaraPulangINACBGS Where KdKondisiPulangSIMRS='" & rsE("KdKondisiPulang") & "'"
                        Call msubRecFO(rsF, strSQL)
                        If IsNull(rsF("KdCaraPulangINACBGS")) = True Then
                            strCaraPulang = ""
                        Else
                            strCaraPulang = rsF("KdCaraPulangINACBGS")
                        End If
                    End If
                Else
                    'Rawat Jalan
                    strSQL = "Select KdKondisiPulang, KdRuangan From PasienRJPulang Where NoPendaftaran='" & ParamNopendaftaran & "'"
                    Call msubRecFO(rsE, strSQL)
                    If rsE.RecordCount <> 0 Then
                        If IsNull(rsE("KdKondisiPulang")) = True Then
                            strCaraPulang = ""
                        Else
                            strSQL = "Select KdCaraPulangINACBGS From MappingCaraPulangINACBGS Where KdKondisiPulangSIMRS='" & rsE("KdKondisiPulang") & "'"
                            Call msubRecFO(rsF, strSQL)
                            If IsNull(rsF("KdCaraPulangINACBGS")) = True Then
                                strCaraPulang = ""
                            Else
                                strCaraPulang = rsF("KdCaraPulangINACBGS")
                            End If
                        End If
                        
                        'Khusus poli eksekutif, mencari biaya tambahan untuk poli eksekutif
                        strSQL = "Select NamaExternal From Ruangan Where KdRuangan='" & rsE("KdRuangan") & "'"
                        Call msubRecFO(rsG, strSQL)
                        strNamaPoli = rsG("NamaRuangan")
                        If strNamaPoli = "Poli Eksekutif" Then
                            strSQL = "Select Value From SettingGlobal Where Prefix='BiayaTambahanPoliEksekutif'"
                            Call msubRecFO(rsG, strSQL)
                            If rsG.RecordCount <> 0 Then
                                curBiayaTambahanPoliEksekutif = rsG("Value")
                            Else
                                curBiayaTambahanPoliEksekutif = 0
                            End If
                            
                        Else
                            curBiayaTambahanPoliEksekutif = 0
                        End If
                        
                        
                    Else
                        strCaraPulang = ""
                    End If
                    
                End If
            End If
                        
            
            
            'Kelas perawatan(Kelas Tanggungan)
            strKelasPerawatan = 1
            If strJenisRawat = "1" Then
                If rs("KdKelasDitanggung") = "01" Then
                    strKelasPerawatan = "3"
                ElseIf rs("KdKelasDitanggung") = "02" Then
                    strKelasPerawatan = "2"
                ElseIf rs("KdKelasDitanggung") = "03" Then
                    strKelasPerawatan = "1"
                End If
            Else
                'Saya patok IGD ke kelas reguler karena IGD tidak ada yang eksekutif. IGD masuk ke rawat jalan di E-Klaim INACBGS
                strKelasPerawatan = "3"
                
                'Replace Kelas perawatan jika ke poli eksekutif
                If strNamaPoli = "Poli Eksekutif" Then
                    strKelasPerawatan = "1"
                End If
            End If
            
            
            
            
            'Menentukan derajat kelas tanggungan
            If strJenisRawat = 1 Then
                If strKelasPerawatan = 1 Then
                    strDerajatKelasTanggungan = 3
                ElseIf strKelasPerawatan = 2 Then
                    strDerajatKelasTanggungan = 2
                Else
                    strDerajatKelasTanggungan = 1
                End If
            End If
            
            
            'Pengecekan apakah ada pernah naik kelas atau tidak
            strAdaKenaikanKelas = "0"
            intLamaHariRawatNaikKelas = 0
            intCountHariYgSama = 0
            
            If strJenisRawat = 1 Then
                strSQL = "Select TglMasuk, TglKeluar, DerajatKelas, NamaEksternal, KdRuangan, KdKelas, KdKelasPel From V_DerajatKelasPemakaianKamar " _
                       & " Where NoPendaftaran='" & ParamNopendaftaran & "' And DerajatKelas > '" & strDerajatKelasTanggungan & "' Order By DerajatKelas Asc"
                Call msubRecFO(rsF, strSQL)
                
                If rsF.RecordCount > 0 Then
                    strAdaKenaikanKelas = "1"
                    
                    
                    'Jumlah hari rawat naik kelas
                    For i = 1 To rsF.RecordCount
                        
                        'Untuk kasus naik kelas berkali-kali di hari yang sama
                        If DateDiff("d", dateTglMasukNaikKelasSebelumnya, rsF("TglMasuk")) = 0 Then
                            intCountHariYgSama = intCountHariYgSama + 1
                        End If
                        
                        
                        If DateDiff("d", rsF("TglMasuk"), rsF("TglKeluar")) = 0 Then
                            intLamaHariRawatNaikKelas = intLamaHariRawatNaikKelas + 1
                        Else
                            intLamaHariRawatNaikKelas = intLamaHariRawatNaikKelas + (DateDiff("d", rsF("TglMasuk"), rsF("TglKeluar")) + 1)
                        End If
                          
                        If i = rsF.RecordCount Then
                            strNamaKelasTertinggiNaikKelas = rsF("NamaEksternal")
                            strGlobalKdRuanganNaikKelas = rsF("KdRuangan")
                            strGlobalKdKelasNaikKelas = rsF("KdKelasPel")
                        End If
                        
                        dateTglMasukNaikKelasSebelumnya = Format(rsF("TglMasuk"), "yyyy-MM-dd HH:mm:ss")
                        
                        rsF.MoveNext
                        
                        
                    Next i
                        
                    'Untuk kasus naik kelas berkali-kali di hari yang sama
                    intLamaHariRawatNaikKelas = intLamaHariRawatNaikKelas - intCountHariYgSama
                    intCountHariYgSama = 0
                
                End If
                
                
                'Ada perawatan intensif
                strSQL = "Select Value From SettingGlobal Where Prefix='KdKelasIntensif'"
                Call msubRecFO(rsE, strSQL)
                If rsE.RecordCount <> 0 Then
                    strKdKelasIntensif = rsE("Value")
                Else
                    strKdKelasIntensif = "08"
                End If
                    
                
                strSQL = "Select TglMasuk, TglKeluar From V_DerajatKelasPemakaianKamar Where NoPendaftaran='" & ParamNopendaftaran & "' And KdKelasPel='" & strKdKelasIntensif & "'"
                Call msubRecFO(rsE, strSQL)
                If rsE.RecordCount <> 0 Then
                    
                    strAdaPerawatanIntensif = "1"
                    
                    For i = 1 To rsE.RecordCount
                        
                        'Untuk kasus masuk intensif berkali-kali di hari yang sama
                        If DateDiff("d", dateTglMasukIntensifSebelumnya, rsE("TglMasuk")) = 0 Then
                            intCountHariYgSama = intCountHariYgSama + 1
                        End If
                        
                        If DateDiff("d", rsE("TglMasuk"), rsE("TglKeluar")) = 0 Then
                            intLamaHariRawatIntensif = intLamaHariRawatIntensif + 1
                        Else
                            intLamaHariRawatIntensif = intLamaHariRawatIntensif + (DateDiff("d", rsE("TglMasuk"), rsE("TglKeluar")) + 1)
                        End If
                        rsE.MoveNext
                    Next i
                    
                    'Untuk kasus masuk intensif berkali-kali di hari yang sama
                    intLamaHariRawatNaikKelas = intLamaHariRawatNaikKelas - intCountHariYgSama
                    
                Else
                    strAdaPerawatanIntensif = 0
                End If
                
                
                
            End If

            
            strSQL = "SELECT SUBSTRING(NoCM,5,2)+SUBSTRING(NoCM,3,2)+SUBSTRING(NoCM,1,2) from Pasien where NoCM='" & rs("NoCM") & "'"
            Set rs503 = Nothing
            Call msubRecFO(rs503, strSQL)
            NoCMTerm = rs503(0).Value
                
             If strGlobalINACBGVersiEklaim = "5.2" Then
                strTmpBalikanDariWebServiceINACBGS = context.SyncronPasien(rs("NoKartu"), rs("NoSJP"), NoCMTerm, rs("NamaPasien"), _
                                                                        Format(rs("TglLahir"), "yyyy-MM-dd HH:mm:ss"), IIf(rs("JenisKelamin") = "L", "1", "2"), _
                                                                        strDiagnosa, strDiagnosaTindakan, strCaraPulang, strJenisRawat, _
                                                                        Format(rs("TglMasuk"), "yyyy-MM-dd HH:mm:ss"), _
                                                                        Format(rs("TglPulang"), "yyyy-MM-dd HH:mm:ss"), _
                                                                        strKelasPerawatan, _
                                                                        getDetailTarif(ParamNopendaftaran, "NonBedah"), _
                                                                        getDetailTarif(ParamNopendaftaran, "Bedah"), _
                                                                        getDetailTarif(ParamNopendaftaran, "Konsultasi"), _
                                                                        getDetailTarif(ParamNopendaftaran, "TenagaAhli"), _
                                                                        getDetailTarif(ParamNopendaftaran, "Keperawatan"), _
                                                                        getDetailTarif(ParamNopendaftaran, "Penunjang"), _
                                                                        getDetailTarif(ParamNopendaftaran, "Radiologi"), _
                                                                        getDetailTarif(ParamNopendaftaran, "Laboratorium"), _
                                                                        getDetailTarif(ParamNopendaftaran, "PelayananDarah"), _
                                                                        getDetailTarif(ParamNopendaftaran, "Rehabilitasi"), _
                                                                        getDetailTarif(ParamNopendaftaran, "Kamar"), _
                                                                        getDetailTarif(ParamNopendaftaran, "RawatIntensif"), _
                                                                        getDetailTarif(ParamNopendaftaran, "Obat"), _
                                                                        getDetailTarif(ParamNopendaftaran, "Alkes"), _
                                                                        getDetailTarif(ParamNopendaftaran, "BMHP"), _
                                                                        0, _
                                                                        curBiayaTambahanPoliEksekutif, strNamaDokter, "CP", strGlobalINACBGJenisPasienId, _
                                                                        strGlobalINACBGJenisPasienNama, strGlobalINACBGNIKPegawai, strAdaKenaikanKelas, strNamaKelasTertinggiNaikKelas, _
                                                                        intLamaHariRawatNaikKelas, ParamKoefisienTambahanBiayaKeVIP, strAdaPerawatanIntensif, intLamaHariRawatIntensif, "0")
                
                
                
            Else
                strTmpBalikanDariWebServiceINACBGS = context.SyncronPasien51(rs("NoKartu"), _
                                                                    rs("NoSJP"), _
                                                                    rs("NoCM"), _
                                                                    rs("NamaPasien"), _
                                                                    Format(rs("TglLahir"), "yyyy-MM-dd HH:mm:ss"), _
                                                                    IIf(rs("JenisKelamin") = "L", "1", "2"), _
                                                                    strDiagnosa, _
                                                                    strDiagnosaTindakan, _
                                                                    strCaraPulang, _
                                                                    strJenisRawat, _
                                                                    Format(rs("TglMasuk"), "yyyy-MM-dd HH:mm:ss"), _
                                                                    Format(rs("TglPulang"), "yyyy-MM-dd HH:mm:ss"), _
                                                                    strKelasPerawatan, _
                                                                    Format(ParamTotalBiaya), _
                                                                    curBiayaTambahanPoliEksekutif, _
                                                                    strNamaDokter, _
                                                                    strGlobalINACBGKodeTarifRs, _
                                                                    strGlobalINACBGJenisPasienId, _
                                                                    strGlobalINACBGJenisPasienNama, _
                                                                    strGlobalINACBGNIKPegawai, _
                                                                    strAdaKenaikanKelas, _
                                                                    strNamaKelasTertinggiNaikKelas, _
                                                                    intLamaHariRawatNaikKelas, _
                                                                    ParamKoefisienTambahanBiayaKeVIP, strAdaPerawatanIntensif, intLamaHariRawatIntensif)
            End If
                
                
            
'            MsgBox strTmpBalikanDariWebServiceINACBGS
            SyncronINACBGPerPasien = strTmpBalikanDariWebServiceINACBGS
            
            SyncronINACBGPerPasien = ""
            
            'Global
            strGlobalAdaKenaikanKelas = strAdaKenaikanKelas
            strGlobalKelasPerawatan = strKelasPerawatan
            strGlobalNamaKelasTertinggiNaikKelas = strNamaKelasTertinggiNaikKelas
                
        End If
    End If



Exit Function

duaTambahDuaSamaDenganLima:

Call msubPesanError
'Resume 0

End Function


Public Function HapusClaim(ParamNopendaftaran As String) As String
On Error GoTo duaTambahDuaSamaDenganLima

    Dim tmpBalikanDariWebServiceINACBGS As String
    Dim context As BridgingInaCbg.context
    Dim Hasil As String
    
    
    If (Dir(strGlobalPathReferenceBridgingINACBG) <> "") Then
        
        
        If strGlobalUrlINACBG <> "" Then
            Set context = New BridgingInaCbg.context
            context.SetEndpoint strGlobalUrlINACBG
            context.SetKey strGlobalINACBGKeyEnkripDanDekripEklaim
            
            'Antisipasi jika masih pakai 5.1
            If strGlobalINACBGVersiEklaim = "5.1" Then
                If strGlobalINACBGUrlEklaim51 <> "" Then
                    context.SetEndpoint strGlobalINACBGUrlEklaim51
                End If
            End If
            
            strSQL = "Select NoSJP From V_BridgingINACBGS Where NoPendaftaran='" & ParamNopendaftaran & "'"
            Call msubRecFO(rs, strSQL)
            If rs.RecordCount > 0 Then
                tmpBalikanDariWebServiceINACBGS = context.HapusClaim(rs("NoSJP"), strGlobalINACBGNIKPegawai)
            Else
                MsgBox "NoSJP tidak ada", vbInformation
            End If
            
'            MsgBox "Hapus klaim " & tmpBalikanDariWebServiceINACBGS
            
        End If
    End If

Exit Function

duaTambahDuaSamaDenganLima:
Call msubPesanError
'Resume 0
End Function


Public Function FinalKlaim(ParamNopendaftaran As String) As String
On Error GoTo duaTambahDuaSamaDenganLima

    Dim tmpBalikanDariWebServiceINACBGS As String
    Dim context As BridgingInaCbg.context
    Dim Hasil As String
    
    
    If (Dir(strGlobalPathReferenceBridgingINACBG) <> "") Then
        
        
        If strGlobalUrlINACBG <> "" Then
            Set context = New BridgingInaCbg.context
            context.SetEndpoint strGlobalUrlINACBG
            context.SetKey strGlobalINACBGKeyEnkripDanDekripEklaim
            
            'Antisipasi jika masih pakai 5.1
            If strGlobalINACBGVersiEklaim = "5.1" Then
                If strGlobalINACBGUrlEklaim51 <> "" Then
                    context.SetEndpoint strGlobalINACBGUrlEklaim51
                End If
            End If
            
            strSQL = "Select NoSJP From V_BridgingINACBGS Where NoPendaftaran='" & ParamNopendaftaran & "'"
            Call msubRecFO(rs, strSQL)
            
            If rs.RecordCount > 0 Then
                tmpBalikanDariWebServiceINACBGS = context.FinalClaim(rs("NoSJP"), strGlobalINACBGNIKPegawai)
            Else
                MsgBox "NoSJP tidak ada", vbInformation
            End If
            
'            MsgBox "Hapus klaim " & tmpBalikanDariWebServiceINACBGS
            
        End If
    End If

Exit Function

duaTambahDuaSamaDenganLima:
Call msubPesanError
End Function

Public Function EditFinalKlaim(ParamNopendaftaran As String) As String
On Error GoTo duaTambahDuaSamaDenganLima

    Dim tmpBalikanDariWebServiceINACBGS As String
    Dim context As BridgingInaCbg.context
    Dim Hasil As String
    
    
    If (Dir(strGlobalPathReferenceBridgingINACBG) <> "") Then
        
        
        If strGlobalUrlINACBG <> "" Then
            Set context = New BridgingInaCbg.context
            context.SetEndpoint strGlobalUrlINACBG
            context.SetKey strGlobalINACBGKeyEnkripDanDekripEklaim
            
            'Antisipasi jika masih pakai 5.1
            If strGlobalINACBGVersiEklaim = "5.1" Then
                If strGlobalINACBGUrlEklaim51 <> "" Then
                    context.SetEndpoint strGlobalINACBGUrlEklaim51
                End If
            End If
            
            strSQL = "Select NoSJP From V_BridgingINACBGS Where NoPendaftaran='" & ParamNopendaftaran & "'"
            Call msubRecFO(rs, strSQL)
            If rs.RecordCount > 0 Then
                tmpBalikanDariWebServiceINACBGS = context.EditUlangFinalClaim(rs("NoSJP"))
            Else
                MsgBox "NoSJP tidak ada", vbInformation
            End If
            
'            MsgBox "Hapus klaim " & tmpBalikanDariWebServiceINACBGS
            
        End If
    End If

Exit Function

duaTambahDuaSamaDenganLima:
Call msubPesanError
End Function

Public Function CetakKlaim(ParamNopendaftaran As String) As String
On Error GoTo duaTambahDuaSamaDenganLima

    Dim tmpBalikanDariWebServiceINACBGS As String
    Dim context As BridgingInaCbg.context
    
    If (Dir(strGlobalPathReferenceBridgingINACBG) <> "") Then
        
        
        If strGlobalUrlINACBG <> "" Then
            Set context = New BridgingInaCbg.context
            context.SetEndpoint strGlobalUrlINACBG
            context.SetKey strGlobalINACBGKeyEnkripDanDekripEklaim
            
            'Antisipasi jika masih pakai 5.1
            If strGlobalINACBGVersiEklaim = "5.1" Then
                If strGlobalINACBGUrlEklaim51 <> "" Then
                    context.SetEndpoint strGlobalINACBGUrlEklaim51
                End If
            End If
            
            strSQL = "Select NoSJP From V_BridgingINACBGS Where NoPendaftaran='" & ParamNopendaftaran & "'"
            Call msubRecFO(rs, strSQL)
            
            If rs.RecordCount > 0 Then
                
                
                If cariSettingGlobal("PathCetakanKlaim") <> "" Then
                    tmpBalikanDariWebServiceINACBGS = context.CetakClaim(rs("NoSJP"), cariSettingGlobal("PathCetakanKlaim"))
                Else
                    tmpBalikanDariWebServiceINACBGS = context.CetakClaim(rs("NoSJP"), "D:\\Eklaim\\")
                End If
                
                
            Else
                MsgBox "NoSJP tidak ada", vbInformation
            End If
            
        End If
    End If

Exit Function

duaTambahDuaSamaDenganLima:
Call msubPesanError

End Function


Public Function SimulasiTarifGrouperStage2(ParamNopendaftaran As String, ParamSpecialCMG As String) As String()
On Error GoTo duaTambahDuaSamaDenganLima

    Dim tmpBalikanDariWebServiceINACBGS() As String
    Dim context As BridgingInaCbg.context
    
    
    If (Dir(strGlobalPathReferenceBridgingINACBG) <> "") Then
        
        
        If strGlobalUrlINACBG <> "" Then
            Set context = New BridgingInaCbg.context
            
            context.SetEndpoint strGlobalUrlINACBG
            context.SetKey strGlobalINACBGKeyEnkripDanDekripEklaim
            
            'Antisipasi jika masih pakai 5.1
            If strGlobalINACBGVersiEklaim = "5.1" Then
                If strGlobalINACBGUrlEklaim51 <> "" Then
                    context.SetEndpoint strGlobalINACBGUrlEklaim51
                End If
            End If
            
            strSQL = "Select NoSJP From V_BridgingINACBGS Where NoPendaftaran='" & ParamNopendaftaran & "'"
            Call msubRecFO(rs, strSQL)
            
            If rs.RecordCount > 0 Then
                tmpBalikanDariWebServiceINACBGS = context.SimulasiTarifGrouperStage2(rs("NoSJP"), ParamSpecialCMG)
                Call InsertValueDariArryaHasilGrouperKeVariabel(tmpBalikanDariWebServiceINACBGS)
            Else
                MsgBox "NoSJP tidak ada", vbInformation
            End If
            
        End If
    End If

Exit Function

duaTambahDuaSamaDenganLima:
Call msubPesanError
End Function


Private Function getDetailTarif(ParamNopendaftaran As String, ParamNamaDetailTarif As String) As Currency
    
    If ParamNamaDetailTarif = "Obat" Then
        strsqlE = "Select Sum(TotalBiaya) As 'DetailTarif' From DetailPemakaianAlkes Where NoPendaftaran='" & ParamNopendaftaran & "' "
        Call msubRecFO(rsE, strsqlE)
    Else
        strsqlE = "Select Sum(Tarif) As 'DetailTarif' From V_DetailTarif Where NoPendaftaran='" & ParamNopendaftaran & "' And NamaDetailTarif='" & ParamNamaDetailTarif & "'"
        Call msubRecFO(rsE, strsqlE)
    End If
    
    
    If rsE.RecordCount > 0 Then
        If IsNull(rsE("DetailTarif")) = False Then
            getDetailTarif = rsE("DetailTarif")
            currGlobalTmpTotalDetailTarif = currGlobalTmpTotalDetailTarif + rsE("DetailTarif")
        End If
    End If
            
    
    Exit Function
    
duaDuaLima:
    msubPesanError
    
'Resume 0
End Function

Public Function cariSettingGlobal(ParamPrefix As String) As String
On Error GoTo duaDuaLima
    strsqlx6 = "Select Value from SettingGlobal Where Prefix='" & ParamPrefix & "'"
    Call msubRecFO(rsM, strsqlx6)
    
    If rsM.RecordCount > 0 Then
        cariSettingGlobal = rsM("Value")
    Else
        cariSettingGlobal = ""
    End If

Exit Function
duaDuaLima:
    cariSettingGlobal = ""
    Call msubPesanError
End Function

Public Sub InsertValueDariArryaHasilGrouperKeVariabel(ParamArrHasilSimulasiTarifGrouper() As String)

On Error GoTo duaTambahDuaSamaDenganLima

Dim strTmpSpecialCMGOptionCode As String
Dim strTmpSpecialCMGOptionDescription As String
Dim intCountingSpecialCMGOption As Integer

intCountingSpecialCMGOption = 0

currGlobalTarifCBG = 0
currGlobalTarifSubAcute = 0
currGlobalTarifChronic = 0
currGlobalBiayaTambahan = 0
currGlobalTarifINACBGKelas1 = 0
currGlobalTarifINACBGKelas2 = 0
currGlobalTarifINACBGKelas3 = 0

currGlobalTotalSpecialCMG = 0

Dim i As Integer
Dim j As Integer

j = 1
        For i = LBound(ParamArrHasilSimulasiTarifGrouper) To UBound(ParamArrHasilSimulasiTarifGrouper)
            Dim arr() As String
            arr = Split(ParamArrHasilSimulasiTarifGrouper(i), ":")
            Select Case arr(0)
                    
                    Case "response->cbg->tariff"
                        currGlobalTarifCBG = CCur(arr(1))
                        
                    Case "response->sub_acute->tariff"
                        currGlobalTarifSubAcute = CCur(arr(1))
                         
                    Case "response->chronic->tariff"
                        currGlobalTarifChronic = CCur(arr(1))
                        
                    Case "response->add_payment_amt"
                        currGlobalBiayaTambahan = CCur(arr(1))
                    
                    
                    Case "special_cmg_option->code"
                        strTmpSpecialCMGOptionCode = arr(1)
                        
                    Case "special_cmg_option->description"
                        strTmpSpecialCMGOptionDescription = arr(1)
                        
                    Case "special_cmg_option->type"
                        If strTmpSpecialCMGOptionCode <> "" And strTmpSpecialCMGOptionDescription <> "" Then
                            
                            ReDim Preserve globalSpecialCMGOption(intCountingSpecialCMGOption + 1)
                            globalSpecialCMGOption(intCountingSpecialCMGOption).Code = strTmpSpecialCMGOptionCode
                            globalSpecialCMGOption(intCountingSpecialCMGOption).Description = strTmpSpecialCMGOptionDescription
                            globalSpecialCMGOption(intCountingSpecialCMGOption).Type = arr(1)
                            
                            intCountingSpecialCMGOption = intCountingSpecialCMGOption + 1
                            strTmpSpecialCMGOptionCode = ""
                            strTmpSpecialCMGOptionDescription = ""
                        End If
                    
                    
                    Case "response->special_cmg->tariff"
                        currGlobalTotalSpecialCMG = currGlobalTotalSpecialCMG + CCur(arr(1))
                        
                    Case "tarif_alt->tarif_inacbg"
                        If j = 1 Then
                            If currGlobalTarifINACBGKelas1 = 0 Then
                                currGlobalTarifINACBGKelas1 = CCur(arr(1))
                            End If
                        End If
                        
                        If j = 2 Then
                            If currGlobalTarifINACBGKelas1 <> 0 Then
                                currGlobalTarifINACBGKelas2 = CCur(arr(1))
                            End If
                        End If
                        
                        If j = 3 Then
                            If currGlobalTarifINACBGKelas2 <> 0 Then
                                currGlobalTarifINACBGKelas3 = CCur(arr(1))
                            End If
                        End If
                        j = j + 1
            End Select
        Next i

Exit Sub

duaTambahDuaSamaDenganLima:
msubPesanError

'Resume 0
End Sub

Public Sub inisialisasiObjekContext()

    strGlobalPathReferenceBridgingINACBG = cariSettingGlobal("INACBGPathReferenceBridging")
    strGlobalINACBGVersiEklaim = cariSettingGlobal("INACBGVersiEklaim")
    strGlobalINACBGUrlEklaim51 = cariSettingGlobal("UrlEklaim51")
    strGlobalUrlINACBG = cariSettingGlobal("UrlInaCbg")
    strGlobalINACBGKeyEnkripDanDekripEklaim = cariSettingGlobal("INACBGKeyEnkripDanDekripEklaim")
    strGlobalINACBGKodeTarifRs = cariSettingGlobal("INACBGKodeTarifRs")
    strGlobalINACBGJenisPasienId = cariSettingGlobal("INACBGJenisPasienId")
    strGlobalINACBGJenisPasienNama = cariSettingGlobal("INACBGJenisPasienNama")
    strGlobalINACBGNIKPegawai = cariSettingGlobal("INACBGNIKPegawai")
    
End Sub



Public Function cekStatusVerifikasiTagihan(ParamNopendaftaran As String) As Boolean

On Error GoTo duaTambahDuaSamaDenganLima

    strSQL = "Select NoPendaftaran From VerifikasiTagihan Where NoPendaftaran='" & ParamNopendaftaran & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount > 0 Then
        cekStatusVerifikasiTagihan = True
    Else
        cekStatusVerifikasiTagihan = False
    End If
    
Exit Function

duaTambahDuaSamaDenganLima:
Call msubPesanError
End Function

Public Function cariTotalKlaimBPJSDariDB(ParamNopendaftaran) As Currency


On Error GoTo duaTambahDuaSamaDenganLima

    strSQL = "Select TotalKlaim From TotalBiayaKlaimBPJS Where NoPendaftaran='" & ParamNopendaftaran & "'"
    Call msubRecFO(rs, strSQL)
    
    
    If rs.RecordCount > 0 Then
        cariTotalKlaimBPJSDariDB = rs("TotalKlaim")
    Else
        cariTotalKlaimBPJSDariDB = 0
    End If
    
Exit Function

duaTambahDuaSamaDenganLima:
Call msubPesanError
End Function







