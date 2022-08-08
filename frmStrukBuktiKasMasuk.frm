VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmStrukBuktiKasMasuk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Bukti Kas Masuk"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStrukBuktiKasMasuk.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   8790
   Begin VB.CommandButton cmdTutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   7200
      TabIndex        =   53
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Frame fraRincianTagihan 
      Caption         =   "Rincian Tagihan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   8775
      Begin VB.TextBox txtTotPemb 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   5400
         TabIndex        =   24
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtTotTanggRS 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   3480
         TabIndex        =   22
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtTotBbnPjmn 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         TabIndex        =   20
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtTotDisc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   5400
         TabIndex        =   18
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtTotPpn 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   3480
         TabIndex        =   16
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtTotBiaya 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total Pembebasan"
         Height          =   210
         Left            =   5400
         TabIndex        =   23
         Top             =   960
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "Total Tanggungan RS"
         Height          =   300
         Left            =   3480
         TabIndex        =   21
         Top             =   960
         Width           =   1815
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "Total Beban Penjamin"
         Height          =   240
         Left            =   1560
         TabIndex        =   19
         Top             =   960
         Width           =   1815
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Discount"
         Height          =   210
         Left            =   5400
         TabIndex        =   17
         Top             =   240
         Width           =   1365
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Total Ppn."
         Height          =   210
         Left            =   3480
         TabIndex        =   15
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya"
         Height          =   210
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame fraBKM 
      Enabled         =   0   'False
      Height          =   4095
      Left            =   0
      TabIndex        =   25
      Top             =   3120
      Width           =   8775
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         Height          =   375
         Left            =   4320
         TabIndex        =   11
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox txtNamaBank 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   2280
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1440
         Width           =   6255
      End
      Begin VB.TextBox txtNoKartu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txtNamaPenyetor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox txtKet 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   2280
         MaxLength       =   100
         TabIndex        =   9
         Top             =   2880
         Width           =   6255
      End
      Begin VB.TextBox txtUangKembali 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   6600
         TabIndex        =   8
         Top             =   2520
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpTglBKM 
         Height          =   330
         Left            =   6600
         TabIndex        =   0
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy hh:mm"
         Format          =   127139843
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin VB.TextBox txtJmlByr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   2280
         MaxLength       =   8
         TabIndex        =   7
         Top             =   2520
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo dcCaraBayar 
         Height          =   330
         Left            =   2280
         TabIndex        =   2
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483624
         ForeColor       =   128
         Text            =   ""
      End
      Begin VB.TextBox txtNoBKM 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2280
         TabIndex        =   33
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   5760
         TabIndex        =   10
         Top             =   3600
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dcTransaksi 
         Height          =   330
         Left            =   2280
         TabIndex        =   1
         Top             =   720
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483624
         ForeColor       =   128
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcJenisKartu 
         Height          =   330
         Left            =   6480
         TabIndex        =   3
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483624
         ForeColor       =   128
         Text            =   ""
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kartu"
         Height          =   210
         Left            =   5520
         TabIndex        =   37
         Top             =   1140
         Width           =   885
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Nama Bank"
         Height          =   210
         Left            =   240
         TabIndex        =   36
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "No. Kartu / Rekening"
         Height          =   210
         Left            =   240
         TabIndex        =   35
         Top             =   1860
         Width           =   1725
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Nama Penyetor"
         Height          =   210
         Left            =   240
         TabIndex        =   34
         Top             =   2220
         Width           =   1260
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Uang Kembalian"
         Height          =   210
         Left            =   5160
         TabIndex        =   32
         Top             =   2580
         Width           =   1290
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Nama Transaksi"
         Height          =   210
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Jml. Bayar (Rp. )"
         Height          =   210
         Left            =   240
         TabIndex        =   30
         Top             =   2580
         Width           =   1350
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Transaksi"
         Height          =   210
         Left            =   5280
         TabIndex        =   29
         Top             =   420
         Width           =   1110
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   210
         Left            =   240
         TabIndex        =   28
         Top             =   2940
         Width           =   945
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cara Bayar"
         Height          =   210
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "No. Bukti Kas Masuk"
         Height          =   210
         Left            =   240
         TabIndex        =   26
         Top             =   405
         Width           =   1635
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Total Tagihan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   47
      Top             =   2040
      Width           =   8775
      Begin VB.CheckBox chkDetail 
         Caption         =   "Detail"
         Height          =   255
         Left            =   7680
         TabIndex        =   52
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtSisaTagihan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   5400
         TabIndex        =   50
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sisa Tagihan"
         Height          =   210
         Left            =   5400
         TabIndex        =   51
         Top             =   360
         Width           =   1110
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Total Tagihan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   49
         Top             =   240
         Width           =   1710
      End
      Begin VB.Label lblTotalTagihan 
         Caption         =   "Rp. 100.000.000,00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   48
         Top             =   600
         Width           =   2850
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Tagihan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   38
      Top             =   1080
      Width           =   8775
      Begin VB.TextBox txtRuangan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   6240
         TabIndex        =   45
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtKelPasien 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   4560
         TabIndex        =   44
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtNoStruk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   40
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtTglStruk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2760
         TabIndex        =   39
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan"
         Height          =   210
         Left            =   6240
         TabIndex        =   46
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Kelompok Pasien"
         Height          =   210
         Left            =   4560
         TabIndex        =   43
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "No. Struk"
         Height          =   210
         Left            =   1320
         TabIndex        =   42
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Struk"
         Height          =   210
         Left            =   2760
         TabIndex        =   41
         Top             =   240
         Width           =   1140
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmStrukBuktiKasMasuk.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6960
      Picture         =   "frmStrukBuktiKasMasuk.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmStrukBuktiKasMasuk.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Menu mnuutilitas 
      Caption         =   "Utilitas"
      Begin VB.Menu mnucetak 
         Caption         =   "Cetak Kuitansi"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmStrukBuktiKasMasuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSimpan_Click()
    On Error GoTo errSave
    If funcCekValidasi = False Then Exit Sub
    If CCur(txtSisaTagihan.Text) <> 0 And typPenjaminSisaTagihanApotik.blnStatus = False Then
        frmStrukBuktiKasMasuk.Enabled = False
        With frmPenjaminSisaTagihanApotik
            .Show
            strSQL = "SELECT * FROM v_S_PasienApotik WHERE NoStruk='" & txtNoStruk.Text & "'"
            msubRecFO rs, strSQL
            .txtNoPendaftaran.Text = rs("NoPendaftaran").Value & ""
            .txtNoCM.Text = rs("NoCM").Value & ""
            .txtNamaPasien.Text = rs("NamaLengkap").Value & ""
            .txtSex.Text = rs("JenisKelamin").Value & ""
            .txtThn.Text = rs("UmurTahun").Value & ""
            .txtBln.Text = rs("UmurBulan").Value & ""
            .txtHari.Text = rs("UmurHari").Value & ""
        End With
        Exit Sub
    End If
    If CCur(txtSisaTagihan.Text) = 0 And typPenjaminSisaTagihanApotik.blnStatus = True Then
        typPenjaminSisaTagihanApotik.blnStatus = False
    End If
    If sp_BKM(dbcmd) = False Then Exit Sub
    If typPenjaminSisaTagihanApotik.blnStatus = True Then
        If sp_PenjaminST(dbcmd) = False Then Exit Sub
        typPenjaminSisaTagihanApotik.blnStatus = False
    End If
    fraBKM.Enabled = False
    txtNoStruk.SetFocus
errSave:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
    If blnStatusFrmUtama = True Then frmDaftarTagihanPasienApotik.Enabled = True
    blnStatusFrmUtama = False
End Sub

Private Sub dcCaraBayar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJmlByr.SetFocus
End Sub

Private Sub dcTransaksi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcCaraBayar.SetFocus
End Sub

Private Sub dtpTglBKM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcTransaksi.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    subLoadDC
    dtpTglBKM.Value = Now
    typPenjaminSisaTagihanApotik.blnStatus = False
End Sub

Private Sub txtJmlByr_Change()
    If txtJmlByr.Text = "" Then txtJmlByr.Text = 0
    txtJmlByr = Format(txtJmlByr, "#,###,###,###,##0")
    txtJmlByr.SelStart = Len(txtJmlByr.Text)
    If CCur(lblTotalTagihan.Caption) - CCur(txtJmlByr.Text) >= 0 Then
        txtSisaTagihan.Text = FormatCurrency(CCur(lblTotalTagihan.Caption) - CCur(txtJmlByr.Text), 2)
        txtUangKembali.Text = FormatCurrency(0, 2)
    Else
        txtSisaTagihan.Text = FormatCurrency(0, 2)
        txtUangKembali.Text = FormatCurrency(CCur(txtJmlByr.Text) - CCur(lblTotalTagihan.Caption), 2)
    End If
End Sub

Private Sub txtJmlByr_GotFocus()
    txtJmlByr.SelStart = 0
    txtJmlByr.SelLength = Len(txtJmlByr.Text)
End Sub

Private Sub txtJmlByr_KeyPress(KeyAscii As Integer)
    SetKeyPressToNumber KeyAscii
    If KeyAscii = 13 Then txtKet.SetFocus
End Sub

Private Sub txtKet_KeyPress(KeyAscii As Integer)
    msubSetDeleteKeyComma KeyAscii
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNoStruk_KeyPress(KeyAscii As Integer)
    msubSetDeleteKeyComma KeyAscii
    If KeyAscii = 13 Then
        If funcLoadDataStruk(txtNoStruk.Text) = False Then Exit Sub
        fraBKM.Enabled = True
        dtpTglBKM.SetFocus
    End If
End Sub

'untuk membersihkan tampilan data
Private Sub subClearData()
    txttotbiaya.Text = ""
    txtTotPpn.Text = ""
    txtTotDisc.Text = ""
    txtTotBbnPjmn.Text = ""
    txtTotTanggRS.Text = ""
    txtTotPemb.Text = ""
    txtSisaTagihan.Text = ""
    txtNoStruk.Text = ""
    txtTglStruk.Text = ""
    txtNoBKM.Text = ""
    dtpTglBKM.Value = Now
    dcTransaksi.Text = ""
    dcCaraBayar.Text = ""
    txtJmlByr.Text = ""
    txtUangKembali.Text = ""
    txtKet.Text = ""
End Sub

'untuk loading data combo
Private Sub subLoadDC()
    On Error GoTo errLoad
    strSQL = "SELECT KdTransaksi,NamaTransaksi FROM DaftarTransaksi Order By NamaTransaksi"
    msubRecFO rs, strSQL
    Set dcTransaksi.RowSource = rs
    dcTransaksi.BoundColumn = rs(0).Name
    dcTransaksi.ListField = rs(1).Name

    strSQL = "SELECT KdCaraBayar,CaraBayar FROM CaraBayar"
    msubRecFO rs, strSQL
    Set dcCaraBayar.RowSource = rs
    dcCaraBayar.BoundColumn = rs(0).Name
    dcCaraBayar.ListField = rs(1).Name
    Exit Sub
errLoad:
    msubPesanError
End Sub

'untuk loading data struk apotik
Private Function funcLoadDataStruk(strNoStruk As String) As Boolean
    On Error GoTo errLoad
    funcLoadDataStruk = False
    strSQL = "SELECT * FROM v_S_StrukApotik WHERE NoStruk='" & strNoStruk & "'"
    msubRecFO rs, strSQL
    If rs.RecordCount <> 0 Then
        txttotbiaya.Text = rs("TotalBiaya").Value
        txtTotPpn.Text = rs("TotalPpn").Value
        txtTotDisc.Text = rs("TotalDiscount").Value
        txtTotBbnPjmn.Text = rs("JmlHutangPenjamin").Value
        txtTotTanggRS.Text = rs("JmlTanggunganRS").Value
        txtTotPemb.Text = rs("JmlPembebasan").Value
        txtSisaTagihan.Text = rs("SisaTagihan").Value
        txtTglStruk.Text = rs("TglStruk").Value
        lblTotalTagihan.Caption = FormatCurrency(rs("TotalBiaya").Value + rs("TotalPpn").Value - rs("TotalDiscount").Value - rs("JmlHutangPenjamin").Value - rs("JmlTanggunganRS").Value - rs("JmlPembebasan").Value, 2)
        funcLoadDataStruk = True
    Else
        subClearData
        fraBKM.Enabled = False
        MsgBox "No.Struk tersebut tidak terdaftar", vbCritical, "Validasi"
    End If
    Exit Function
errLoad:
    fraBKM.Enabled = False
    msubPesanError
End Function

'untuk mengecek validasi data yang akan disimpan
Private Function funcCekValidasi() As Boolean
    funcCekValidasi = False
    If txtNoStruk.Text = "" Then
        MsgBox "No.Struk Pasien harus diisi", vbCritical, "Validasi"
        txtNoStruk.SetFocus
        Exit Function
    ElseIf dcTransaksi.Text = "" Then
        MsgBox "Pilihan nama transaksi harus diisi", vbCritical, "Validasi"
        dcTransaksi.SetFocus
        Exit Function
    ElseIf dcCaraBayar.Text = "" Then
        MsgBox "Pilihan cara bayar harus diisi", vbCritical, "Validasi"
        dcCaraBayar.SetFocus
        Exit Function
    ElseIf txtJmlByr.Text = "" Then
        MsgBox "Jumlah bayar harus diisi", vbCritical, "Validasi"
        txtJmlByr.SetFocus
        Exit Function
    End If
    funcCekValidasi = True
End Function

'Store procedure untuk menyimpan atau mengubah Struk Bukti Kas Masuk
Private Function sp_BKM(ByVal adoCommand As ADODB.Command) As Boolean
    On Error GoTo errSp_BKM
    Dim strLokal As String
    sp_BKM = False
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoStruk", adChar, adParamInput, 10, txtNoStruk.Text)
        .Parameters.Append .CreateParameter("TglBKM", adDate, adParamInput, , Format(dtpTglBKM.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdTransaksi", adVarChar, adParamInput, 5, dcTransaksi.BoundText)
        .Parameters.Append .CreateParameter("KdCaraBayar", adChar, adParamInput, 2, dcCaraBayar.BoundText)
        .Parameters.Append .CreateParameter("JmlBayar", adCurrency, adParamInput, , CCur(txtJmlByr.Text))
        .Parameters.Append .CreateParameter("SisaTagihan", adCurrency, adParamInput, , CCur(txtSisaTagihan.Text))
        If txtKet.Text = "" Then
            .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, Null)
        Else
            .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, txtKet.Text)
        End If
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuanganKasir)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, UserID)
        .Parameters.Append .CreateParameter("OutputNoBKM", adChar, adParamOutput, 10, Null)
        .Parameters.Append .CreateParameter("AtasNamaPembayar", adVarChar, adParamInput, 50, Null)

        .ActiveConnection = dbConn
        .CommandText = "Add_StrukBuktiKasMasuk"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data Struk Bukti Kas Masuk", vbCritical, "Validasi"
        Else
            If Not IsNull(.Parameters("OutputNoBKM").Value) Then txtNoBKM.Text = .Parameters("OutputNoBKM").Value
            If Len(txtNoBKM.Text) = 0 Then
                strLokal = "SELECT NoBKM from StrukBuktiKasMasuk where tglBKM = '" & Format(dtpTglBKM.Value, "yyyy/MM/dd HH:mm:ss") & "' and kdRuangan = '" & mstrKdRuanganKasir & "' and idUser = '" & UserID & "'"
                Call msubRecFO(rs, strLokal)
                txtNoBKM.Text = rs("NoBKM").Value
            End If
            MsgBox "Pemasukan data Struk Bukti Kas Masuk sukses", vbInformation, "Validasi"
            sp_BKM = True
            Call Add_HistoryLoginActivity("Add_StrukBuktiKasMasuk")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Function
errSp_BKM:
    Call deleteADOCommandParameters(adoCommand)
    Set adoCommand = Nothing
    msubPesanError
End Function

'Store procedure untuk menyimpan penjamin sisa tagihan pasien
Private Function sp_PenjaminST(ByVal adoCommand As ADODB.Command) As Boolean
    On Error GoTo errSp_PenjaminST
    sp_PenjaminST = False
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoBKM", adChar, adParamInput, 10, txtNoBKM.Text)
        .Parameters.Append .CreateParameter("NamaLengkap", adVarChar, adParamInput, 50, typPenjaminSisaTagihanApotik.strNamaLengkap)
        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(typPenjaminSisaTagihanApotik.dTglLahir, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 1, typPenjaminSisaTagihanApotik.strJenisKelamin)
        .Parameters.Append .CreateParameter("NoIdentitas", adVarChar, adParamInput, 50, typPenjaminSisaTagihanApotik.strNoIdentitas)
        .Parameters.Append .CreateParameter("Hubungan", adVarChar, adParamInput, 50, typPenjaminSisaTagihanApotik.strHubungan)
        .Parameters.Append .CreateParameter("Alamat", adVarChar, adParamInput, 100, typPenjaminSisaTagihanApotik.strAlamat)
        If typPenjaminSisaTagihan.strTelepon = "" Then
            .Parameters.Append .CreateParameter("Telepon", adVarChar, adParamInput, 15, Null)
        Else
            .Parameters.Append .CreateParameter("Telepon", adVarChar, adParamInput, 15, typPenjaminSisaTagihanApotik.strTelepon)
        End If
        .Parameters.Append .CreateParameter("Propinsi", adVarChar, adParamInput, 25, typPenjaminSisaTagihanApotik.strPropinsi)
        .Parameters.Append .CreateParameter("Kota", adVarChar, adParamInput, 25, typPenjaminSisaTagihanApotik.strKota)
        .Parameters.Append .CreateParameter("Kecamatan", adVarChar, adParamInput, 25, typPenjaminSisaTagihanApotik.strKecamatan)
        .Parameters.Append .CreateParameter("Kelurahan", adVarChar, adParamInput, 25, typPenjaminSisaTagihanApotik.strKelurahan)
        If typPenjaminSisaTagihan.strRTRW = "" Then
            .Parameters.Append .CreateParameter("RTRW", adVarChar, adParamInput, 7, Null)
        Else
            .Parameters.Append .CreateParameter("RTRW", adVarChar, adParamInput, 7, typPenjaminSisaTagihanApotik.strRTRW)
        End If
        If typPenjaminSisaTagihan.strKodePos = "" Then
            .Parameters.Append .CreateParameter("KodePos", adChar, adParamInput, 5, Null)
        Else
            .Parameters.Append .CreateParameter("KodePos", adChar, adParamInput, 7, typPenjaminSisaTagihanApotik.strKodePos)
        End If
        .Parameters.Append .CreateParameter("TglBKM", adDate, adParamInput, , Format(dtpTglBKM.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("JmlBayar", adInteger, adParamInput, , CCur(txtJmlByr.Text))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, UserID)
        .Parameters.Append .CreateParameter("SisaTagihan", adInteger, adParamInput, , CCur(txtSisaTagihan.Text))

        .ActiveConnection = dbConn
        .CommandText = "Add_PenjaminSisaTagihanPasienApotik"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan penyimpanan data penjamin Pasien", vbCritical, "Validasi"
        Else
            MsgBox "Penyimpanan data penjamin Pasien sukses", vbInformation, "Validasi"
            sp_PenjaminST = True
            Call Add_HistoryLoginActivity("Add_PenjaminSisaTagihanPasienApotik")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Function
errSp_PenjaminST:
    Call deleteADOCommandParameters(adoCommand)
    Set adoCommand = Nothing
    msubPesanError
End Function
