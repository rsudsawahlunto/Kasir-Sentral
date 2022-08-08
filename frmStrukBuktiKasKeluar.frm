VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmStrukBuktiKasKeluar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Bukti Kas Keluar"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStrukBuktiKasKeluar.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10005
   Begin VB.CommandButton cmdTutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   8400
      TabIndex        =   37
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame fraBKK 
      Height          =   4215
      Left            =   0
      TabIndex        =   19
      Top             =   2640
      Width           =   9975
      Begin VB.TextBox txtNamaPenyetor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox txtNoKartu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txtNamaBank 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1440
         Width           =   6135
      End
      Begin VB.TextBox txtKet 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   8
         Top             =   2880
         Width           =   7575
      End
      Begin MSComCtl2.DTPicker dtpTglBKK 
         Height          =   330
         Left            =   6360
         TabIndex        =   0
         Top             =   320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy hh:mm"
         Format          =   146407427
         UpDown          =   -1  'True
         CurrentDate     =   38448
      End
      Begin VB.TextBox txtJmlByr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   2160
         MaxLength       =   17
         TabIndex        =   7
         Top             =   2520
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo dcCaraBayar 
         Height          =   330
         Left            =   2160
         TabIndex        =   3
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483624
         ForeColor       =   128
         Text            =   ""
      End
      Begin VB.TextBox txtNoBKK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2160
         TabIndex        =   26
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   6960
         TabIndex        =   9
         Top             =   3600
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dcTransaksi 
         Height          =   330
         Left            =   2160
         TabIndex        =   2
         Top             =   720
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483624
         ForeColor       =   128
         Text            =   ""
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Nama Penerima"
         Height          =   210
         Left            =   240
         TabIndex        =   36
         Top             =   2220
         Width           =   1260
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "No. Rekening"
         Height          =   210
         Left            =   240
         TabIndex        =   35
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Nama Bank"
         Height          =   210
         Left            =   240
         TabIndex        =   34
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Nama Transaksi"
         Height          =   210
         Left            =   240
         TabIndex        =   25
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Jml. Bayar (Rp. )"
         Height          =   210
         Left            =   240
         TabIndex        =   24
         Top             =   2580
         Width           =   1350
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Transaksi"
         Height          =   210
         Left            =   5160
         TabIndex        =   23
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   210
         Left            =   240
         TabIndex        =   22
         Top             =   2940
         Width           =   945
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cara Bayar"
         Height          =   210
         Left            =   240
         TabIndex        =   21
         Top             =   1140
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "No. Bukti Kas Keluar"
         Height          =   210
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1635
      End
   End
   Begin VB.Frame fraRincianTagihan 
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
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   9975
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   6360
         TabIndex        =   33
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   8160
         TabIndex        =   31
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1680
         TabIndex        =   28
         Top             =   480
         Width           =   4575
      End
      Begin VB.TextBox txtNoTerima 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   240
         MaxLength       =   10
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtTotDisc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   4320
         TabIndex        =   16
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtTotPpn 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2280
         TabIndex        =   14
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtTotBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Terima"
         Height          =   210
         Left            =   6360
         TabIndex        =   32
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Jatuh Tempo"
         Height          =   210
         Left            =   8160
         TabIndex        =   30
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nama Supplier"
         Height          =   210
         Left            =   1680
         TabIndex        =   29
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "No. Terima"
         Height          =   210
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblTotalTagihan 
         Alignment       =   1  'Right Justify
         Caption         =   "Rp. 0,00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6480
         TabIndex        =   18
         Top             =   1200
         Width           =   3090
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Total Tagihan"
         Height          =   210
         Left            =   6480
         TabIndex        =   17
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Discount"
         Height          =   210
         Left            =   4320
         TabIndex        =   15
         Top             =   960
         Width           =   1215
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Ppn"
         Height          =   210
         Left            =   2280
         TabIndex        =   13
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya"
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   885
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   38
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
      Picture         =   "frmStrukBuktiKasKeluar.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8160
      Picture         =   "frmStrukBuktiKasKeluar.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmStrukBuktiKasKeluar.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Menu mnuUtilitas 
      Caption         =   "Utilitas"
      Begin VB.Menu mnuCetak 
         Caption         =   "Cetak Kuitansi"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmStrukBuktiKasKeluar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSimpan_Click()
    On Error GoTo errSave
    If funcCekValidasi = False Then Exit Sub
    If sp_BKK(dbcmd) = False Then Exit Sub
    fraBKK.Enabled = False
    txtNoTerima.SetFocus
errSave:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
    If blnStatusFrmUtama = True Then frmDaftarTagihanSupplier.Enabled = True
    blnStatusFrmUtama = False
End Sub

Private Sub dcCaraBayar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcCaraBayar.BoundText = "01" Then
            txtNamaBank.Enabled = False
            txtNoKartu.Enabled = False
            txtNamaPenyetor.Enabled = True
            txtNamaPenyetor.SetFocus
        ElseIf dcCaraBayar.BoundText = "03" Then
            txtNamaBank.Enabled = True
            txtNoKartu.Enabled = True
            txtNamaPenyetor.Enabled = True
            txtNamaBank.SetFocus
        ElseIf dcCaraBayar.BoundText = "04" Then
            txtNamaBank.Enabled = False
            txtNoKartu.Enabled = False
            txtNamaPenyetor.Enabled = True
            txtNamaPenyetor.SetFocus
        ElseIf dcCaraBayar.BoundText = "05" Then
            txtNamaBank.Enabled = False
            txtNoKartu.Enabled = False
            txtNamaPenyetor.Enabled = True
            txtNamaPenyetor.SetFocus
        ElseIf dcCaraBayar.BoundText = "06" Then
            txtNamaBank.Enabled = True
            txtNoKartu.Enabled = True
            txtNamaPenyetor.Enabled = True
            txtNamaBank.SetFocus
        End If
    End If
End Sub

Private Sub dcTransaksi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcCaraBayar.SetFocus
End Sub

Private Sub dtpTglBKK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcTransaksi.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    subLoadDC
    dtpTglBKK.Value = Now
    typPenjaminSisaTagihanApotik.blnStatus = False

    Set rs = Nothing
    strSQL = "SELECT * FROM V_D_DaftarTagihanSupplier WHERE NoTerima = '" & frmDaftarTagihanSupplier.dgDaftarTagihanSupplier.Columns(0).Value & "'"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    txtNoTerima.Text = rs(0).Value
    Text1.Text = rs(1).Value
    Text3.Text = rs(3).Value
    Text2.Text = rs(4).Value
    txtTotBiaya.Text = Format(rs(5).Value, "#,###,###,###,##0")
    txtTotPpn.Text = Format(rs(6).Value, "#,###,###,###,##0")
    txtTotDisc.Text = Format(rs(7).Value, "#,###,###,###,##0")
    lblTotalTagihan.Caption = Format(rs(9).Value, "#,###,###,###,##0")
    If IsNull(rs(11).Value) Then
        txtNoBKK.Text = ""
        dtpTglBKK.Value = Now
        dcTransaksi.Text = ""
        dcCaraBayar.Text = ""
        txtNamaBank.Text = ""
        txtNoKartu.Text = ""
        txtNamaPenyetor.Text = ""
        txtJmlByr.Text = rs(9).Value
        txtKet.Text = ""
        cmdSimpan.Enabled = True
    Else
        txtNoBKK.Text = rs(11).Value
        dtpTglBKK.Value = rs(12).Value
        dcTransaksi.Text = rs(13).Value
        dcCaraBayar.Text = rs(14).Value
        If IsNull(rs(15).Value) Then
            txtNamaBank.Text = ""
        Else
            txtNamaBank.Text = rs(15).Value
        End If
        If IsNull(rs(16).Value) Then
            txtNoKartu.Text = ""
        Else
            txtNoKartu.Text = rs(16).Value
        End If
        If IsNull(rs(17).Value) Then
            txtNamaPenyetor.Text = ""
        Else
            txtNamaPenyetor.Text = rs(17).Value
        End If
        txtJmlByr.Text = rs(18).Value
        If IsNull(rs(19).Value) Then
            txtKet.Text = ""
        Else
            txtKet.Text = rs(19).Value
        End If
        cmdSimpan.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDaftarTagihanSupplier.Enabled = True
End Sub

Private Sub mnucetak_Click()
    strCetak = "BKK"
    frmCetakKwitansi.Show
End Sub

Private Sub txtJmlByr_Change()
    If txtJmlByr.Text = "" Then txtJmlByr.Text = 0
    txtJmlByr = Format(txtJmlByr, "#,###,###,###,##0")
    txtJmlByr.SelStart = Len(txtJmlByr.Text)
End Sub

Private Sub txtNamaBank_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNoKartu.SetFocus
End Sub

Private Sub txtNoKartu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaPenyetor.SetFocus
End Sub

Private Sub txtNamaPenyetor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJmlByr.SetFocus
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

Private Sub txtNoTerima_KeyPress(KeyAscii As Integer)
    msubSetDeleteKeyComma KeyAscii
    If KeyAscii = 13 Then
        If funcLoadDataTerima(txtNoTerima.Text) = False Then Exit Sub
        fraBKK.Enabled = True
        dtpTglBKK.SetFocus
    End If
End Sub

'untuk membersihkan tampilan data
Private Sub subClearData()
    txtTotBiaya.Text = ""
    txtTotPpn.Text = ""
    txtTotDisc.Text = ""
    txtNoTerima.Text = ""
    txtNoBKK.Text = ""
    dtpTglBKK.Value = Now
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
    strSQL = "SELECT KdCaraBayar,CaraBayar FROM CaraBayar WHERE KdCaraBayar <> '02'"
    msubRecFO rs, strSQL
    Set dcCaraBayar.RowSource = rs
    dcCaraBayar.BoundColumn = rs(0).Name
    dcCaraBayar.ListField = rs(1).Name
    Exit Sub
errLoad:
    msubPesanError
End Sub

'untuk loading data struk apotik
Private Function funcLoadDataTerima(strNoTerima As String) As Boolean
    On Error GoTo errLoad
    funcLoadDataTerima = False
    strSQL = "SELECT * FROM TagihanSupplier WHERE NoTerima='" & strFilter & "'"
    msubRecFO rs, strSQL
    If rs.RecordCount <> 0 Then
        txtTotBiaya.Text = rs("TotalTagihan").Value
        txtTotPpn.Text = rs("TotalPpn").Value
        txtTotDisc.Text = rs("TotalDiscount").Value
        lblTotalTagihan.Caption = FormatCurrency(rs("TotalTagihan").Value + rs("TotalPpn").Value - rs("TotalDiscount").Value, 2)
        funcLoadDataTerima = True
    Else
        subClearData
        fraBKK.Enabled = False
        MsgBox "No.Struk tersebut tidak terdaftar", vbCritical, "Validasi"
    End If
    Exit Function
errLoad:
    fraBKK.Enabled = False
    msubPesanError
End Function

'untuk mengecek validasi data yang akan disimpan
Private Function funcCekValidasi() As Boolean
    funcCekValidasi = False
    If txtNoTerima.Text = "" Then
        MsgBox "No.Terima Supplier harus diisi", vbCritical, "Validasi"
        txtNoTerima.SetFocus
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

'Store procedure untuk menyimpan atau mengubah Struk Bukti Kas Keluar
Private Function sp_BKK(ByVal adoCommand As ADODB.Command) As Boolean
    On Error GoTo errSp_BKK
    Dim strLokal As String
    sp_BKK = False
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, txtNoTerima.Text)
        .Parameters.Append .CreateParameter("NoStruk", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("TglBKK", adDate, adParamInput, , Format(dtpTglBKK.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdTransaksi", adVarChar, adParamInput, 5, dcTransaksi.BoundText)
        .Parameters.Append .CreateParameter("KdCaraBayar", adChar, adParamInput, 2, dcCaraBayar.BoundText)
        If txtNamaBank.Text = "" Then
            .Parameters.Append .CreateParameter("NamaBank", adVarChar, adParamInput, 100, Null)
        Else
            .Parameters.Append .CreateParameter("NamaBank", adVarChar, adParamInput, 100, txtNamaBank.Text)
        End If
        If txtNoKartu.Text = "" Then
            .Parameters.Append .CreateParameter("NoAccount", adVarChar, adParamInput, 50, Null)
        Else
            .Parameters.Append .CreateParameter("NoAccount", adVarChar, adParamInput, 50, txtNoKartu.Text)
        End If
        If txtNamaPenyetor.Text = "" Then
            .Parameters.Append .CreateParameter("AtasNama", adVarChar, adParamInput, 50, Null)
        Else
            .Parameters.Append .CreateParameter("AtasNama", adVarChar, adParamInput, 50, txtNamaPenyetor.Text)
        End If
        .Parameters.Append .CreateParameter("JmlBayar", adCurrency, adParamInput, , CCur(txtJmlByr.Text))
        If txtKet.Text = "" Then
            .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, Null)
        Else
            .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, txtKet.Text)
        End If
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuanganKasir)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, UserID)
        .Parameters.Append .CreateParameter("OutputNoBKK", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "Add_StrukBuktiKasKeluar"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data Struk Bukti Kas Keluar", vbCritical, "Validasi"
        Else
            If Not IsNull(.Parameters("OutputNoBKK").Value) Then txtNoBKK.Text = .Parameters("OutputNoBKK").Value
            If Len(txtNoBKK.Text) = 0 Then
                strLokal = "SELECT NoBKK from StrukBuktiKasKeluar where tglBKK = '" & Format(dtpTglBKK.Value, "yyyy/MM/dd HH:mm:ss") & "' and kdRuangan = '" & mstrKdRuanganKasir & "' and idUser = '" & UserID & "'"
                Call msubRecFO(rs, strLokal)
                txtNoBKK.Text = rs("NoBKK").Value
            End If
            MsgBox "Pemasukan data Struk Bukti Kas Keluarsukses", vbInformation, "Validasi"
            sp_BKK = True
            Call Add_HistoryLoginActivity("Add_StrukBuktiKasKeluar")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Function
errSp_BKK:
    Call deleteADOCommandParameters(adoCommand)
    Set adoCommand = Nothing
    msubPesanError
End Function
