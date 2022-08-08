VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmStrukBuktiKasKeluarU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Struk Bukti Kas Keluar"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStrukBuktiKasKeluarU.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7725
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   6120
      TabIndex        =   21
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Frame fraBKK 
      Height          =   4215
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   7695
      Begin VB.CheckBox chkStatus 
         Caption         =   "Karyawan RS"
         Height          =   255
         Left            =   5280
         TabIndex        =   22
         Top             =   2198
         Width           =   1455
      End
      Begin VB.TextBox txtNamaPenyetor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   1800
         Width           =   3855
      End
      Begin VB.TextBox txtNamaBank 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1440
         Width           =   5295
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox txtKet 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   7
         Top             =   2880
         Width           =   5295
      End
      Begin MSComCtl2.DTPicker dtpTglBKK 
         Height          =   330
         Left            =   5520
         TabIndex        =   0
         Top             =   340
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy hh:mm"
         Format          =   146931715
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
         TabIndex        =   6
         Top             =   2520
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo dcCaraBayar 
         Height          =   330
         Left            =   2160
         TabIndex        =   2
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
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
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         Top             =   3600
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dcTransaksi 
         Height          =   330
         Left            =   2160
         TabIndex        =   1
         Top             =   720
         Width           =   5295
         _ExtentX        =   9340
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
         TabIndex        =   20
         Top             =   2220
         Width           =   1260
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "No. Rekening"
         Height          =   210
         Left            =   240
         TabIndex        =   19
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Nama Bank"
         Height          =   210
         Left            =   240
         TabIndex        =   18
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Nama Transaksi"
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Jml. Bayar (Rp. )"
         Height          =   210
         Left            =   240
         TabIndex        =   15
         Top             =   2580
         Width           =   1350
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Transaksi"
         Height          =   210
         Left            =   4320
         TabIndex        =   14
         Top             =   400
         Width           =   1110
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   2940
         Width           =   945
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cara Bayar"
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   1140
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "No. Bukti Kas Keluar"
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   405
         Width           =   1635
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   23
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
      Picture         =   "frmStrukBuktiKasKeluarU.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   5880
      Picture         =   "frmStrukBuktiKasKeluarU.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmStrukBuktiKasKeluarU.frx":4B79
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
Attribute VB_Name = "frmStrukBuktiKasKeluarU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSimpan_Click()
    On Error GoTo errSave
    If funcCekValidasi = False Then Exit Sub
    If sp_BKK(dbcmd) = False Then Exit Sub
    cmdSimpan.Enabled = False
    cmdTambah.SetFocus
errSave:
End Sub

Private Sub cmdTambah_Click()
    txtNoBKK.Text = ""
    dtpTglBKK.Value = Format(Now, "dd/MM/yyyy HH:mm:ss")
    dcTransaksi.Text = ""
    dcCaraBayar.Text = ""
    txtNamaBank.Text = ""
    txtNoKartu.Text = ""
    txtNamaPenyetor.Text = ""
    txtJmlByr.Text = ""
    txtKet.Text = ""
    cmdSimpan.Enabled = True
    dtpTglBKK.SetFocus
End Sub

Private Sub cmdTutup_Click()
    Unload Me
    If blnStatusFrmUtama = True Then frmDaftarKasKeluar.Enabled = True
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDaftarKasKeluar.Enabled = True
End Sub

Private Sub mnucetak_Click()
    strCetak = "BKKUmum"
    frmCetakKwitansi.Show
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

Private Sub txtJmlByr_Change()
    If txtJmlByr.Text = "" Then txtJmlByr.Text = 0
    txtJmlByr = Format(txtJmlByr, "#,###,###,###,##0")
    txtJmlByr.SelStart = Len(txtJmlByr.Text)
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

'untuk mengecek validasi data yang akan disimpan
Private Function funcCekValidasi() As Boolean
    funcCekValidasi = False
    If dcTransaksi.Text = "" Then
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
Private Function sp_BKK(ByVal adoCommand As ADODB.Command) As Boolean
    On Error GoTo errSp_BKK
    Dim strLokal As String
    sp_BKK = False
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
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
        .CommandText = "dbo.Add_StrukBuktiKasKeluarUmum"
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
            MsgBox "Pemasukan data Struk Bukti Kas Keluar sukses", vbInformation, "Validasi"
            sp_BKK = True
            Call Add_HistoryLoginActivity("Add_StrukBuktiKasKeluarUmum")
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
