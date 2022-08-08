VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmTagihan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifist2000 - Daftar Tagihan Pasien"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTagihan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   14790
   Begin MSDataGridLib.DataGrid dgPasien 
      Height          =   5055
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   6120
      Width           =   14775
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   12960
         TabIndex        =   1
         Top             =   300
         Width           =   1695
      End
      Begin VB.CommandButton cmdTagihan 
         Caption         =   "&Tagihan Pasien"
         Height          =   495
         Left            =   11160
         TabIndex        =   0
         Top             =   300
         Width           =   1695
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12960
      Picture         =   "frmTagihan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmTagihan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTagihan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmTagihan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsB As New ADODB.recordset

Private Sub cmdTagihan_Click()
    Dim i As Integer
    On Error GoTo errLoad
    'cek pasien RI
    If dgPasien.ApproxCount = 0 Then Exit Sub

    strSQL = "SELECT * FROM RegistrasiIGD WHERE NoPendaftaran = '" & dgPasien.Columns("No. Registrasi").Value & "' AND StatusPulang = 'T'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then
        MsgBox "Pasien belum keluar dari IGD", vbCritical
        Exit Sub
    End If

    strSQL = "SELECT NamaRuangan FROM v_PasienAktifPakaiKamar WHERE NoPendaftaran='" _
    & dgPasien.Columns("No. Registrasi").Value & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then MsgBox "Pasien belum keluar dari Rawat Inap ( " & rs(0) & " )", vbCritical, "Validasi": dgPasien.SetFocus: Exit Sub

    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & dgPasien.Columns("No. Registrasi").Value & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    Else
        MsgBox "Lengkapi dahulu data penjamin pasien", vbCritical, "Validasi"
        Call cmdUbahKelPasien_Click
        Exit Sub
    End If

    If mstrKdPenjaminPasien <> "2222222222" Then
        strSQL = "SELECT * FROM PemakaianAsuransi WHERE NoPendaftaran='" & dgPasien.Columns("No. Registrasi").Value & "'"
        Call msubRecFO(rsB, strSQL)
        If rsB.RecordCount = 0 Then
            MsgBox "Lengkapi dahulu data penjamin pasien", vbCritical, "Validasi"
            Call cmdUbahKelPasien_Click
            Exit Sub
        End If
    End If

    If mstrKdPenjaminPasien <> "2222222222" Then
        Call frmCariPasien.PostingHutangPenjaminPasien_AU("A")
    End If
    

    strSQL = "SELECT * FROM RegistrasiIGD WHERE NoPendaftaran = '" & dgPasien.Columns("No. Registrasi").Value & "'"
    Call msubRecFO(rs, strSQL)
    
    If rs.RecordCount = 0 Then
        
        strSQLx = "Select * from V_DaftarPasienBelumBayar_New" & _
              " WHERE (NoCM like '%" & mstrNoCM & "%') AND TglPulang BETWEEN '" & Format(frmCariPasien.dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmCariPasien.dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'" & mstrFilter

        Call msubRecFO(rsx, strSQLx)
        
        For i = 1 To rsx.RecordCount
        
            If rsx("Ruangan").Value = "Gawat Darurat" Then
                MsgBox "Pasien belum membayar tagihan di transaksi IGD", vbCritical
                Exit Sub
            Else
                Call subLoadFormTP
         
            End If
            
        Next i
     
     Else
     
        Call subLoadFormTP
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdUbahKelPasien_Click()
    On Error GoTo hell
    If dgPasien.ApproxCount = 0 Then Exit Sub
    strSQL = "SELECT KdInstalasi FROM dbo.Ruangan WHERE KdRuangan = '" & dgPasien.Columns("KdRuanganAkhir") & "'"
    Call msubRecFO(rs, strSQL)
    mstrKdInstalasi = rs.Fields("KdInstalasi")
    '20090216 Ubah Jenis Harus Admin dan Verifikator
    Set rs = Nothing
    Call msubRecFO(rs, "select idPegawai from V_AdminKasir where idPegawai = '" & strIDPegawaiAktif & "'")
    If rs.EOF = True Then Exit Sub
    Call subLoadFormJP
    Exit Sub
hell:
End Sub

Private Sub cmdTutup_Click()
    On Error Resume Next
    Unload Me
    frmCariPasien.Enabled = True
End Sub


Private Sub dgPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTagihan.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subLoadDataPasien
End Sub

'untuk load data pasien di form transaksi pelayanan
Private Sub subLoadFormTP()
    On Error GoTo hell
    mstrNoPen = dgPasien.Columns("No. Registrasi").Value
    mstrNoCM = dgPasien.Columns(1).Value
    With frmTagihanPasien
        .Show
        .txtNoPendaftaran.Text = mstrNoPen
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = dgPasien.Columns("Nama Pasien").Value
        .txtSex.Text = dgPasien.Columns("JK").Value
        .txtJenisPasien.Text = dgPasien.Columns("Jenis Pasien").Value
        Call .txtNoPendaftaran_KeyPress(13)
    End With
    Exit Sub
hell:
End Sub

'untuk load data pasien
Public Sub subLoadDataPasien()
   
     strSQL = "Select * from V_DaftarPasienBelumBayar_New" & _
              " WHERE (NoCM like '%" & mstrNoCM & "%') AND TglPulang BETWEEN '" & Format(frmCariPasien.dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(frmCariPasien.dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'" & mstrFilter
   
    Call msubRecFO(rsB, strSQL)
    Set dgPasien.DataSource = rsB
    With dgPasien
        .Columns(0).Width = 1300
        .Columns(0).Caption = "No. Registrasi"
        .Columns(1).Width = 750 'NoCM
        .Columns(2).Width = 2750 'Nama Pasien
        .Columns(3).Width = 350 'JK
        .Columns(4).Width = 1500 'Umur
        .Columns(5).Width = 1750 'Jenis Pasien
        .Columns(6).Width = 1590 'Nama Penjamin
        .Columns(7).Width = 1590 'TglPendaftaran
        .Columns(8).Width = 2210 'Ruangan
        .Columns(9).Width = 1590 'TglPulang
        .Columns(10).Width = 0 'Alamat
        .Columns(11).Width = 500 'Tahun
        .Columns(12).Width = 0 'Bulan
        .Columns(13).Width = 0 'Hari
        .Columns(14).Width = 0 'KdRuangan
        .Columns(15).Width = 0 'KdInstalasi
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadFormJP()
    On Error GoTo hell
    mstrNoPen = dgPasien.Columns("No. Registrasi").Value
    mstrNoCM = dgPasien.Columns("No. CM").Value
    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)

    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If

    With frmUbahJenisPasien
        .Show
        .txtNamaFormPengirim.Text = Me.Name
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = dgPasien.Columns("Nama Pasien").Value
        If dgPasien.Columns("JK").Value = "P" Then
            .txtJK.Text = "Perempuan"
        Else
            .txtJK.Text = "Laki-laki"
        End If
        .txtThn.Text = dgPasien.Columns("UmurTahun").Value
        .txtBln.Text = dgPasien.Columns("UmurBulan").Value
        .txtHr.Text = dgPasien.Columns("UmurHari").Value
        .lblNoPendaftaran.Visible = False
        .txtNoPendaftaran.Visible = False
        .txtTglPendaftaran.Text = dgPasien.Columns("TglPendaftaran").Value
        .dcJenisPasien.BoundText = mstrKdJenisPasien
        .dcPenjamin.BoundText = mstrKdPenjaminPasien

    End With
    Exit Sub
hell:
End Sub

