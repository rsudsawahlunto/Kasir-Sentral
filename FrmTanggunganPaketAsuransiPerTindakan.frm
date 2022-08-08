VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form FrmTanggunganPaketAsuransiPerTindakan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Tanggungan Paket Asuransi Per Tindakan"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTanggunganPaketAsuransiPerTindakan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   10500
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   13
      Top             =   1005
      Width           =   10455
      Begin VB.TextBox txtPersenTRSfromSelisih 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8400
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtJmlTanggungan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6600
         MaxLength       =   12
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo dcPenjamin 
         Height          =   330
         Left            =   6600
         TabIndex        =   2
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcKelas 
         Height          =   330
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcPelayanan 
         Height          =   330
         Left            =   2640
         TabIndex        =   4
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcKelompokPasien 
         Height          =   330
         Left            =   3840
         TabIndex        =   1
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcPaket 
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "% TRS from Selisih"
         Height          =   210
         Left            =   8400
         TabIndex        =   21
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nama Paket Asuransi"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kelompok Pasien"
         Height          =   210
         Index           =   1
         Left            =   3840
         TabIndex        =   19
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pelayanan"
         Height          =   210
         Index           =   0
         Left            =   2640
         TabIndex        =   16
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Pelayanan"
         Height          =   210
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah Tanggungan"
         Height          =   210
         Left            =   6600
         TabIndex        =   17
         Top             =   1080
         Width           =   1650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nama Penjamin"
         Height          =   210
         Index           =   0
         Left            =   6600
         TabIndex        =   14
         Top             =   360
         Width           =   1245
      End
   End
   Begin MSDataGridLib.DataGrid dgTanggunganAsuransiNonPaket 
      Height          =   4575
      Left            =   0
      TabIndex        =   7
      Top             =   3000
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   8070
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
            LCID            =   1033
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
            LCID            =   1033
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   18
      Top             =   7560
      Width           =   10455
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   4440
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtCariPelayanan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   5640
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   9150
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   6855
         TabIndex        =   10
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   8050
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cari Pelayanan"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   300
         Width           =   1155
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
      Picture         =   "FrmTanggunganPaketAsuransiPerTindakan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8640
      Picture         =   "FrmTanggunganPaketAsuransiPerTindakan.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "FrmTanggunganPaketAsuransiPerTindakan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "FrmTanggunganPaketAsuransiPerTindakan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub subLoadDcSource()
On Error GoTo errLoad

    Call msubDcSource(dcPaket, rs, "Select KdPaket, NamaPaket from PaketAsuransi where StatusEnabled=1 ORDER BY NamaPaket")
    Call msubDcSource(dcKelompokPasien, rs, "Select * from KelompokPasien where StatusEnabled=1")
    Call msubDcSource(dcPenjamin, rs, "SELECT IdPenjamin, NamaPenjamin From V_PenjaminPasien where StatusEnabled=1")
    Call msubDcSource(dcKelas, rs, "Select * from KelasPelayanan where StatusEnabled=1 ORDER BY DeskKelas")
    Call msubDcSource(dcPelayanan, rs, "Select * from ListPelayananRS where StatusEnabled=1 ORDER BY NamaPelayanan")

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdBatal_Click()
On Error GoTo errLoad
    Call subKosong
    Call subLoadDcSource
    Call subLoadGridSource
    dcPaket.SetFocus
Exit Sub
errLoad:
End Sub

Private Sub cmdCetak_Click()
    If dgTanggunganAsuransiNonPaket.ApproxCount = 0 Then Exit Sub
'    FrmCetakPaketAsuransiPerTindakan.Show
End Sub

Private Sub cmdHapus_Click()
On Error GoTo errLoad
    
    If Periksa("datacombo", dcKelompokPasien, "Kelompok pasien kosong") = False Then Exit Sub
    If Periksa("datacombo", dcPenjamin, "Nama penjamin kosong") = False Then Exit Sub
    If Periksa("datacombo", dcKelas, "Nama kelas kosong") = False Then Exit Sub
    If Periksa("datacombo", dcPelayanan, "Nama pelayanan kosong") = False Then Exit Sub
    
    If MsgBox("Apakah anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    If sp_TanggunganAsuransiPaketPerTindakan("D") = False Then Exit Sub
    
    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
    Call cmdBatal_Click

Exit Sub
errLoad:
    MsgBox "Penghapusan Gagal, Data Sudah Terpakai !", vbOKOnly, "Informasi"
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo errLoad
    If Periksa("datacombo", dcPaket, "Paket asuransi kosong") = False Then Exit Sub
    If Periksa("datacombo", dcKelompokPasien, "Kelompok pasien kosong") = False Then Exit Sub
    If Periksa("datacombo", dcPenjamin, "Nama penjamin kosong") = False Then Exit Sub
    If Periksa("datacombo", dcKelas, "Nama kelas kosong") = False Then Exit Sub
    If Periksa("datacombo", dcPelayanan, "Nama pelayanan kosong") = False Then Exit Sub
    If Periksa("nilai", txtJmlTanggungan, "Jumlah tanggungan kosong") = False Then Exit Sub
    
    
    Set rs = Nothing
    strSQL = "SELECT DISTINCT dbo.DaftarTMNonTanggungan.IdPenjamin, dbo.DaftarTMNonTanggungan.KdPelayananRS " & _
             "FROM dbo.TanggunganAsuransiNonPaket INNER JOIN " & _
             "dbo.ConvertPaketAsuransiToPelayanan ON dbo.TanggunganAsuransiNonPaket.KdPaket = dbo.ConvertPaketAsuransiToPelayanan.KdPaket FULL OUTER JOIN " & _
             "dbo.DaftarTMNonTanggungan ON dbo.TanggunganAsuransiNonPaket.KdKelompokPasien = dbo.DaftarTMNonTanggungan.KdKelompokPasien " & _
             "Where dbo.DaftarTMNonTanggungan.IdPenjamin = '" & dcPenjamin.BoundText & "' and dbo.DaftarTMNonTanggungan.KdPelayananRS = '" & dcPelayanan.BoundText & "'"
    Call msubRecFO(rs, strSQL)
    
    If rs.EOF = True Then
        If sp_TanggunganAsuransiPaketPerTindakan("A") = False Then Exit Sub
        MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    Else
        MsgBox "Pelayanan " & dcPelayanan.Text & " sudah di input di settingan lain", vbCritical
        dcPelayanan.Text = ""
        dcPelayanan.SetFocus
    End If
    
    Call cmdBatal_Click
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcKelas_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
    If dcKelas.MatchedWithList = True Then dcPelayanan.SetFocus
        If Len(Trim(dcKelas.Text)) = 0 Then dcPelayanan.SetFocus: Exit Sub
        strSQL = "SELECT KdKelas, DeskKelas FROM KelasPelayanan WHERE (DeskKelas LIKE '%" & dcKelas.Text & "%') and StatusEnabled=1 ORDER BY DeskKelas"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then dcKelas.Text = "": Exit Sub
        dcKelas.BoundText = rs(0).Value
        dcPelayanan.SetFocus
    End If
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKelas_LostFocus()
If dcKelas.MatchedWithList = False Then dcKelas.Text = ""
End Sub

Private Sub dcKelompokPasien_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
    If dcKelompokPasien.MatchedWithList = True Then dcPenjamin.SetFocus
    End If
End Sub

Private Sub dcKelompokPasien_LostFocus()
If dcKelompokPasien.MatchedWithList = False Then dcKelompokPasien.Text = ""
End Sub

Private Sub dcPaket_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
    If dcPaket.MatchedWithList = True Then dcKelompokPasien.SetFocus
        If Len(Trim(dcPaket.Text)) = 0 Then dcKelompokPasien.SetFocus: Exit Sub
        strSQL = "SELECT KdPaket, NamaPaket FROM PaketAsuransi WHERE (NamaPaket LIKE '%" & dcPaket.Text & "%') and StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then dcPaket.Text = "": Exit Sub
        dcPaket.BoundText = rs(0).Value
        dcKelompokPasien.SetFocus
    End If

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPaket_LostFocus()
If dcPaket.MatchedWithList = False Then dcPaket.Text = ""
End Sub

Private Sub dcPelayanan_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
    If dcPelayanan.MatchedWithList = True Then txtJmlTanggungan.SetFocus
        If Len(Trim(dcPelayanan.Text)) = 0 Then txtJmlTanggungan.SetFocus: Exit Sub
        strSQL = "SELECT KdPelayananRS FROM ListPelayananRS WHERE (NamaPelayanan LIKE '%" & dcPelayanan.Text & "%') and StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then dcPelayanan.Text = "": Exit Sub
        dcPelayanan.BoundText = rs(0).Value
        txtJmlTanggungan.SetFocus
    End If
Exit Sub
errLoad:
    Call msubPesanError
End Sub


Private Sub dcPelayanan_LostFocus()
If dcPelayanan.MatchedWithList = False Then dcPelayanan.Text = ""
End Sub

Private Sub dcPenjamin_GotFocus()
On Error GoTo errLoad
Dim tempKode As String
    
    tempKode = dcPenjamin.BoundText
    strSQL = "SELECT IdPenjamin, NamaPenjamin From V_PenjaminPasien WHERE (KdKelompokPasien = '" & dcKelompokPasien.BoundText & "') and StatusEnabled=1"
    Call msubDcSource(dcPenjamin, rs, strSQL)
    dcPenjamin.BoundText = tempKode

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPenjamin_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
    If dcPenjamin.MatchedWithList = True Then dcKelas.SetFocus
        If Len(Trim(dcPenjamin.Text)) = 0 Then dcKelas.SetFocus: Exit Sub
        strSQL = "SELECT IdPenjamin, NamaPenjamin FROM Penjamin WHERE (NamaPenjamin LIKE '%" & dcPenjamin.Text & "%') and StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then dcPenjamin.Text = "": Exit Sub
        dcPenjamin.BoundText = rs(0).Value
        dcKelas.SetFocus
    End If
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPenjamin_LostFocus()
If dcPenjamin.MatchedWithList = False Then dcPenjamin.Text = ""
End Sub

Private Sub dgTanggunganAsuransiNonPaket_KeyPress(KeyAscii As Integer)

On Error GoTo hell
    If KeyAscii = 13 Then cmdHapus.SetFocus
    With dgTanggunganAsuransiNonPaket
        If .ApproxCount = 0 Then Exit Sub
        dcPaket.BoundText = .Columns(6).Value
        dcKelompokPasien.BoundText = .Columns(7).Value
        dcPenjamin.BoundText = .Columns(8).Value
        dcPelayanan.BoundText = .Columns(9).Value
        dcKelas.BoundText = .Columns(10).Value
        txtJmlTanggungan.Text = Format(.Columns(5).Value, "#,###")
        txtPersenTRSfromSelisih.Text = .Columns(11).Value
    End With
Exit Sub
hell:
End Sub

Private Sub dgTanggunganAsuransiNonPaket_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo hell
    
    With dgTanggunganAsuransiNonPaket
        If .ApproxCount = 0 Then Exit Sub
        dcPaket.BoundText = .Columns(6).Value
        dcKelompokPasien.BoundText = .Columns(7).Value
        dcPenjamin.BoundText = .Columns(8).Value
        dcPelayanan.BoundText = .Columns(9).Value
        dcKelas.BoundText = .Columns(10).Value
        txtJmlTanggungan.Text = Format(.Columns(5).Value, "#,###")
        txtPersenTRSfromSelisih.Text = .Columns(11).Value
    End With
Exit Sub
hell:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo errLoad

     Call centerForm(Me, MDIUtama)
     Call PlayFlashMovie(Me)
     Call cmdBatal_Click
  
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()
On Error GoTo errLoad
    strSQL = "select * from V_TanggunganPaketAsuransiPerTindakan WHERE NamaPelayanan LIKE '%" & txtCariPelayanan.Text & "%'"
    Call msubRecFO(rs, strSQL)
    With dgTanggunganAsuransiNonPaket
        Set .DataSource = rs
        .Columns(0).Width = 1500
        .Columns(1).Width = 1500
        .Columns(2).Width = 1250
        .Columns(3).Width = 2500
        .Columns(4).Width = 1250
        .Columns(5).Width = 1200
        .Columns(6).Width = 0
        .Columns(7).Width = 0
        .Columns(8).Width = 0
        .Columns(9).Width = 0
        .Columns(10).Width = 0
        .Columns(11).Width = 650
         
        .Columns(0).Caption = "Paket Asuransi"
        .Columns(1).Caption = "Kelompok Pasien"
        .Columns(2).Caption = "Penjamin"
        .Columns(3).Caption = "Pelayananan"
        .Columns(4).Caption = "Kelas"
        .Columns(5).Caption = "Jml Tanggungan"
        .Columns(11).Caption = "% TRS"
    End With
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub subKosong()
    dcPaket.BoundText = ""
    dcKelompokPasien.BoundText = ""
    dcPenjamin.BoundText = ""
    dcKelas.BoundText = ""
    dcPelayanan.BoundText = ""
    txtJmlTanggungan.Text = ""
    txtPersenTRSfromSelisih.Text = ""
End Sub

Private Sub txtCariPelayanan_KeyPress(KeyAscii As Integer)
    Call subLoadGridSource
    Call SetKeyPressToChar(KeyAscii)
    If KeyAscii = 13 Then dgTanggunganAsuransiNonPaket.SetFocus
    
End Sub

Private Sub txtJmlTanggungan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPersenTRSfromSelisih.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtJmlTanggungan_LostFocus()
    txtJmlTanggungan.Text = Format(txtJmlTanggungan.Text, "#,###")
End Sub

Private Function sp_TanggunganAsuransiPaketPerTindakan(f_Status As String) As Boolean
    sp_TanggunganAsuransiPaketPerTindakan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdPaket", adVarChar, adParamInput, 3, dcPaket.BoundText)
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, dcKelompokPasien.BoundText)
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, dcPenjamin.BoundText)
        .Parameters.Append .CreateParameter("kdkelas", adChar, adParamInput, 2, dcKelas.BoundText)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, dcPelayanan.BoundText)
        .Parameters.Append .CreateParameter("jmltanggungan", adCurrency, adParamInput, , CCur(txtJmlTanggungan.Text))
        .Parameters.Append .CreateParameter("PersenTRSfromSelisih", adDouble, adParamInput, , IIf(txtPersenTRSfromSelisih.Text = "", 0, txtPersenTRSfromSelisih.Text))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_TanggunganAsuransiPaketPerTindakan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_TanggunganAsuransiPaketPerTindakan = False
        End If
        
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub txtPersenTRSfromSelisih_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
    Call SetKeyPressToNumber(KeyAscii)
End Sub
