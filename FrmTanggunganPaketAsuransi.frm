VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form FrmTanggunganPaketAsuransi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Tarif Tanggungan Paket Asuransi"
   ClientHeight    =   8970
   ClientLeft      =   765
   ClientTop       =   1770
   ClientWidth     =   10350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTanggunganPaketAsuransi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   10350
   Begin VB.CommandButton cmdcetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   3000
      TabIndex        =   24
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   8750
      TabIndex        =   12
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   5880
      TabIndex        =   11
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   8400
      Width           =   1335
   End
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
      Height          =   6375
      Left            =   0
      TabIndex        =   14
      Top             =   1080
      Width           =   10335
      Begin VB.TextBox txtPersenTRSfromSelisih 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8280
         MaxLength       =   8
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtJmlTanggungan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8280
         MaxLength       =   12
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo dcPenjamin 
         Height          =   330
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   4095
         _ExtentX        =   7223
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
         Left            =   4680
         TabIndex        =   4
         Top             =   1320
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSDataListLib.DataCombo dcPaket 
         Height          =   330
         Left            =   360
         TabIndex        =   0
         Top             =   600
         Width           =   4095
         _ExtentX        =   7223
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
         Left            =   4680
         TabIndex        =   1
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
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
      Begin MSDataGridLib.DataGrid dgTanggunganPaketAsuransi 
         Height          =   4455
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "% TRS from Selisih"
         Height          =   210
         Left            =   8280
         TabIndex        =   20
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nama Penjamin"
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah Tanggungan"
         Height          =   210
         Left            =   8280
         TabIndex        =   18
         Top             =   1080
         Width           =   1650
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nama Kelas"
         Height          =   210
         Left            =   4680
         TabIndex        =   17
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nama Paket Asuransi"
         Height          =   210
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kelompok Pasien"
         Height          =   210
         Index           =   1
         Left            =   4680
         TabIndex        =   15
         Top             =   360
         Width           =   1365
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   13
      Top             =   7320
      Width           =   10335
      Begin VB.TextBox txtParameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1440
         TabIndex        =   7
         Top             =   300
         Width           =   3495
      End
      Begin MSDataListLib.DataCombo cboKlpPasien 
         Height          =   330
         Left            =   6720
         TabIndex        =   8
         Top             =   300
         Width           =   3375
         _ExtentX        =   5953
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kelompok Pasien"
         Height          =   210
         Index           =   2
         Left            =   5280
         TabIndex        =   23
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nama Paket"
         Height          =   210
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   21
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
      Picture         =   "FrmTanggunganPaketAsuransi.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8520
      Picture         =   "FrmTanggunganPaketAsuransi.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   1800
      Picture         =   "FrmTanggunganPaketAsuransi.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8745
   End
End
Attribute VB_Name = "FrmTanggunganPaketAsuransi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboKlpPasien_Change()
    Call subLoadGridSource
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
    If dgTanggunganPaketAsuransi.ApproxCount = 0 Then Exit Sub
'    FrmCetakTanggunganPaketAsuransi.Show
End Sub

Private Sub cmdHapus_Click()
On Error GoTo errLoad
    
    If Periksa("datacombo", dcPaket, "Nama paket kosong") = False Then Exit Sub
    If Periksa("datacombo", dcKelompokPasien, "Kelompok pasien kosong") = False Then Exit Sub
    If Periksa("datacombo", dcPenjamin, "Nama penjamin kosong") = False Then Exit Sub
    If Periksa("datacombo", dcKelas, "Nama kelas kosong") = False Then Exit Sub
    
    If MsgBox("Yakin data ini akan dihapus", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    If sp_TanggunganAsuransiPAKET("D") = False Then Exit Sub
    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
    Call cmdBatal_Click

Exit Sub
errLoad:
    MsgBox "Data digunakan, tidak dapat dihapus", vbOKOnly, "Informasi"
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo errLoad

    If Periksa("datacombo", dcPaket, "Nama paket kosong") = False Then Exit Sub
    If Periksa("datacombo", dcKelompokPasien, "Kelompok pasien kosong") = False Then Exit Sub
    If Periksa("datacombo", dcPenjamin, "Nama penjamin kosong") = False Then Exit Sub
    If Periksa("datacombo", dcKelas, "Nama kelas kosong") = False Then Exit Sub
    If Periksa("nilai", txtJmlTanggungan, "Jumlah tanggungan kosong") = False Then Exit Sub
    
    If sp_TanggunganAsuransiPAKET("A") = False Then Exit Sub
    MsgBox "Data telah disimpan..", vbInformation, "Informasi"
    Call cmdBatal_Click

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Set rs = Nothing
    Unload Me
End Sub

Private Sub dcKelas_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 13 Then
        strSQL = "SELECT KdKelas, DeskKelas FROM KelasPelayanan WHERE (DeskKelas LIKE '%" & dcKelas.Text & "%') and StatusEnabled=1 ORDER BY DeskKelas"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then dcKelas.BoundText = rs(0).Value: txtJmlTanggungan.SetFocus
    End If
Exit Sub
errLoad:
End Sub

Private Sub dcKelompokPasien_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then txtPersenTRSfromSelisih.SetFocus
End Sub

Private Sub dcPaket_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad

    If KeyAscii = 13 Then
        strSQL = "SELECT KdPaket, NamaPaket FROM PaketAsuransi WHERE (NamaPaket LIKE '%" & dcPaket.Text & "%') and StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then dcPaket.BoundText = rs(0).Value: dcKelompokPasien.SetFocus
    End If

Exit Sub
errLoad:
    Call msubPesanError
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
    If KeyAscii = 13 Then
        strSQL = "SELECT IdPenjamin, NamaPenjamin FROM Penjamin WHERE (NamaPenjamin LIKE '%" & dcPenjamin.Text & "%') and StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then dcPenjamin.BoundText = rs(0).Value: dcKelas.SetFocus
    End If
End Sub

Private Sub dgTanggunganPaketAsuransi_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdHapus.SetFocus
 On Error GoTo errLoad
    
    Call msubDcSource(dcPenjamin, rs, "SELECT IdPenjamin, NamaPenjamin From V_PenjaminPasien where StatusEnabled=1")
    
    With dgTanggunganPaketAsuransi
        If .ApproxCount = 0 Then Exit Sub
        dcPaket.BoundText = .Columns(5).Value
        dcKelompokPasien.BoundText = .Columns(6).Value
        dcPenjamin.BoundText = .Columns(7).Value
        dcKelas.BoundText = .Columns(8).Value
        txtJmlTanggungan.Text = Format(.Columns(4).Value, "#,###")
        txtPersenTRSfromSelisih.Text = .Columns(9).Value
    End With
Exit Sub
errLoad:
End Sub

Private Sub dgTanggunganPaketAsuransi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo errLoad
    
    Call msubDcSource(dcPenjamin, rs, "SELECT IdPenjamin, NamaPenjamin From V_PenjaminPasien where StatusEnabled=1")
    
    With dgTanggunganPaketAsuransi
        If .ApproxCount = 0 Then Exit Sub
        dcPaket.BoundText = .Columns(5).Value
        dcKelompokPasien.BoundText = .Columns(6).Value
        dcPenjamin.BoundText = .Columns(7).Value
        dcKelas.BoundText = .Columns(8).Value
        txtJmlTanggungan.Text = Format(.Columns(4).Value, "#,###")
        txtPersenTRSfromSelisih.Text = .Columns(9).Value
    End With
Exit Sub
errLoad:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo errLoad

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subKosong
    Call subLoadDcSource
    Call subLoadGridSource
 
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDcSource()
On Error GoTo errLoad

    Call msubDcSource(dcPaket, rs, "Select KdPaket, NamaPaket from PaketAsuransi where StatusEnabled=1 ORDER BY NamaPaket")
    Call msubDcSource(dcKelompokPasien, rs, "Select * from KelompokPasien where StatusEnabled=1")
    Call msubDcSource(cboKlpPasien, rs, "Select * from KelompokPasien where StatusEnabled=1")
    Call msubDcSource(dcPenjamin, rs, "SELECT IdPenjamin, NamaPenjamin From V_PenjaminPasien where StatusEnabled=1")
    Call msubDcSource(dcKelas, rs, "Select * from KelasPelayanan where StatusEnabled=1 ORDER BY DeskKelas")

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()
On Error GoTo errLoad
    strSQL = "select TOP 100 * from V_M_TanggunganPaketAsuransi WHERE NamaPaket LIKE '%" & txtParameter.Text & "%' and JenisPasien like '%" & cboKlpPasien.Text & "%'"
    Call msubRecFO(rs, strSQL)
    With dgTanggunganPaketAsuransi
        Set .DataSource = rs
        .Columns(0).Width = 1500
        .Columns(1).Width = 2500
        .Columns(2).Width = 2500
        .Columns(3).Width = 1400
        .Columns(4).Width = 1200
        .Columns(4).NumberFormat = "#,###"
        .Columns(4).Alignment = dbgRight
        .Columns(5).Width = 0
        .Columns(6).Width = 0
        .Columns(7).Width = 0
        .Columns(8).Width = 0
        .Columns(9).Width = 650
         
        .Columns(0).Caption = "Nama Paket"
        .Columns(1).Caption = "Kelompok Pasien"
        .Columns(2).Caption = "Penjamin"
        .Columns(3).Caption = "Kelas"
        .Columns(4).Caption = "Jml Tanggungan"
        .Columns(9).Caption = "% TRS"
    End With
Exit Sub
errLoad:
    Set rs = Nothing
    Call msubPesanError
End Sub

Private Sub subKosong()
    dcPaket.BoundText = ""
    dcKelompokPasien.BoundText = ""
    dcPenjamin.BoundText = ""
    dcKelas.Text = ""
    txtJmlTanggungan.Text = ""
    txtPersenTRSfromSelisih.Text = ""
End Sub

Private Sub txtJmlTanggungan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtJmlTanggungan_LostFocus()
    txtJmlTanggungan.Text = Format(txtJmlTanggungan.Text, "#,###")
End Sub

Private Sub txtParameter_Change()
    Call subLoadGridSource
End Sub

Private Function sp_TanggunganAsuransiPAKET(f_Status As String) As Boolean
    sp_TanggunganAsuransiPAKET = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdPaket", adVarChar, adParamInput, 3, dcPaket.BoundText)
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, dcKelompokPasien.BoundText)
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, dcPenjamin.BoundText)
        .Parameters.Append .CreateParameter("kdkelas", adChar, adParamInput, 2, dcKelas.BoundText)
        .Parameters.Append .CreateParameter("jmltanggungan", adCurrency, adParamInput, , CCur(txtJmlTanggungan.Text))
        .Parameters.Append .CreateParameter("PersenTRSfromSelisih", adDouble, adParamInput, , IIf(txtPersenTRSfromSelisih.Text = "", 0, txtPersenTRSfromSelisih.Text))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        
        .ActiveConnection = dbConn
        .CommandText = "AUD_TanggunganpaketAsuransi"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_TanggunganAsuransiPAKET = False
        End If
        
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then dgTanggunganPaketAsuransi.SetFocus
Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtPersenTRSfromSelisih_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcPenjamin.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub
