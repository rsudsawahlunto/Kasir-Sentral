VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmValidasiUlangGagalBayar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifist2000 - Validasi Ulang Gagal Bayar"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmValidasiUlangGagalBayar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   12165
   Begin VB.TextBox txtNoBKM 
      Height          =   375
      Left            =   6120
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid dgPembayaranOtomatis 
      Height          =   4935
      Left            =   0
      TabIndex        =   10
      Top             =   2520
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   8705
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
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   12135
      Begin VB.OptionButton Option1 
         Caption         =   "No Struk is NULL"
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton optNoBKMNull 
         Caption         =   "No BKM is NULL"
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.Frame fraPeriode 
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6240
         TabIndex        =   13
         Top             =   120
         Width           =   5775
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   127795203
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   2
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   127795203
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   14
            Top             =   315
            Width           =   255
         End
      End
      Begin VB.Label lblJumData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data 0/0"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   7560
      Width           =   12135
      Begin VB.TextBox txtCariNoCM 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   5130
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdHapusData 
         Caption         =   "&HapusData"
         Height          =   495
         Left            =   6840
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   10260
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdBayarOtomatis 
         Caption         =   "Bayar &Ulang"
         Height          =   495
         Left            =   8550
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Masukkan No CM"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   16
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
      Picture         =   "frmValidasiUlangGagalBayar.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   10320
      Picture         =   "frmValidasiUlangGagalBayar.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmValidasiUlangGagalBayar.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmValidasiUlangGagalBayar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFilter As String
Dim rsB As New ADODB.recordset
Dim rsJenisPasien As New ADODB.recordset
Dim QJenisPasien As String
Dim DTglStruk As Date
Dim DTglBKM As Date
Dim sKdRuanganKasir As String

Public Function sp_PembatalanStrukPelayananKasir(f_KdRuangan As String) As Boolean
    On Error GoTo errLoad

    sp_PembatalanStrukPelayananKasir = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        If optNoBKMNull.Value = True Then
            .Parameters.Append .CreateParameter("NoBKM", adChar, adParamInput, 10, Null)
            .Parameters.Append .CreateParameter("NoStruk", adChar, adParamInput, 10, dgPembayaranOtomatis.Columns("NoStruk"))
        Else
            .Parameters.Append .CreateParameter("NoBKM", adChar, adParamInput, 10, dgPembayaranOtomatis.Columns("NoBKM"))
            .Parameters.Append .CreateParameter("NoStruk", adChar, adParamInput, 10, Null)
        End If
        .Parameters.Append .CreateParameter("PembayaranKe", adTinyInt, adParamInput, , 1)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "Add_PembatalanStrukPelayananKasir"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Add_PembatalanStrukPelayananKasir")
        End If
        Set dbcmd = Nothing
        Call deleteADOCommandParameters(dbcmd)
    End With
    Call cmdCari_Click

    Exit Function
errLoad:
    sp_PembatalanStrukPelayananKasir = False
    Call msubPesanError
End Function

'Store procedure untuk mengisi struk billing pasien
Private Function sp_AddStrukBuktiKasMasuk() As Boolean
    On Error GoTo errLoad
    Dim strLokal As String
    sp_AddStrukBuktiKasMasuk = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        If optNoBKMNull.Value = True Then
            .Parameters.Append .CreateParameter("TglBKM", adDate, adParamInput, , Format(DTglStruk, "yyyy/MM/dd HH:mm:ss"))
        Else
            .Parameters.Append .CreateParameter("TglBKM", adDate, adParamInput, , Format(DTglBKM, "yyyy/MM/dd HH:mm:ss"))
        End If
        .Parameters.Append .CreateParameter("KdCaraBayar", adChar, adParamInput, 2, "01")
        .Parameters.Append .CreateParameter("KdJenisKartu", adChar, adParamInput, 2, Null)
        .Parameters.Append .CreateParameter("NamaBank", adVarChar, adParamInput, 100, Null)
        .Parameters.Append .CreateParameter("NoKartu", adVarChar, adParamInput, 50, Null)
        .Parameters.Append .CreateParameter("AtasNama", adVarChar, adParamInput, 50, Null)
        .Parameters.Append .CreateParameter("JmlBayar", adCurrency, adParamInput, , mcurAll_HrsDibyr)
        .Parameters.Append .CreateParameter("Administrasi", adCurrency, adParamInput, , 0)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, Null)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, sKdRuanganKasir)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, noidpegawai)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("OutputNoBKM", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_StrukBuktiKasMasukPelayananPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("RETURN_VALUE").Value <> 0 Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Struk Billing Pasien", vbCritical, "Validasi"
            sp_AddStrukBuktiKasMasuk = False
        Else
            If Not IsNull(.Parameters("OutputNoBKM").Value) Then txtNoBKM.Text = .Parameters("OutputNoBKM").Value
            If Len(txtNoBKM.Text) = 0 Then
                If optNoBKMNull.Value = True Then
                    strLokal = "SELECT NoBKM from StrukBuktiKasMasuk where tglBKM = '" & Format(DTglStruk, "yyyy/MM/dd HH:mm:ss") & "' and kdRuangan = sKdRuanganKasir and idUser = '" & noidpegawai & "'"
                Else
                    strLokal = "SELECT NoBKM from StrukBuktiKasMasuk where tglBKM = '" & Format(DTglBKM, "yyyy/MM/dd HH:mm:ss") & "' and kdRuangan = sKdRuanganKasir and idUser = '" & noidpegawai & "'"
                End If
                Call msubRecFO(rs, strLokal)
                txtNoBKM.Text = rs("NoBKM").Value
            End If
            Call Add_HistoryLoginActivity("Add_StrukBuktiKasMasukPelayananPasien")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    sp_AddStrukBuktiKasMasuk = False
    Call msubPesanError("-Add_StrukBuktiKasMasukPelayananPasien")
End Function

'Store procedure untuk mengisi struk billing pasien
Private Function sp_AddStruk(ByVal adoCommand As ADODB.Command, strStsByr As String) As Boolean
    On Error GoTo errLoad
    Dim strLokal As String
    sp_AddStruk = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoBKM", adChar, adParamInput, 10, txtNoBKM.Text)
        .Parameters.Append .CreateParameter("OutputNoStruk", adChar, adParamOutput, 10, Null)
        If optNoBKMNull.Value = True Then
            .Parameters.Append .CreateParameter("TglStruk", adDate, adParamInput, , Format(DTglStruk, "yyyy/MM/dd HH:mm:ss"))
        Else
            .Parameters.Append .CreateParameter("TglStruk", adDate, adParamInput, , Format(DTglBKM, "yyyy/MM/dd HH:mm:ss"))
        End If
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, mstrNoCM)
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, mstrKdJenisPasien)
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, mstrKdPenjaminPasien)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, sKdRuanganKasir)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, noidpegawai)
        .Parameters.Append .CreateParameter("TotalBiaya", adCurrency, adParamInput, , CCur(mcurBayar))
        .Parameters.Append .CreateParameter("JmlHutangPenjamin", adCurrency, adParamInput, , CCur(mcurAll_TP))
        .Parameters.Append .CreateParameter("JmlTanggunganRS", adCurrency, adParamInput, , CCur(mcurAll_TRS))
        .Parameters.Append .CreateParameter("JmlPembebasan", adCurrency, adParamInput, , CCur(mcurAll_Pemb))
        .Parameters.Append .CreateParameter("JmlHrsDibayar", adCurrency, adParamInput, , CCur(mcurAll_HrsDibyr))
        .Parameters.Append .CreateParameter("JmlDiscount", adCurrency, adParamInput, , "0")

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_NoStrukPelayananPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Struk Billing Pasien", vbCritical, "Validasi"
            sp_AddStruk = False
        Else
            If Not IsNull(.Parameters("OutputNoStruk").Value) Then mstrNoStruk = .Parameters("OutputNoStruk").Value
            If Len(mstrNoStruk) = 0 Then
                If optNoBKMNull.Value = True Then
                    strLokal = "SELECT NoStruk from StrukPelayananPasien where tglStruk = '" & Format(DTglStruk, "yyyy/MM/dd HH:mm:ss") & "' and NoPendaftaran = '" & mstrNoPen & "' and NoCM = '" & mstrNoCM & "' and idUser = '" & noidpegawai & "'"
                Else
                    strLokal = "SELECT NoStruk from StrukPelayananPasien where tglStruk = '" & Format(DTglBKM, "yyyy/MM/dd HH:mm:ss") & "' and NoPendaftaran = '" & mstrNoPen & "' and NoCM = '" & mstrNoCM & "' and idUser = '" & noidpegawai & "'"
                End If
                Call msubRecFO(rs, strLokal)
                mstrNoStruk = rs("NoStruk").Value
            End If
            Call Add_HistoryLoginActivity("Add_NoStrukPelayananPasien")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Function
errLoad:
    msubPesanError ("-Add_NoStrukPelayananPasien")
End Function

Private Sub cmdBayarOtomatis_Click()
    On Error GoTo hell:
    If sKdRuanganKasir = "" Then Exit Sub
    cmdHapusData.Enabled = True
    cmdBayarOtomatis.Enabled = False
    QJenisPasien = "SELECT NoCM,KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rsJenisPasien, QJenisPasien)
    If rsJenisPasien.EOF = False Then
        mstrNoCM = rsJenisPasien("NoCM").Value
        mstrKdJenisPasien = rsJenisPasien("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rsJenisPasien("IdPenjamin")), "2222222222", rsJenisPasien("IdPenjamin"))
    End If
    strSQL = "SELECT     SUM(BiayaTotal) AS BiayaTotal, SUM(TotalHutangPenjamin) AS TotalHutangPenjamin, SUM(TotalTanggunganRS) AS TotalTanggunganRS, " & _
    " SUM(TotalPembebasan) AS TotalPembebasan, SUM(TotalHarusDibayar) AS TotalHarusDibayar " & _
    " From V_RincianTotalDetailBiayaPelayanan WHERE NoPendaftaran='" & mstrNoPen & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        curTarif = 0
        curTP = 0
        curTRS = 0
        curPemb = 0
        mcurAll_HrsDibyr = 0
        mcurBayar = 0
        mcurAll_TP = 0
        mcurAll_TRS = 0
        mcurAll_Pemb = 0
    Else
        curTarif = IIf(IsNull(rs("BiayaTotal")), 0, rs("BiayaTotal"))
        curTP = IIf(IsNull(rs("TotalHutangPenjamin")), 0, rs("TotalHutangPenjamin"))
        curTRS = IIf(IsNull(rs("TotalTanggunganRS")), 0, rs("TotalTanggunganRS"))
        curPemb = IIf(IsNull(rs("TotalPembebasan")), 0, rs("TotalPembebasan"))
        mcurAll_HrsDibyr = curTarif - (curTP + curTRS + curPemb)
        mcurBayar = curTarif
        mcurAll_TP = curTP
        mcurAll_TRS = curTRS
        mcurAll_Pemb = curPemb
    End If
    Set rs = Nothing
    If optNoBKMNull.Value = True Then
        strSQL = "SELECT  NoPendaftaran, TglBKM" & _
        " FROM  ConvertSBKMToNoPendaftaran WHERE NoPendaftaran ='" & mstrNoPen & "' AND TglBKM = '" & Format(DTglStruk, "yyyy/MM/dd HH:mm:ss") & "'"
    Else
        strSQL = "SELECT  NoPendaftaran, TglBKM" & _
        " FROM  ConvertSBKMToNoPendaftaran WHERE NoPendaftaran ='" & mstrNoPen & "' AND TglBKM = '" & Format(DTglBKM, "yyyy/MM/dd HH:mm:ss") & "'"
    End If
    Call msubRecFO(rs, strSQL)

    If rs.EOF = False Then
        If optNoBKMNull.Value = True Then
            Set rs = Nothing
            strSQL = "Delete ConvertSBKMToNoPendaftaran WHERE NoPendaftaran ='" & mstrNoPen & "' AND TglBKM = '" & Format(DTglStruk, "yyyy/MM/dd HH:mm:ss") & "'"
            dbConn.Execute strSQL
        Else
            Set rs = Nothing
            strSQL = "Delete ConvertSBKMToNoPendaftaran WHERE NoPendaftaran ='" & mstrNoPen & "' AND TglBKM = '" & Format(DTglBKM, "yyyy/MM/dd HH:mm:ss") & "'"
            dbConn.Execute strSQL
        End If
    End If
    If sp_AddStrukBuktiKasMasuk() = False Then Exit Sub
    If sp_AddStruk(dbcmd, 1) = False Then Exit Sub
    mstrNoBKM = txtNoBKM.Text
    fStatusPiutang = "TM"
    fStatusBayarSemua = "Y"
    MsgBox "Validasi Data yang Gagal Bayar BERHASIL", vbInformation, "Informasi"
    Exit Sub
hell:
    msubPesanError
End Sub

Public Sub cmdCari_Click()
    On Error Resume Next
    MousePointer = vbHourglass
    Call subLoadPembayaranOtomatis
    MousePointer = vbDefault
End Sub

Private Sub cmdCetak_Click()
    If dgPembayaranOtomatis.ApproxCount = 0 Then Exit Sub
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
End Sub

Private Sub cmdRincianTotal_Click()
    Me.Enabled = False
    frmRincianTotalDetailBiaya.Show
End Sub

Private Sub cmdHapusData_Click()
    On Error Resume Next
    cmdBayarOtomatis.Enabled = True
    cmdHapusData.Enabled = False
    If dgPembayaranOtomatis.ApproxCount = 0 Then Exit Sub
    mstrNoPen = dgPembayaranOtomatis.Columns("NoPendaftaran").Value
    sKdRuanganKasir = dgPembayaranOtomatis.Columns("KdRuanganKasir")
    If optNoBKMNull.Value = True Then
        DTglStruk = dgPembayaranOtomatis.Columns("TglStruk")
    Else
        DTglBKM = dgPembayaranOtomatis.Columns("NoBKM")
    End If
    If optNoBKMNull.Value = True Then
        If sp_PembatalanStrukPelayananKasir(dgPembayaranOtomatis.Columns("KdRuanganKasir")) = False Then Exit Sub
    Else
        If sp_PembatalanStrukPelayananKasir(dgPembayaranOtomatis.Columns("KdRuanganKasir")) = False Then Exit Sub
    End If
    MsgBox "Hapus Data gagal bayar Berhasil - Untuk Validasi Ulang Tekan Tombol Bayar Ulang!!", vbInformation + vbOKOnly, "Informasi"
    cmdBayarOtomatis.SetFocus
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgPembayaranOtomatis_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    LblJumData.Caption = "Data " & dgPembayaranOtomatis.Bookmark & "/" & dgPembayaranOtomatis.ApproxCount
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpAwal.Value = Format(Now, "dd MMMM yyyy 00:00:00")
    dtpAkhir.Value = Format(Now, "dd MMMM yyyy 23:59:59")
    mstrFilter = ""
    cmdBayarOtomatis.Enabled = False
    optNoBKMNull.Value = True
    Call subLoadPembayaranOtomatis
End Sub

'untuk load data pasien di form transaksi pelayanan
Private Sub subLoadFormTP()
    On Error GoTo hell
    mstrNoPen = dgDaftarPasien.Columns("NoPendaftaran").Value
    With frmTagihanPasien
        .Show
        .txtNoPendaftaran.Text = mstrNoPen
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = dgDaftarPasien.Columns("NamaPasien").Value
        .txtSex.Text = dgDaftarPasien.Columns("JK").Value
        .txtThn.Text = dgDaftarPasien.Columns("UmurTahun")
        .txtBln.Text = dgDaftarPasien.Columns("UmurBulan")
        .txtHari.Text = dgDaftarPasien.Columns("UmurHari")
        .txtJenisPasien.Text = dgDaftarPasien.Columns("JenisPasien").Value
        Call .txtNoPendaftaran_KeyPress(13)
    End With
hell:
End Sub

'untuk load data pasien
Private Sub subLoadPembayaranOtomatis()
    If optNoBKMNull.Value = True Then
        strSQL = "SELECT * " & _
        " FROM V_StrukGagalGenerateNoBKM" & _
        " WHERE TglStruk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:59") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "' " & mstrFilter
        Call msubRecFO(rs, strSQL)
        Set dgPembayaranOtomatis.DataSource = rs
        With dgPembayaranOtomatis
            .Columns("NoBKM").Width = 1500
            .Columns("NoStruk").Width = 1500
            .Columns("NoPendaftaran").Width = 1500
            .Columns("No. CM").Width = 1500
            .Columns("RuanganKasir").Width = 2500
            .Columns("KdRuanganKasir").Width = 0
            .Columns("KdKelompokPasien").Width = 0
            .Columns("IdPenjamin").Width = 0
        End With
    Else
        strSQL = "SELECT * " & _
        " FROM V_SBKMGagalGenerateNoStruk" & _
        " WHERE TglBKM BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd hh:mm:59") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd hh:mm:59") & "' " & mstrFilter
        Call msubRecFO(rs, strSQL)
        Set dgPembayaranOtomatis.DataSource = rs
        With dgPembayaranOtomatis
            .Columns("NoBKM").Width = 1500
            .Columns("NoStruk").Width = 1500
            .Columns("RuanganKasir").Width = 2500
            .Columns("Keterangan").Width = 3000
            .Columns("KdRuanganKasir").Width = 0
        End With
    End If
End Sub

Private Function sp_BayarOtomatis(f_Pendaftaran As String, f_NoCM As String) As Boolean
    On Error GoTo errLoad
    sp_BayarOtomatis = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglStruk", adDate, adParamInput, , Format(DTPTglTransfer.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdCaraBayar", adChar, adParamInput, 2, dcCaraBayar.BoundText)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_Pendaftaran)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, f_NoCM)
        .Parameters.Append .CreateParameter("Keterangan", adChar, adParamInput, 100, txtKeterangan.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "Add_StrukPelayananPasienBayarByBackOffice"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam Pembayaran otomatis pasien", vbCritical, "Validasi"
            sp_BayarOtomatis = False
            MousePointer = vbDefault
        Else
            Call Add_HistoryLoginActivity("Add_StrukPelayananPasienBayarByBackOffice")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    msubPesanError
    sp_BayarOtomatis = False
    cmdSimpan.Enabled = True
    MousePointer = vbDefault
End Function

Private Function sp_Loop_AddTransferBPOAToHutangPenjamin(f_Pendaftaran As String) As Boolean
    On Error GoTo errLoad
    sp_Loop_AddTransferBPOAToHutangPenjamin = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_Pendaftaran)
        .Parameters.Append .CreateParameter("TglTransfer", adDate, adParamInput, , Format(DTPTglTransfer.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "Loop_AddTransferBPOAToHutangPenjamin"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam transfer hutang penjamin", vbCritical, "Validasi"
            sp_Loop_AddTransferBPOAToHutangPenjamin = False
            cmdSettingData.Enabled = True
        Else
            Call Add_HistoryLoginActivity("Loop_AddTransferBPOAToHutangPenjamin")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    msubPesanError
    MousePointer = vbDefault
    sp_Loop_AddTransferBPOAToHutangPenjamin = False
    cmdSettingData.Enabled = True
End Function

Private Function sp_Loop_AddTransferBPTMToHutangPenjamin(f_Pendaftaran As String) As Boolean
    On Error GoTo errLoad

    sp_Loop_AddTransferBPTMToHutangPenjamin = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_Pendaftaran)
        .Parameters.Append .CreateParameter("TglTransfer", adDate, adParamInput, , Format(DTPTglTransfer.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "Loop_AddTransferBPTMToHutangPenjamin"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam transfer hutang penjamin", vbCritical, "Validasi"
            sp_Loop_AddTransferBPTMToHutangPenjamin = False
            MousePointer = vbDefault
        Else
            Call Add_HistoryLoginActivity("Loop_AddTransferBPTMToHutangPenjamin")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    msubPesanError
    sp_Loop_AddTransferBPTMToHutangPenjamin = False
    cmdSettingData.Enabled = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    blnFrmCariPasien = False
End Sub

Private Sub Option1_Click()
    Call optNoBKMNull_Click
End Sub

Private Sub optNoBKMNull_Click()
    If optNoBKMNull.Value = True Then
        fraPeriode.Caption = "Periode - Tanggal Struk"
        cmdHapusData.Enabled = True
        cmdBayarOtomatis.Enabled = False
        txtCariNoCM.Visible = True
        Label1.Visible = True
    Else
        fraPeriode.Caption = "Periode - Tanggal BKM"
        cmdHapusData.Enabled = False
        cmdBayarOtomatis.Enabled = False
        txtCariNoCM.Visible = False
        Label1.Visible = False
    End If
End Sub

Private Sub txtCariNoCM_Change()
    mstrFilter = "AND NoCM LIKE '%" & txtCariNoCM.Text & "%'"
End Sub

Private Sub txtCariNoCM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCari.SetFocus
End Sub

