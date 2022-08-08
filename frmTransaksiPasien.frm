VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmTransaksiPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Transaksi Pelayanan Pasien"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTransaksiPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   11070
   Begin VB.Frame Frame1 
      Caption         =   "Transaksi Pelayanan Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   0
      TabIndex        =   20
      Top             =   2040
      Width           =   11055
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   8640
         TabIndex        =   31
         Top             =   6120
         Width           =   2055
      End
      Begin VB.TextBox txtGrandTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   29
         Top             =   6120
         Width           =   2415
      End
      Begin TabDlg.SSTab sstTP 
         Height          =   5655
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   9975
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Pe&layanan Tindakan"
         TabPicture(0)   =   "frmTransaksiPasien.frx":0CCA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdHapusData"
         Tab(0).Control(1)=   "cmdTambahPT"
         Tab(0).Control(2)=   "txtTindakanTotal"
         Tab(0).Control(3)=   "dgTindakan"
         Tab(0).Control(4)=   "Label1"
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Pemakaian &Obat && Alkes"
         TabPicture(1)   =   "frmTransaksiPasien.frx":0CE6
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "dgObatAlkes"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "txtAlkesTotal"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmdTambahPOA"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Pemakaian &Kamar"
         TabPicture(2)   =   "frmTransaksiPasien.frx":0D02
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "dgKamar"
         Tab(2).ControlCount=   1
         Begin VB.CommandButton cmdHapusData 
            Caption         =   "&Hapus Data"
            Height          =   375
            Left            =   -68550
            TabIndex        =   35
            Top             =   5040
            Width           =   2055
         End
         Begin VB.CommandButton cmdTambahPOA 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   8520
            TabIndex        =   33
            Top             =   5040
            Width           =   2055
         End
         Begin VB.CommandButton cmdTambahPT 
            Caption         =   "&Tambah Data"
            Height          =   375
            Left            =   -66480
            TabIndex        =   32
            Top             =   5040
            Width           =   2055
         End
         Begin VB.TextBox txtAlkesTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   3360
            TabIndex        =   27
            Top             =   5040
            Width           =   2415
         End
         Begin VB.TextBox txtTindakanTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   -72000
            TabIndex        =   25
            Top             =   5040
            Width           =   2415
         End
         Begin MSDataGridLib.DataGrid dgTindakan 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   22
            Top             =   720
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   7223
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgObatAlkes 
            Height          =   4095
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   7223
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dgKamar 
            Height          =   4095
            Left            =   -74760
            TabIndex        =   24
            Top             =   720
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   7223
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total Biaya Pemakaian Obat && Alkes"
            Height          =   210
            Left            =   240
            TabIndex        =   28
            Top             =   5100
            Width           =   2925
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Biaya Pelayanan Tindakan"
            Height          =   210
            Left            =   -74760
            TabIndex        =   26
            Top             =   5100
            Width           =   2550
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya Pelayanan"
         Height          =   210
         Left            =   360
         TabIndex        =   30
         Top             =   6195
         Width           =   1755
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Data Pasien"
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
      TabIndex        =   0
      Top             =   960
      Width           =   11055
      Begin VB.Frame Frame5 
         Caption         =   "Umur"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   580
         Left            =   5520
         TabIndex        =   7
         Top             =   350
         Width           =   2415
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   10
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   900
            MaxLength       =   6
            TabIndex        =   9
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   8
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Left            =   2130
            TabIndex        =   13
            Top             =   270
            Width           =   165
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Left            =   1350
            TabIndex        =   12
            Top             =   277
            Width           =   240
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Left            =   550
            TabIndex        =   11
            Top             =   277
            Width           =   285
         End
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2640
         TabIndex        =   4
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5040
         TabIndex        =   3
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtJenisPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   8040
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtTglDaftar 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   9360
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1560
         TabIndex        =   18
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   2640
         TabIndex        =   17
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "JK"
         Height          =   210
         Left            =   5160
         TabIndex        =   16
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pasien"
         Height          =   210
         Left            =   8040
         TabIndex        =   15
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pendaftaran"
         Height          =   210
         Left            =   9360
         TabIndex        =   14
         Top             =   360
         Width           =   1365
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   34
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
      Picture         =   "frmTransaksiPasien.frx":0D1E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9240
      Picture         =   "frmTransaksiPasien.frx":36DF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTransaksiPasien.frx":4467
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmTransaksiPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Store procedure untuk menghapus biaya pelayanan pasien
Private Sub sp_DelBiayaPelayanan(s_NoPendaftaran As String, s_KdRuangan As String, s_KdPelayanan As String, s_TglPelayanan As Date, s_KdPegawai As String)
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, s_NoPendaftaran)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, s_KdRuangan)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, s_KdPelayanan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(s_TglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, s_KdPegawai)

        .ActiveConnection = dbConn
        .CommandText = "Delete_BiayaPelayananNew"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penghapusan Biaya Pelayanan Pasien", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_BiayaPelayanan")
        End If
        Set dbcmd = Nothing
    End With
    Exit Sub
End Sub

Private Sub cmdHapusData_Click()
    On Error GoTo errHapusData

    If dgTindakan.ApproxCount = 0 Then Exit Sub
    If dgTindakan.Columns("Status Bayar").Value = "Sudah DiBayar" Then
        MsgBox "Pelayanan yang sudah dibayar tidak dapat dihapus", vbCritical, "Validasi"
        Exit Sub
    End If
    vbMsgboxRslt = MsgBox("Apakah anda yakin akan menghapus pelayanan '" _
    & dgTindakan.Columns("NamaPelayanan").Value & "'" & vbNewLine _
    & "Dengan tanggal pelayanan '" & dgTindakan.Columns("TglPelayanan").Value _
    & "'", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgboxRslt = vbNo Then Exit Sub

    strSQL = "SELECT KdRuangan From V_BiayaPelayananTindakan WHERE NoPendaftaran = '" & mstrNoPen & "' And KdPelayananRS ='" & dgTindakan.Columns("KdPelayananRS").Value & "'"
    Call msubRecFO(rs, strSQL)

    Call sp_DelBiayaPelayanan(mstrNoPen, rs(0), dgTindakan.Columns("KdPelayananRS").Value, dgTindakan.Columns("TglPelayanan").Value, strIDPegawaiAktif)
    Call subLoadPelayananDidapat
    MsgBox "Penghapusan data berhasil", vbInformation, "Informasi"

    Exit Sub
errHapusData:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdTambahPOA_Click()
    frmPemakaianObatAlkes.Show
End Sub

Private Sub cmdTambahPT_Click()
    frmTindakan.Show
    frmTindakan.txtNamaFormPengirim.Text = Me.Name
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgKamar_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKamar
    WheelHook.WheelHook dgKamar
End Sub

Private Sub dgObatAlkes_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgObatAlkes
    WheelHook.WheelHook dgObatAlkes
End Sub

Private Sub dgTindakan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgTindakan
    WheelHook.WheelHook dgTindakan
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errLoad

    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    Select Case KeyCode
        Case vbKey1
            If strCtrlKey Then sstTP.Tab = 0
        Case vbKey2
            If strCtrlKey Then sstTP.Tab = 1
        Case vbKey3
            If strCtrlKey Then sstTP.Tab = 2
    End Select

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subLoadPelayananDidapat
    Call subPemakaianObatAlkes
    sstTP.Tab = 0
    If mblnAdmin = True Then cmdHapusData.Visible = True Else cmdHapusData.Visible = False
End Sub

'Untuk meload pelayanan yang sudah pernah didapat
Public Sub subLoadPelayananDidapat()
    strSQL = "SELECT TglPelayanan,JenisPelayanan,NamaPelayanan,NamaRuangan AS [Ruang Pelayanan]," _
    & "Kelas,JenisTarif,CITO,JmlPelayanan as Jml,Total as Tarif,BiayaTotal," _
    & "DokterPemeriksa,[Status Bayar],KdPelayananRS FROM V_BiayaPelayananTindakan WHERE " _
    & "NoPendaftaran='" & mstrNoPen & "' ORDER BY TglPelayanan"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgTindakan.DataSource = rs
    With dgTindakan
        .Columns(0).Width = 1600
        .Columns(1).Width = 2000
        .Columns(2).Width = 2000
        .Columns(3).Width = 1600
        .Columns(4).Width = 900
        .Columns(5).Width = 1000
        .Columns(6).Width = 500
        .Columns(7).Width = 400
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 900
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 1000
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 2400
        .Columns(11).Width = 1200
        .Columns(12).Width = 0
    End With

    strSQL = "SELECT sum(BiayaTotal) as TotalBayar FROM V_BiayaPelayananTindakan " _
    & "WHERE NoPendaftaran='" _
    & mstrNoPen & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If IsNull(rs.Fields(0).Value) = False Then
        txtTindakanTotal.Text = FormatCurrency(rs.Fields(0).Value, 2)
    Else
        txtTindakanTotal.Text = FormatCurrency(0, 2)
    End If
    If txtAlkesTotal.Text = "" Then
        txtAlkesTotal.Text = 0
        txtAlkesTotal.Text = FormatCurrency(txtAlkesTotal.Text, 2)
    End If
    txtGrandTotal.Text = CCur(txtTindakanTotal.Text) + CCur(txtAlkesTotal.Text)
    txtGrandTotal.Text = FormatCurrency(txtGrandTotal.Text, 2)
End Sub

'Untuk meload pemakaian obat dan alkes yang sudah pernah didapat
Public Sub subPemakaianObatAlkes()
    strSQL = "SELECT TglPelayanan,[Detail Jenis Brg],NamaBarang," _
    & "NamaRuangan AS [Ruang Pelayanan],Kelas,JenisTarif,SatuanJml as Sat,JmlBarang as Jml," _
    & "HargaSatuan as Tarif,BiayaTotal,DokterPemeriksa,[Status Bayar] " _
    & "FROM V_BiayaPemakaianObatAlkes WHERE NoPendaftaran='" _
    & mstrNoPen & "' ORDER BY TglPelayanan"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgObatAlkes.DataSource = rs
    With dgObatAlkes
        .Columns(0).Width = 1600
        .Columns(1).Width = 2000
        .Columns(2).Width = 2000
        .Columns(3).Width = 1600
        .Columns(4).Width = 900
        .Columns(5).Width = 1000
        .Columns(6).Width = 400
        .Columns(7).Width = 400
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 900
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 1000
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 2400
        .Columns(11).Width = 1200
    End With

    strSQL = "SELECT sum(BiayaTotal) as TotalBayar FROM V_BiayaPemakaianObatAlkes " _
    & "WHERE NoPendaftaran='" _
    & mstrNoPen & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If IsNull(rs.Fields(0).Value) = False Then
        txtAlkesTotal.Text = FormatCurrency(rs.Fields(0).Value, 2)
    Else
        txtAlkesTotal.Text = FormatCurrency(0, 2)
    End If
    If txtTindakanTotal.Text = "" Then
        txtTindakanTotal.Text = 0
        txtTindakanTotal.Text = FormatCurrency(txtTindakanTotal.Text, 2)
    End If
    txtGrandTotal.Text = CCur(txtTindakanTotal.Text) + CCur(txtAlkesTotal)
    txtGrandTotal.Text = FormatCurrency(txtGrandTotal.Text, 2)
End Sub

