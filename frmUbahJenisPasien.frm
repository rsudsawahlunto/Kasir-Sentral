VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUbahJenisPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Asuransi Pasien"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUbahJenisPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   10485
   Begin VB.TextBox txtNamaFormPengirim2 
      Height          =   495
      Left            =   4080
      TabIndex        =   75
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtNoBKM 
      Height          =   375
      Left            =   4800
      TabIndex        =   71
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtKdInstalasi 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      TabIndex        =   69
      Text            =   "txtKdInstalasi"
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraDataRujukan 
      Caption         =   "Data Rujukan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   58
      Top             =   6360
      Width           =   10455
      Begin VB.TextBox txtNoRujukan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   480
         Width           =   3975
      End
      Begin MSDataListLib.DataCombo dcAsalRujukan 
         Height          =   330
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
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
      Begin MSComCtl2.DTPicker dtpTglDirujuk 
         Height          =   315
         Left            =   360
         TabIndex        =   28
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
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
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   154730499
         UpDown          =   -1  'True
         CurrentDate     =   37694
      End
      Begin MSDataListLib.DataCombo dcNamaPerujuk 
         Height          =   330
         Left            =   2520
         TabIndex        =   29
         Top             =   1080
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcNamaAsalRujukan 
         Height          =   330
         Left            =   6600
         TabIndex        =   27
         Top             =   480
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcDiagnosa 
         Height          =   330
         Left            =   6600
         TabIndex        =   30
         Top             =   1080
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama tempat Perujuk = Nama Puskesmas/ Nama Klinik/ Tempat Dokter Praktek/ Nama Rumah Sakit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   240
         TabIndex        =   65
         Top             =   1440
         Width           =   8565
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Perujuk (Dokter, Bidan, Mantri, dll)"
         Height          =   210
         Index           =   21
         Left            =   2520
         TabIndex        =   64
         Top             =   840
         Width           =   3345
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Rujukan"
         Height          =   210
         Index           =   24
         Left            =   2520
         TabIndex        =   63
         Top             =   240
         Width           =   930
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asal Rujukan"
         Height          =   210
         Index           =   25
         Left            =   240
         TabIndex        =   62
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Asal Rujukan (Nama Tempat Rujukan)"
         Height          =   210
         Index           =   27
         Left            =   6600
         TabIndex        =   61
         Top             =   240
         Width           =   3600
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Dirujuk"
         Height          =   210
         Index           =   26
         Left            =   240
         TabIndex        =   60
         Top             =   840
         Width           =   930
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnosa (Penyakit) Rujukan"
         Height          =   210
         Index           =   22
         Left            =   6600
         TabIndex        =   59
         Top             =   840
         Width           =   2325
      End
   End
   Begin VB.Frame fraPemakaianAsuransi 
      Caption         =   "Pemakaian Asuransi  (SP3 = Surat Jaminan Pelayanan)"
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
      TabIndex        =   51
      Top             =   4560
      Width           =   10455
      Begin VB.TextBox txtNoSJP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3840
         MaxLength       =   30
         TabIndex        =   19
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox txtNoKunjungan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         MaxLength       =   1
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtNoBP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         MaxLength       =   3
         TabIndex        =   21
         Text            =   "a24"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtAnakKe 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chkNoSJP 
         Caption         =   "No. SP3 Otomatis"
         Enabled         =   0   'False
         Height          =   210
         Left            =   3840
         TabIndex        =   18
         Top             =   360
         Width           =   3375
      End
      Begin MSDataListLib.DataCombo dcHubungan 
         Height          =   330
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
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
      Begin MSComCtl2.DTPicker dtpTglSJP 
         Height          =   315
         Left            =   8400
         TabIndex        =   20
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
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
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   154730499
         UpDown          =   -1  'True
         CurrentDate     =   37694
      End
      Begin MSDataListLib.DataCombo dcUnitKerja 
         Height          =   330
         Left            =   2040
         TabIndex        =   23
         Top             =   1200
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcKelasDitanggung 
         Height          =   330
         Left            =   7560
         TabIndex        =   24
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kelas Ditanggung"
         Height          =   210
         Index           =   13
         Left            =   7560
         TabIndex        =   72
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. SP3"
         Height          =   210
         Index           =   17
         Left            =   8400
         TabIndex        =   57
         Top             =   360
         Width           =   690
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hubungan Pasien"
         Height          =   210
         Index           =   15
         Left            =   240
         TabIndex        =   56
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anak Ke -"
         Height          =   210
         Index           =   16
         Left            =   2880
         TabIndex        =   55
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. BP"
         Height          =   210
         Index           =   18
         Left            =   240
         TabIndex        =   54
         Top             =   960
         Width           =   555
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kunj. Ke -"
         Height          =   210
         Index           =   20
         Left            =   960
         TabIndex        =   53
         Top             =   960
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bertugas di Unit / Bagian"
         Height          =   210
         Index           =   19
         Left            =   2040
         TabIndex        =   52
         Top             =   960
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   8880
      TabIndex        =   32
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   7200
      TabIndex        =   31
      Top             =   8280
      Width           =   1575
   End
   Begin VB.TextBox txtNamaFormPengirim 
      Height          =   495
      Left            =   0
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTglPendaftaran 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Frame fraDataKartuPeserta 
      Caption         =   "Data Kartu Peserta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   41
      Top             =   2160
      Width           =   10455
      Begin VB.CheckBox chkDiriSendiri 
         Caption         =   "Diri Sendiri"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   9120
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtAlamatPA 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   15
         Top             =   1800
         Width           =   8175
      End
      Begin VB.TextBox txtNipPA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7800
         MaxLength       =   16
         ScrollBars      =   1  'Horizontal
         TabIndex        =   13
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtNamaPA 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1200
         Width           =   3975
      End
      Begin VB.TextBox txtNoKartuPA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1200
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo dcPenjamin 
         Height          =   330
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   4215
         _ExtentX        =   7435
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
      Begin MSComCtl2.DTPicker dtpTglLahirPA 
         Height          =   315
         Left            =   6240
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   154206209
         UpDown          =   -1  'True
         CurrentDate     =   37694
      End
      Begin MSDataListLib.DataCombo dcPerusahaan 
         Height          =   330
         Left            =   4560
         TabIndex        =   9
         Top             =   600
         Width           =   4335
         _ExtentX        =   7646
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
      Begin MSDataListLib.DataCombo dcGolonganAsuransi 
         Height          =   330
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Golongan Asuransi"
         Height          =   210
         Index           =   7
         Left            =   240
         TabIndex        =   73
         Top             =   1560
         Width           =   1485
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Institusi Asal Pasien"
         Height          =   210
         Index           =   2
         Left            =   4560
         TabIndex        =   70
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Peserta"
         Height          =   210
         Index           =   14
         Left            =   2160
         TabIndex        =   47
         Top             =   1560
         Width           =   1230
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Lahir"
         Height          =   210
         Index           =   11
         Left            =   6240
         TabIndex        =   46
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. KTP / SIM Peserta"
         Height          =   210
         Index           =   12
         Left            =   7800
         TabIndex        =   45
         Top             =   960
         Width           =   1845
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Peserta"
         Height          =   210
         Index           =   10
         Left            =   2160
         TabIndex        =   44
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kartu Peserta"
         Height          =   210
         Index           =   9
         Left            =   240
         TabIndex        =   43
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Penjamin"
         Height          =   210
         Index           =   8
         Left            =   240
         TabIndex        =   42
         Top             =   360
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
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
      TabIndex        =   33
      Top             =   960
      Width           =   10455
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo dcJenisPasien 
         Height          =   315
         Left            =   8280
         TabIndex        =   6
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   12
         TabIndex        =   0
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4440
         MaxLength       =   9
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.Frame Frame4 
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
         Height          =   615
         Left            =   5640
         TabIndex        =   34
         Top             =   360
         Width           =   2535
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            MaxLength       =   6
            TabIndex        =   3
            Top             =   250
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            MaxLength       =   6
            TabIndex        =   4
            Top             =   250
            Width           =   375
         End
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   5
            Top             =   250
            Width           =   375
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "thn"
            Height          =   210
            Index           =   4
            Left            =   600
            TabIndex        =   37
            Top             =   302
            Width           =   285
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bln"
            Height          =   210
            Index           =   5
            Left            =   1440
            TabIndex        =   36
            Top             =   302
            Width           =   240
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "hr"
            Height          =   210
            Index           =   6
            Left            =   2280
            TabIndex        =   35
            Top             =   302
            Width           =   165
         End
      End
      Begin VB.Label lblNoPendaftaran 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   7680
         TabIndex        =   67
         Top             =   -120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblJenisPasien 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Pasien"
         Height          =   210
         Left            =   8280
         TabIndex        =   48
         Top             =   360
         Width           =   960
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. CM"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pasien"
         Height          =   210
         Index           =   1
         Left            =   1920
         TabIndex        =   39
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Index           =   3
         Left            =   4440
         TabIndex        =   38
         Top             =   360
         Width           =   1065
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   68
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
   Begin MSComctlLib.ProgressBar pbData 
      Height          =   360
      Left            =   120
      TabIndex        =   74
      Top             =   8400
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   635
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   200
      Scrolling       =   1
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8640
      Picture         =   "frmUbahJenisPasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmUbahJenisPasien.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmUbahJenisPasien.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmUbahJenisPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fTglLahir As Date
Dim fNoPendaftaran As String
Dim fNoSJP As String
Dim fNoBP As String
Dim fNoKunjungan As Integer
Dim fChkNoSJP As String
Dim fDcUnitKerja As String
Dim fNamaAsalRujukan As String
Dim fNamaPerujuk As String
Dim fDiagnosa As String
Dim fAlamatPA As String
Dim fIDPeserta As String
Dim fKdPerusahaan As String

Private Sub subLoadPemakaianAsuransi(s_NoPendaftaran As String, s_IdPenjamin As String)
    On Error GoTo errLoad

    strSQL = "SELECT * FROM v_PemakaianAsuransi WHERE NoPendaftaran = '" & s_NoPendaftaran & "' AND IdPenjamin='" & s_IdPenjamin & "' and StatusEnabled='1'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF = False Then
        dcHubungan.BoundText = IIf(IsNull(rs("KdHubungan")), "", rs("KdHubungan"))
        txtAnakKe.Text = IIf(IsNull(rs("AnakKe")), "", rs("AnakKe"))
        txtNoSJP.Text = IIf(IsNull(rs("NoSJP")), "", rs("NoSJP"))
        dtpTglSJP.Value = IIf(IsNull(rs("TglSJP")), Now, rs("TglSJP"))
        txtNoBP.Text = IIf(IsNull(rs("NoBP")), "", rs("NoBP"))
        txtNoKunjungan.Text = IIf(IsNull(rs("KunjunganKe")), "", rs("KunjunganKe"))
        dcUnitKerja.Text = IIf(IsNull(rs("UnitBagian")), "", rs("UnitBagian"))
    Else
        dcHubungan.BoundText = ""
        txtAnakKe.Text = ""
        txtNoSJP.Text = ""
        dtpTglSJP.Value = Now
        txtNoBP.Text = ""
        txtNoKunjungan.Text = ""
        dcUnitKerja.Text = ""
        dcKelasDitanggung.BoundText = ""
    End If

    Exit Sub
errLoad:
    Call msubPesanError("subLoadPemakaianAsuransi")
End Sub

Private Sub subTampungDataPenjamin()
    typAsuransi.strIdPenjamin = dcPenjamin.BoundText
    typAsuransi.strIdAsuransi = txtNoKartuPA.Text
    typAsuransi.strNoCm = txtNoCM.Text
    typAsuransi.strNamaPeserta = txtNamaPA.Text

    typAsuransi.strIdPeserta = IIf(txtNipPA.Text = "", "-", txtNipPA.Text)   ''allow null

    typAsuransi.strKdGolongan = dcGolonganAsuransi.BoundText
    typAsuransi.dTglLahir = dtpTglLahirPA.Value
    typAsuransi.strAlamat = IIf(txtAlamatPA.Text = "", "-", txtAlamatPA.Text)
    typAsuransi.strNoPendaftaran = IIf(txtNopendaftaran.Text <> "", txtNopendaftaran.Text, mstrNoPen)

    typAsuransi.strHubungan = dcHubungan.BoundText
    typAsuransi.strNoSJP = txtNoSJP.Text
    typAsuransi.dTglSJP = dtpTglSJP.Value
    typAsuransi.strNoBp = IIf(txtNoBP.Text = "", "-", txtNoBP.Text)
    typAsuransi.intNoKunjungan = IIf(val(txtNoKunjungan.Text) = 0, 1, val(txtNoKunjungan.Text))

    typAsuransi.strStatusNoSJP = IIf(chkNoSJP.Value = vbChecked, "O", "M")
    typAsuransi.intAnakKe = IIf(val(txtAnakKe.Text) = 0, 0, val(txtAnakKe.Text))
    typAsuransi.strUnitBagian = IIf(dcUnitKerja.Text = "", "-", Trim(dcUnitKerja.Text))

    typAsuransi.strNoRujukan = IIf(txtNoRujukan.Text = "", "-", txtNoRujukan.Text)
    typAsuransi.strKdRujukanAsal = dcAsalRujukan.BoundText
    typAsuransi.strDetailRujukanAsal = IIf(dcNamaAsalRujukan.Text = "", "-", dcNamaAsalRujukan.Text)
    typAsuransi.strKdDetailRujukanAsal = dcNamaAsalRujukan.BoundText
    typAsuransi.strNamaPerujuk = IIf(dcNamaPerujuk.Text = "", "-", dcNamaPerujuk.Text)

    typAsuransi.dTglDirujuk = dtpTglDirujuk.Value
    typAsuransi.strDiagnosaRujukan = IIf(dcDiagnosa.Text = "", "-", dcDiagnosa.Text)
    typAsuransi.strKdDiagnosa = dcDiagnosa.BoundText

    typAsuransi.strKdKelompokPasien = dcJenisPasien.BoundText
    typAsuransi.strPerusahaanPenjamin = dcPerusahaan.BoundText
    typAsuransi.strKdKelasDitanggung = dcKelasDitanggung.BoundText

    typAsuransi.blnSuksesAsuransi = True
    cmdSimpan.Enabled = False
End Sub

Private Function sp_JenisPasienJoinProgramAskes() As Boolean
    On Error GoTo StatusErr

    MousePointer = vbHourglass
    sp_JenisPasienJoinProgramAskes = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, dcPenjamin.BoundText)
        .Parameters.Append .CreateParameter("IdAsuransi", adVarChar, adParamInput, 25, txtNoKartuPA)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM)
        .Parameters.Append .CreateParameter("KdHubKeluarga", adChar, adParamInput, 2, dcHubungan.BoundText)
        .Parameters.Append .CreateParameter("NamaPeserta", adVarChar, adParamInput, 50, txtNamaPA.Text)
        '5
        .Parameters.Append .CreateParameter("IDPeserta", adVarChar, adParamInput, 16, txtNipPA)
        .Parameters.Append .CreateParameter("KdGolongan", adChar, adParamInput, 2, IIf(Len(Trim(dcGolonganAsuransi.Text)) = 0, Null, Trim(dcGolonganAsuransi.BoundText)))
        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(dtpTglLahirPA, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Alamat", adVarChar, adParamInput, 100, txtAlamatPA)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        '10
        .Parameters.Append .CreateParameter("KdHubungan", adChar, adParamInput, 2, dcHubungan.BoundText)
        .Parameters.Append .CreateParameter("NoSJP", adVarChar, adParamInput, 30, IIf(Len(Trim(txtNoSJP.Text)) = 0, Null, Trim(txtNoSJP.Text)))
        .Parameters.Append .CreateParameter("TglSJP", adDate, adParamInput, , Format(dtpTglSJP, "yyyy/MM/dd hh:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("NoBP", adChar, adParamInput, 3, IIf(Len(Trim(txtNoBP.Text)) = 0, Null, Trim(txtNoBP.Text)))
        '15
        .Parameters.Append .CreateParameter("KunjunganKe", adInteger, adParamInput, , IIf(val(txtNoKunjungan.Text) = 0, "1", txtNoKunjungan.Text))
        .Parameters.Append .CreateParameter("OutputNoSJP", adVarChar, adParamOutput, 30, Null)
        .Parameters.Append .CreateParameter("StatusNoSJP", adChar, adParamInput, 1, IIf(chkNoSJP.Value = vbChecked, "O", "M"))
        .Parameters.Append .CreateParameter("AnakKe", adInteger, adParamInput, , val(txtAnakKe.Text))
        .Parameters.Append .CreateParameter("UnitBagian", adVarChar, adParamInput, 50, IIf(Len(Trim(dcUnitKerja.Text)) = 0, Null, Trim(dcUnitKerja.Text)))
        '20
        .Parameters.Append .CreateParameter("KdPaket", adVarChar, adParamInput, 3, Null)
        .Parameters.Append .CreateParameter("NoRujukan", adVarChar, adParamInput, 30, txtNoRujukan.Text)
        .Parameters.Append .CreateParameter("KdRujukanAsal", adChar, adParamInput, 2, dcAsalRujukan.BoundText)
        .Parameters.Append .CreateParameter("DetailRujukanAsal", adVarChar, adParamInput, 100, IIf(Len(Trim(dcNamaAsalRujukan.Text)) = 0, Null, dcNamaAsalRujukan.Text))
        .Parameters.Append .CreateParameter("KdDetailRujukanAsal", adChar, adParamInput, 8, IIf(chkNoSJP.Value = vbChecked, "12345678", dcNamaAsalRujukan.BoundText))
        '25
        .Parameters.Append .CreateParameter("NamaPerujuk", adVarChar, adParamInput, 50, IIf(Len(Trim(dcNamaPerujuk.Text)) = 0, Null, Trim(dcNamaPerujuk.Text)))
        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , Format(dtpTglDirujuk.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("DiagnosaRujukan", adVarChar, adParamInput, 100, IIf(Len(Trim(dcDiagnosa.Text)) = 0, Null, Trim(dcDiagnosa.Text)))
        .Parameters.Append .CreateParameter("KdDiagnosa", adVarChar, adParamInput, 7, dcDiagnosa.BoundText)
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, dcJenisPasien.BoundText)
        .Parameters.Append .CreateParameter("KdInstitusiAsal", adVarChar, adParamInput, 4, IIf(dcPerusahaan.Text = "", Null, dcPerusahaan.BoundText))
        .Parameters.Append .CreateParameter("KdKelasDiTanggung", adChar, adParamInput, 2, dcKelasDitanggung.BoundText)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_JenisPasienJoinProgramAskesNew"
'        .CommandText = "dbo.AU_AsuransiPasienJoinProgramAskes"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_JenisPasienJoinProgramAskes = False
        Else
            txtNoSJP.Text = IIf(IsNull(.Parameters("OutputNoSJP")), "", .Parameters("OutputNoSJP"))
            cmdSimpan.Enabled = False
'            Call Add_HistoryLoginActivity("Update_JenisPasienJoinProgramAskesNew")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    MousePointer = vbDefault

    Exit Function
StatusErr:
    cmdSimpan.Enabled = True
    MousePointer = vbDefault
    sp_JenisPasienJoinProgramAskes = False
    Call msubPesanError("sp_JenisPasienJoinProgramAskes")
    MsgBox "Ulangi proses simpan ", vbCritical, "Validasi"
End Function

'Store procedure untuk mengisi asuransi pasien
Private Sub sp_AsuransiPasien(ByVal adoCommand As ADODB.Command)
Dim xrtSQL As String
    Set dbcmd = New ADODB.Command
    
    MousePointer = vbHourglass
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, typAsuransi.strIdPenjamin)
        .Parameters.Append .CreateParameter("IdAsuransi", adVarChar, adParamInput, 25, typAsuransi.strIdAsuransi)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, mstrNoCM)
        .Parameters.Append .CreateParameter("KdHubKeluarga", adChar, adParamInput, 2, typAsuransi.strHubungan)
        .Parameters.Append .CreateParameter("NamaPeserta", adVarChar, adParamInput, 50, typAsuransi.strNamaPeserta)
        
        .Parameters.Append .CreateParameter("IDPeserta", adVarChar, adParamInput, 16, typAsuransi.strIdPeserta)
        .Parameters.Append .CreateParameter("KdGolongan", adChar, adParamInput, 2, IIf(Len(Trim(typAsuransi.strKdGolongan)) = 0, Null, Trim(typAsuransi.strKdGolongan)))
        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(typAsuransi.dTglLahir, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("Alamat", adVarChar, adParamInput, 100, typAsuransi.strAlamat)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        
        .Parameters.Append .CreateParameter("KdHubungan", adChar, adParamInput, 2, typAsuransi.strHubungan)
        If typAsuransi.strNoSJP <> "" Then
            .Parameters.Append .CreateParameter("NoSJP", adVarChar, adParamInput, 30, typAsuransi.strNoSJP)
        Else
            .Parameters.Append .CreateParameter("NoSJP", adVarChar, adParamInput, 30, Null)
        End If
        .Parameters.Append .CreateParameter("TglSJP", adDate, adParamInput, , Format(typAsuransi.dTglSJP, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("NoBP", adVarChar, adParamInput, 3, IIf(Len(Trim(typAsuransi.strNoBp)) = 0, Null, Trim(typAsuransi.strNoBp)))
        
        .Parameters.Append .CreateParameter("KunjunganKe", adInteger, adParamInput, , typAsuransi.intNoKunjungan)
        .Parameters.Append .CreateParameter("OutputNoSJP", adVarChar, adParamOutput, 30, Null)
        .Parameters.Append .CreateParameter("StatusNoSJP", adChar, adParamInput, 1, typAsuransi.strStatusNoSJP)
        .Parameters.Append .CreateParameter("AnakKe", adInteger, adParamInput, , typAsuransi.intAnakKe)
        .Parameters.Append .CreateParameter("UnitBagian", adVarChar, adParamInput, 50, IIf(Len(Trim(typAsuransi.strUnitBagian)) = 0, Null, Trim(typAsuransi.strUnitBagian)))
        
        .Parameters.Append .CreateParameter("KdPaket", adVarChar, adParamInput, 3, Null)
        .Parameters.Append .CreateParameter("NoRujukan", adVarChar, adParamInput, 30, typAsuransi.strNoRujukan)
        .Parameters.Append .CreateParameter("KdRujukanAsal", adChar, adParamInput, 2, typAsuransi.strKdRujukanAsal)
        .Parameters.Append .CreateParameter("DetailRujukanAsal", adVarChar, adParamInput, 100, typAsuransi.strDetailRujukanAsal)
        .Parameters.Append .CreateParameter("KdDetailRujukanAsal", adChar, adParamInput, 8, typAsuransi.strKdDetailRujukanAsal)
        
        .Parameters.Append .CreateParameter("NamaPerujuk", adVarChar, adParamInput, 50, typAsuransi.strNamaPerujuk)
        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , typAsuransi.dTglDirujuk)
        .Parameters.Append .CreateParameter("DiagnosaRujukan", adVarChar, adParamInput, 100, typAsuransi.strDiagnosaRujukan)
        .Parameters.Append .CreateParameter("KdDiagnosa", adVarChar, adParamInput, 7, typAsuransi.strKdDiagnosa)
        
        '###24-4-2008 by john ----'edit splakuk
        xrtSQL = "SELECT  KdinstitusiAsal, InstitusiAsal FROM InstitusiAsalPasien WHERE InstitusiAsal LIKE '" & typAsuransi.strPerusahaanPenjamin & "%' or KdInstitusiAsal LIKE '" & typAsuransi.strPerusahaanPenjamin & "' and StatusEnabled='1'"
        Call msubRecFO(rsx, xrtSQL)
        .Parameters.Append .CreateParameter("KdInstitusiAsal", adVarChar, adParamInput, 4, IIf(Len(Trim(rsx(0).Value)) = 0, Null, Trim(rsx(0).Value)))
        .Parameters.Append .CreateParameter("KdKelasDiTanggung", adChar, adParamInput, 2, typAsuransi.strKdKelasDitanggung)
        
        .ActiveConnection = dbConn
        .CommandText = "AU_AsuransiPasienJoinProgramAskes"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan Asuransi Pasien", vbCritical, "Validasi"
            mstrNoSJP = typAsuransi.strNoSJP
        Else
            mstrNoSJP = typAsuransi.strNoSJP
'            Call Add_HistoryLoginActivity("AU_AsuransiPasienJoinProgramAskes")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub subKosong()
    txtNoCM.Text = ""
    txtNamaPasien.Text = ""
    txtJK.Text = ""
    txtThn.Text = ""
    txtBln.Text = ""
    txtHr.Text = ""
    txtNopendaftaran.Text = ""

    chkDiriSendiri.Value = vbUnchecked

    dcPenjamin.BoundText = ""
    txtNoKartuPA.Text = ""
    txtNamaPA.Text = ""
    dtpTglLahirPA.Value = Now
    txtNipPA.Text = ""
    dcKelasDitanggung.BoundText = ""
    txtAlamatPA.Text = ""

    dcHubungan.BoundText = ""
    txtAnakKe.Text = ""
    chkNoSJP.Value = vbUnchecked
    dtpTglSJP.Value = Now
    txtNoBP.Text = ""
    txtNoKunjungan.Text = ""
    dcUnitKerja.BoundText = ""

    dcAsalRujukan.BoundText = ""
    txtNoRujukan.Text = ""
    dcNamaAsalRujukan.BoundText = ""
    dtpTglDirujuk.Value = Now
    dcNamaPerujuk.BoundText = ""
    dcDiagnosa.BoundText = ""
    dcGolonganAsuransi.BoundText = ""
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    Call msubDcSource(dcJenisPasien, rs, "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where StatusEnabled='1' order by JenisPasien")
    Call msubDcSource(dcHubungan, rs, "SELECT KdHubungan, NamaHubungan FROM HubunganPesertaAsuransi where StatusEnabled='1'")
    Call msubDcSource(dcGolonganAsuransi, rs, "SELECT     KdGolongan, NamaGolongan FROM GolonganAsuransi where StatusEnabled='1'")
    Call msubDcSource(dcUnitKerja, rs, "SELECT KdRuangan, NamaRuangan FROM Ruangan where StatusEnabled='1' ORDER BY NamaRuangan")
    Call msubDcSource(dcAsalRujukan, rs, "SELECT KdRujukanAsal, RujukanAsal FROM RujukanAsal where StatusEnabled='1'")
    strSQL = "SELECT KdDetailRujukanAsal, DetailRujukanAsal" & _
    " FROM DetailRujukanAsal " & _
    " WHERE (KdRujukanAsal = '" & dcAsalRujukan.BoundText & "')"
    Call msubDcSource(dcNamaAsalRujukan, rs, strSQL)
    Call msubDcSource(dcNamaPerujuk, rs, "SELECT KodeDokter, NamaDokter FROM V_DaftarDokter")
    Call msubDcSource(dcDiagnosa, rs, "SELECT KdDiagnosa, NamaDiagnosa FROM Diagnosa where StatusEnabled='1' ORDER BY NamaDiagnosa")

    strSQL = "SELECT  KdInstitusiAsal, InstitusiAsal FROM InstitusiAsalPasien where StatusEnabled='1' order by InstitusiAsal"
    Call msubDcSource(dcPerusahaan, rs, strSQL)

    Exit Sub
errLoad:
    Set rs = Nothing
    Call msubPesanError
End Sub

Private Sub chkDiriSendiri_Click()
    On Error GoTo errLoad
    If chkDiriSendiri.Value = 1 Then
        strSQL = "SELECT NamaLengkap, NoIdentitas, Alamat,TglLahir FROM v_S_RegistrasiDataPasien WHERE NocM='" & txtNoCM.Text & "'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount <> 0 Then
            txtNamaPA.Text = rs("NamaLengkap")
            txtNipPA.Text = rs("NoIdentitas") & ""
            txtAlamatPA.Text = rs("Alamat") & ""
            dtpTglLahirPA.Value = Format(rs("TglLahir"), "dd/mm/yyyy")
            dcHubungan.Text = "Peserta"
        Else
            txtNamaPA.Text = ""
            txtNipPA.Text = ""
            txtAlamatPA.Text = ""
            dtpTglLahirPA.Value = Now
            dcHubungan.Text = ""
        End If
    Else
        txtNamaPA.Text = ""
        txtNipPA.Text = ""
        txtAlamatPA.Text = ""
        dcHubungan.Text = ""
        dtpTglLahirPA.Value = Now
    End If
    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub chkDiriSendiri_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNoKartuPA.SetFocus
End Sub

Private Sub chkNoSJP_Click()
    If chkNoSJP.Value = vbChecked Then txtNoSJP.Enabled = False Else txtNoSJP.Enabled = True
End Sub

Private Sub chkNoSJP_KeyPress(KeyAscii As Integer)
    If chkDiriSendiri.Value = vbChecked Then dtpTglSJP.SetFocus Else txtNoSJP.SetFocus
End Sub


Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

'    Call msubRecFO(dbRst, "SELECT IdPenjamin FROM PenjaminKelompokPasien WHERE KdKelompokPasien = '" & dcJenisPasien.BoundText & "'")
'    If dbRst(0).Value = "2222222222" Then
'
'        If sp_UpdateJenisPasienUmum(dcJenisPasien.BoundText, txtNoPendaftaran.Text) = False Then Exit Sub
'
'        MousePointer = vbHourglass
'
'        MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
'        cmdSimpan.Enabled = False
'        MousePointer = vbDefault
'        Exit Sub
'    End If

    If Periksa("datacombo", dcPenjamin, "Penjamin belum di isi") = False Then Exit Sub
    If Periksa("text", txtNoKartuPA, "Nomor kartu belum di isi") = False Then Exit Sub
    If Periksa("text", txtNamaPA, "Nama peserta asuransi ?") = False Then Exit Sub
    If Periksa("datacombo", dcKelasDitanggung, "Kelas ditanggung ?") = False Then Exit Sub
    If Periksa("datacombo", dcHubungan, "Hubungan peserta asuransi belum di isi") = False Then Exit Sub
    If Periksa("text", txtNoSJP, "No SP3 harus diisi") = False Then Exit Sub
    If Periksa("datacombo", dcAsalRujukan, "Asal rujukan belum di isi") = False Then Exit Sub
    If Periksa("datacombo", dcGolonganAsuransi, "Golongan asuransi kosong") = False Then Exit Sub
    If Periksa("text", txtNoRujukan, "No rujukan belum di isi") = False Then Exit Sub
    mstrNoSJP = txtNoSJP.Text
'    If txtNamaFormPengirim.Text = "tampung" Then
'        Call subTampungDataPenjamin
''        Call sp_AsuransiPasien(dbcmd)
'    ElseIf txtNamaFormPengirim2.Text = "frmBayar" Then
        strSQLx = "Select * from PemakaianAsuransiDetail where NoPendaftaran='" & txtNopendaftaran.Text & "' and KdKelompokPasien='" & dcJenisPasien.BoundText & "' and IdPenjamin='" & dcPenjamin.BoundText & "' and NoSJP='" & txtNoSJP.Text & "'"
        Call msubRecFO(rsx, strSQLx)
        If rsx.EOF = True Then
'            strSQL = "Insert into AsuransiPasien values('" & dcPenjamin.BoundText & "' ,'" & txtNoKartuPA.Text & "','" & txtNoCM.Text & "' ,'" & txtNamaPA.Text & "','" & txtNipPA.Text & "','" & dcGolonganAsuransi.BoundText & "','" & Format(dtpTglLahirPA.Value, "yyyy/MM/dd HH:mm:ss") & "','" & txtAlamatPA.Text & "','" & dcPerusahaan.BoundText & "')"
'            Call msubRecFO(rs, strSQL)
'            strSQL = "Insert into PemakaianAsuransi values('" & dcPenjamin.BoundText & "' ,'" & txtNoKartuPA.Text & "','" & txtNoCM.Text & "' ,'" & txtNoPendaftaran.Text & "','" & dcHubungan.BoundText & "','" & txtNoSJP.Text & "','" & Format(dtpTglSJP.Value, "yyyy/MM/dd HH:mm:ss") & "','" & typAsuransi.strNoBp & "'," & typAsuransi.intNoKunjungan & ",'" & typAsuransi.strUnitBagian & "'," & typAsuransi.intAnakKe & ",'" & dcKelasDitanggung.BoundText & "')"
'            Call msubRecFO(rs, strSQL)
            strSQL = "Insert into PemakaianAsuransiDetail values('" & txtNopendaftaran.Text & "' ,'" & dcJenisPasien.BoundText & "','" & dcPenjamin.BoundText & "' ,'" & txtNoSJP.Text & "','" & txtNoKartuPA.Text & "')"
            Call msubRecFO(rs, strSQL)
'            strSQL = "Update PasienDaftar set KdKelompokPasien='" & dcJenisPasien.BoundText & "', IdPenjamin='" & dcPenjamin.BoundText & "' where NoPendaftaran='" & txtNoPendaftaran.Text & "'"
'            Call msubRecFO(rs, strSQL)
            
        Else
'            strSQL = "Update PemakaianAsuransi set IdAsuransi='" & txtNoKartuPA.Text & "', KdHubungan='" & dcHubungan.BoundText & "',TglSJP='" & Format(dtpTglSJP.Value, "yyyy/MM/dd HH:mm:ss") & "',NoBP='" & typAsuransi.strNoBp & "',KunjunganKe=" & typAsuransi.intNoKunjungan & ",UnitBagian='" & typAsuransi.strUnitBagian & "',AnakKe=" & typAsuransi.intAnakKe & ", KdKelasDitanggung='" & dcKelasDitanggung.BoundText & "' where NoPendaftaran='" & txtNoPendaftaran.Text & "' and NoSJP='" & txtNoSJP.Text & "'"
'            Call msubRecFO(rs, strSQL)
            strSQL = "Update PemakaianAsuransiDetail set IdAsuransi='" & txtNoKartuPA.Text & "' where NoPendaftaran='" & txtNopendaftaran.Text & "' and NoSJP='" & txtNoSJP.Text & "' and KdKelompokPasien='" & dcJenisPasien.BoundText & "' and IdPenjamin='" & dcPenjamin.BoundText & "'"
            Call msubRecFO(rs, strSQL)
        End If
        
'    Else
'        If sp_JenisPasienJoinProgramAskes = False Then Exit Sub
        
        cmdSimpan.Enabled = False
'    End If
    MousePointer = vbDefault
    MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
    Exit Sub
errLoad:
    Call msubPesanError("cmdSimpan_Click")
    MousePointer = vbDefault
End Sub

Private Sub cmdTutup_Click()

'    If txtNamaFormPengirim2.Text = "frmBayar" Then
'
'            Call PostingHutangPenjaminPasien_AU("A")
'            Call frmTagihanPasien.txtNoPendaftaran_KeyPress(13)
'            frmTagihanPasien.cmdPerbaikiData_Click
''            frmBayar.txtTAsuransi.Text = FormatCurrency(frmTagihanPasien.txtTAsuransi.Text, 2)
'
'
'            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 1) = dcPenjamin.Text
'            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 7) = dcPenjamin.BoundText
'            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 2) = Format(mcurTM_HrsDibyr_M + mcurOA_HrsDibyr_M, "#,###.00")
'            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 3) = Format(mcurAll_TP_M, "#,###.00")
'            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 4) = 0
'            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 5) = Format(CCur(frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 2) - frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 3)), "#,###.00")
'            frmBayar.dcPenjamin.Visible = False
'            frmBayar.fgMulti.Col = 3
'            frmBayar.fgMulti.SetFocus
'
'
'
'            Call frmBayar.subHitungTotal(frmBayar.fgMulti.Rows - 1, frmBayar.fgMulti.Cols)
'
'
'    End If

    If mblnTemp = True Then
        Unload Me
        mblnTemp = False
        
    Else
        Unload Me
    End If
'    If txtNamaFormPengirim.Text = "frmBayar" Then
'
'            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 1) = dcPenjamin.Text
'            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 7) = dcPenjamin.BoundText
''            fgMulti.TextMatrix(fgMulti.Row, 2) = Format(CCur(txtSisaTagihanMulti.Text), "#,###.00")
''            fgMulti.TextMatrix(fgMulti.Row, 3) = 0
''            fgMulti.TextMatrix(fgMulti.Row, 4) = 0
'            frmBayar.dcPenjamin.Visible = False
'            frmBayar.fgMulti.Col = 3
'            frmBayar.fgMulti.SetFocus
    
    'End If
End Sub

Public Sub PostingHutangPenjaminPasien_AU(strStatus As String)
    On Error GoTo hell_
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, strStatus)

        .ActiveConnection = dbConn
        .CommandText = "dbo.PostingHutangPenjaminPasien_AU"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam proses update HP pasien", vbCritical, "Validasi"
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing

    Exit Sub
hell_:
    msubPesanError
End Sub

Private Sub dcAsalRujukan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcAsalRujukan.MatchedWithList = True Then txtNoRujukan.SetFocus
        strSQL = "SELECT KdRujukanAsal, RujukanAsal FROM RujukanAsal where StatusEnabled='1' and (RujukanAsal LIKE '%" & dcAsalRujukan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcAsalRujukan.Text = ""
            Exit Sub
        End If
        dcAsalRujukan.BoundText = rs(0).Value
        dcAsalRujukan.Text = rs(1).Value
    End If
End Sub

Private Sub dcDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcDiagnosa.MatchedWithList = True Then cmdSimpan.SetFocus
        strSQL = "SELECT KdDiagnosa, NamaDiagnosa FROM Diagnosa where StatusEnabled='1'  and (NamaDiagnosa LIKE '%" & dcDiagnosa.Text & "%')ORDER BY NamaDiagnosa"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcDiagnosa.Text = ""
            Exit Sub
        End If
        dcDiagnosa.BoundText = rs(0).Value
        dcDiagnosa.Text = rs(1).Value
    End If
End Sub

Private Sub dcGolonganAsuransi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcGolonganAsuransi.MatchedWithList = True Then txtAlamatPA.SetFocus
        strSQL = "SELECT     KdGolongan, NamaGolongan FROM GolonganAsuransi where StatusEnabled='1' and (NamaGolongan LIKE '%" & dcGolonganAsuransi.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcGolonganAsuransi.Text = ""
            Exit Sub
        End If
        dcGolonganAsuransi.BoundText = rs(0).Value
        dcGolonganAsuransi.Text = rs(1).Value
    End If
End Sub

Private Sub dcHubungan_Change()
    txtAnakKe.Text = ""
    If dcHubungan.BoundText = "04" Then txtAnakKe.Enabled = True Else txtAnakKe.Enabled = False
End Sub

Private Sub dcJenisPasien_Change()
    On Error GoTo errLoad
    Set rs = Nothing
    rs.Open "select * from v_Penjaminpasien where KdKelompokPasien='" & dcJenisPasien.BoundText & "' and StatusEnabled='1' ORDER BY NamaPenjamin", dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcPenjamin.RowSource = rs
    dcPenjamin.BoundColumn = rs.Fields("idpenjamin").Name
    dcPenjamin.ListField = rs.Fields("namapenjamin").Name
    dcPenjamin.BoundText = ""

    Call msubRecFO(dbRst, "SELECT IdPenjamin FROM PenjaminKelompokPasien WHERE KdKelompokPasien = '" & dcJenisPasien.BoundText & "'")
    If dbRst(0).Value = "2222222222" Then
        fraDataKartuPeserta.Enabled = False
        fraPemakaianAsuransi.Enabled = False
        fraDataRujukan.Enabled = False
        dcPerusahaan.Text = ""
    Else
        fraDataKartuPeserta.Enabled = True
        fraPemakaianAsuransi.Enabled = True
        fraDataRujukan.Enabled = True
    End If
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub dcKelasDitanggung_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

'    tempKode = dcKelasDitanggung.BoundText
    strSQL = "SELECT DISTINCT KdKelas, DeskKelas FROM V_KelasDitanggungPenjamin WHERE (IdPenjamin = '" & dcPenjamin.BoundText & "') AND KdKelompokPasien = '" & dcJenisPasien.BoundText & "'"
    Call msubDcSource(dcKelasDitanggung, rs, strSQL)
'    dcKelasDitanggung.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKelasDitanggung_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKelasDitanggung.MatchedWithList = True Then dcAsalRujukan.SetFocus
        strSQL = "SELECT KdKelas, DeskKelas FROM KelasPelayanan where KdKelas <>'04' and (DeskKelas LIKE '%" & dcKelasDitanggung.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcKelasDitanggung.Text = ""
            Exit Sub
        End If
'        dcKelasDitanggung.BoundText = rs(0).Value
'        dcKelasDitanggung.Text = rs(1).Value
    End If
End Sub

Private Sub dcHubungan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcHubungan.MatchedWithList = True Then If txtAnakKe.Enabled = True Then txtAnakKe.SetFocus Else txtNoSJP.SetFocus
        strSQL = "SELECT KdHubungan, NamaHubungan FROM HubunganPesertaAsuransi where StatusEnabled='1' and (NamaHubungan LIKE '%" & dcHubungan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcHubungan.Text = ""
            Exit Sub
        End If
        dcHubungan.BoundText = rs(0).Value
        dcHubungan.Text = rs(1).Value
    End If
End Sub

Private Sub dcJenisPasien_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then If fraDataKartuPeserta.Enabled = True Then dcPenjamin.SetFocus Else cmdSimpan.SetFocus
End Sub

Private Sub dcNamaAsalRujukan_GotFocus()
    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcNamaAsalRujukan.BoundText
    strSQL = "SELECT DetailRujukanAsal.KdDetailRujukanAsal, DetailRujukanAsal.DetailRujukanAsal" & _
    " FROM DetailRujukanAsal " & _
    " WHERE (KdRujukanAsal = '" & dcAsalRujukan.BoundText & "') and StatusEnabled='1'"
    Set rs = Nothing
    Call msubDcSource(dcNamaAsalRujukan, rs, strSQL)
    dcNamaAsalRujukan.BoundText = tempKode

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcNamaAsalRujukan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcNamaAsalRujukan.MatchedWithList = True Then dtpTglDirujuk.SetFocus
        strSQL = "SELECT KdDetailRujukanAsal, DetailRujukanAsal" & _
        " FROM DetailRujukanAsal " & _
        " WHERE (KdRujukanAsal = '" & dcAsalRujukan.BoundText & "') and (DetailRujukanAsal LIKE '%" & dcNamaAsalRujukan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcNamaAsalRujukan.Text = ""
            Exit Sub
        End If
        dcNamaAsalRujukan.BoundText = rs(0).Value
        dcNamaAsalRujukan.Text = rs(1).Value
    End If
End Sub

Private Sub dcNamaPerujuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcNamaPerujuk.MatchedWithList = True Then dcDiagnosa.SetFocus
        strSQL = "SELECT KodeDokter, NamaDokter FROM V_DaftarDokter where (NamaDokter LIKE '%" & dcNamaPerujuk.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcNamaPerujuk.Text = ""
            Exit Sub
        End If
        dcNamaPerujuk.BoundText = rs(0).Value
        dcNamaPerujuk.Text = rs(1).Value
    End If
End Sub

Private Sub dcPenjaminx()
    On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "SELECT dbo.AsuransiPasien.IdPenjamin, dbo.AsuransiPasien.IdAsuransi, dbo.AsuransiPasien.NoCM, dbo.AsuransiPasien.NamaPeserta, " & _
    " dbo.AsuransiPasien.IDPeserta, dbo.AsuransiPasien.KdGolongan, dbo.AsuransiPasien.TglLahir, dbo.AsuransiPasien.Alamat," & _
    " dbo.AsuransiPasien.KdInstitusiAsal, dbo.InstitusiAsalPasien.InstitusiAsal AS NamaPerusahaan, dbo.InstitusiAsalPasien.StatusEnabled" & _
    " FROM dbo.AsuransiPasien LEFT OUTER JOIN" & _
    " dbo.InstitusiAsalPasien ON dbo.AsuransiPasien.KdInstitusiAsal = dbo.InstitusiAsalPasien.KdInstitusiAsal INNER JOIN" & _
    " dbo.Penjamin ON dbo.AsuransiPasien.IdPenjamin = dbo.Penjamin.IdPenjamin " & _
    " WHERE (AsuransiPasien.NoCM = '" & txtNoCM.Text & "') AND (AsuransiPasien.IdPenjamin = '" & dcPenjamin.BoundText & "') and (dbo.InstitusiAsalPasien.StatusEnabled='1')"
    Call msubRecFO(rs, strSQL)

    
    If rs.EOF = False Then
    If chkDiriSendiri.Value = Unchecked Then
        
        txtNoKartuPA.Text = IIf(IsNull(rs("IdAsuransi")), "", rs("IdAsuransi"))
        txtNamaPA.Text = IIf(IsNull(rs("NamaPeserta")), "", rs("NamaPeserta"))
        txtNipPA.Text = IIf(IsNull(rs("IDPeserta")), "-", rs("IDPeserta"))
        dcGolonganAsuransi.BoundText = IIf(IsNull(rs("KdGolongan")), "", rs("KdGolongan"))
        dtpTglLahirPA.Value = IIf(IsNull(rs("TglLahir")), Now, rs("TglLahir"))
        txtAlamatPA.Text = IIf(IsNull(rs("Alamat")), "", rs("Alamat"))
        dcPerusahaan.Text = IIf(IsNull(rs("NamaPerusahaan")), "", rs("NamaPerusahaan"))
        Call subLoadPemakaianAsuransi(txtNopendaftaran.Text, dcPenjamin.BoundText)
        dcHubungan.SetFocus
     
'    Else
'        txtNoKartuPA.Text = ""
'        txtNamaPA.Text = ""
'        txtNipPA.Text = ""
'        dcGolonganAsuransi.BoundText = ""
'        dtpTglLahirPA.value = Now
'        txtAlamatPA.Text = ""
    End If
    
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPenjamin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcPenjamin.MatchedWithList = True Then dcPerusahaan.SetFocus
        strSQL = "select * from v_Penjaminpasien where KdKelompokPasien='" & dcJenisPasien.BoundText & "' and StatusEnabled='1'  and (NamaPenjamin LIKE '%" & dcPenjamin.Text & "%')ORDER BY NamaPenjamin"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcPenjamin.Text = ""
            Exit Sub
        End If
        dcPenjamin.BoundText = rs(0).Value
        dcPenjamin.Text = rs(1).Value
        Call dcPenjaminx
    End If
End Sub

Private Sub dcPerusahaan_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    If KeyAscii = 13 Then
        If dcPerusahaan.MatchedWithList = True Then chkDiriSendiri.SetFocus
        strSQL = "SELECT  KdInstitusiAsal, InstitusiAsal FROM InstitusiAsalPasien WHERE (InstitusiAsal LIKE '" & dcPerusahaan.Text & "%') and StatusEnabled='1'"
        Set rs = Nothing
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcPerusahaan.Text = ""
            Exit Sub
        End If
        dcPerusahaan.BoundText = rs(0).Value
        dcPerusahaan.Text = rs(1).Value
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcUnitKerja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcUnitKerja.MatchedWithList = True Then dcKelasDitanggung.SetFocus
        strSQL = "SELECT KdRuangan, NamaRuangan FROM Ruangan where StatusEnabled='1'  and (NamaRuangan LIKE '%" & dcUnitKerja.Text & "%')ORDER BY NamaRuangan"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
            dcUnitKerja.Text = ""
            Exit Sub
        End If
        dcUnitKerja.BoundText = rs(0).Value
        dcUnitKerja.Text = rs(1).Value
    End If
End Sub

Private Sub dtpTglDirujuk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcNamaPerujuk.SetFocus
End Sub

Private Sub dtpTglLahirPA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtNipPA.SetFocus
End Sub

Private Sub dtpTglSJP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtNoBP.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpTglLahirPA.Value = Now
    dtpTglSJP.Value = Now
    dtpTglDirujuk.Value = Now

    txtNoBP.Text = ""
    txtNoKunjungan.Text = ""

    If mblnFormDaftarAntrian = True Then txtNoCM.Text = mstrNoCM
    txtNopendaftaran = mstrNoPen
    If mblnAdmin = False Then
        dcJenisPasien.Enabled = False
    Else
        dcJenisPasien.Enabled = True
    End If

    Call subLoadDcSource

    dcJenisPasien.Text = "ASKES PNS"

    Set rs = Nothing
    rs.Open "select * from v_Penjaminpasien where KdKelompokPasien='" & dcJenisPasien.BoundText & "' and StatusEnabled='1' ORDER BY NamaPenjamin", dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcPenjamin.RowSource = rs
    dcPenjamin.BoundColumn = rs.Fields("idpenjamin").Name
    dcPenjamin.ListField = rs.Fields("namapenjamin").Name
    dcPenjamin.BoundText = ""

    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
If cmdSimpan.Enabled = False Then
    If txtNamaFormPengirim2.Text = "frmBayar" Then
    
'            Call PostingHutangPenjaminPasien_AU("A")
'            Call frmTagihanPasien.txtNoPendaftaran_KeyPress(13)
'            frmTagihanPasien.cmdPerbaikiData_Click
'            frmBayar.txtTAsuransi.Text = FormatCurrency(frmTagihanPasien.txtTAsuransi.Text, 2)
    
        
            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 1) = dcPenjamin.Text
            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 7) = dcPenjamin.BoundText
            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 2) = Format(CCur(frmBayar.lblTotalTagihan.Caption), "#,###.00")
            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 3) = Format(CCur(frmBayar.lblTotalTagihan.Caption), "#,###.00")
            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 4) = 0
            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 5) = Format(CCur(frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 2) - frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 3)), "#,###.00")
'            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 8) = mcurAll_TRS_M
'            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 9) = CCur(frmBayar.txtSisaTagihanMultiTM.Text)
'            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 10) = CCur(frmBayar.txtSisaTagihanMultiTM.Text)
'            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 11) = 0
''            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 12) = mcurTM_HrsDibyr_M
''            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 13) = mcurTM_JmlByr_M
'            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 14) = CCur(frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 9) - frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 10))
'            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 15) = CCur(frmBayar.txtSisaTagihanMultiOA.Text)
'            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 16) = CCur(frmBayar.txtSisaTagihanMultiOA.Text)
''            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 17) = mcurOA_TRS_M
''            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 18) = mcurOA_HrsDibyr_M
''            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 19) = mcurOA_JmlByr_M
'            frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 20) = CCur(frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 15) - frmBayar.fgMulti.TextMatrix(frmBayar.fgMulti.Row, 16))
            
            With frmBayar.fgMulti
                    If mblnTM = True Then
                        .TextMatrix(.Row, 9) = CCur(.TextMatrix(.Row, 2))
                        .TextMatrix(.Row, 10) = CCur(.TextMatrix(.Row, 3))
                        .TextMatrix(.Row, 11) = 0
                        .TextMatrix(.Row, 12) = CCur(.TextMatrix(.Row, 2))
                        .TextMatrix(.Row, 13) = CCur(.TextMatrix(.Row, 4))
                        .TextMatrix(.Row, 14) = CCur(.TextMatrix(.Row, 5))

                        .TextMatrix(.Row, 15) = 0
                        .TextMatrix(.Row, 16) = 0
                        .TextMatrix(.Row, 17) = 0
                        .TextMatrix(.Row, 18) = 0
                        .TextMatrix(.Row, 19) = 0
                        .TextMatrix(.Row, 20) = 0
                    ElseIf mblnOA = True Then
                        .TextMatrix(.Row, 9) = 0
                        .TextMatrix(.Row, 10) = 0
                        .TextMatrix(.Row, 11) = 0
                        .TextMatrix(.Row, 12) = 0
                        .TextMatrix(.Row, 13) = 0
                        .TextMatrix(.Row, 14) = 0

                        .TextMatrix(.Row, 15) = CCur(.TextMatrix(.Row, 2))
                        .TextMatrix(.Row, 16) = CCur(.TextMatrix(.Row, 3))
                        .TextMatrix(.Row, 17) = 0
                        .TextMatrix(.Row, 18) = CCur(.TextMatrix(.Row, 2))
                        .TextMatrix(.Row, 19) = CCur(.TextMatrix(.Row, 4))
                        .TextMatrix(.Row, 20) = CCur(.TextMatrix(.Row, 5))

                    ElseIf mblnTM = True And mblnOA = True Then
                        .TextMatrix(.Row, 9) = CCur(.TextMatrix(.Row, 2) / 2)
                        .TextMatrix(.Row, 10) = CCur(.TextMatrix(.Row, 3) / 2)
                        .TextMatrix(.Row, 11) = 0
                        .TextMatrix(.Row, 12) = CCur(.TextMatrix(.Row, 2) / 2)
                        .TextMatrix(.Row, 13) = CCur(.TextMatrix(.Row, 4) / 2)
                        .TextMatrix(.Row, 14) = CCur(.TextMatrix(.Row, 5) / 2)

                        .TextMatrix(.Row, 15) = CCur(.TextMatrix(.Row, 2) / 2)
                        .TextMatrix(.Row, 16) = CCur(.TextMatrix(.Row, 3) / 2)
                        .TextMatrix(.Row, 17) = 0
                        .TextMatrix(.Row, 18) = CCur(.TextMatrix(.Row, 2) / 2)
                        .TextMatrix(.Row, 19) = CCur(.TextMatrix(.Row, 4) / 2)
                        .TextMatrix(.Row, 20) = CCur(.TextMatrix(.Row, 5) / 2)

                    End If
                End With

            Call frmBayar.subHitungTotal(frmBayar.fgMulti.Rows - 1, frmBayar.fgMulti.Cols)
            frmBayar.dcPenjamin.Visible = False
            frmBayar.fgMulti.SetFocus
            frmBayar.fgMulti.Col = 3
            
            
            
            
            
            
    
    End If
End If
End Sub

Private Sub txtAlamatPA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcHubungan.SetFocus
End Sub

Private Sub txtAlamatPA_LostFocus()
    txtAlamatPA = StrConv(txtAlamatPA, vbProperCase)
End Sub

Private Sub txtAnakKe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNoSJP.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtNamaPA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglLahirPA.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtNamaPA_LostFocus()
    txtNamaPA = StrConv(txtNamaPA, vbProperCase)
End Sub

Private Sub txtNipPA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcGolonganAsuransi.SetFocus
End Sub

Private Sub txtNipPA_LostFocus()
    txtNipPA = StrConv(txtNipPA, vbProperCase)
End Sub

Private Sub txtNoBP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcUnitKerja.SetFocus
End Sub

Private Sub txtNoKartuPA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtNamaPA.SetFocus
End Sub

Private Sub txtNoKartuPA_LostFocus()
    On Error GoTo errLoad
    Dim strKdGolongan As String

    strSQL = "SELECT * FROM AsuransiPasien " _
    & "WHERE IdAsuransi='" _
    & txtNoKartuPA.Text & "' and IdPenjamin = '" & dcPenjamin.BoundText & "'"
    
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
   ' If rs.RecordCount = 0 Then Exit Sub
    
    If rs.RecordCount <> 0 Then
     MsgBox "Nomor sudah dipakai oleh pasien lain", vbCritical, "Validasi"
  '   dcPenjamin.Enabled = False
     cmdSimpan.Enabled = False
     Exit Sub
    Else
     dcPenjamin.Enabled = True
     cmdSimpan.Enabled = True
       
    End If
    
    
    
    txtNamaPA.Text = rs.Fields("NamaPeserta").Value
    If Not IsNull(rs.Fields("IDPeserta").Value) Then txtNipPA.Text = rs.Fields("IDPeserta").Value
    dtpTglLahirPA.Value = rs.Fields("TglLahir").Value
    strKdGolongan = rs.Fields("KdGolongan").Value
    If Not IsNull(rs.Fields("Alamat").Value) Then txtAlamatPA.Text = rs.Fields("Alamat").Value
    strSQL = "SELECT NamaGolongan,KdGolongan FROM GolonganAsuransi WHERE KdGolongan='" & strKdGolongan & "' and StatusEnabled='1'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    dcKelasDitanggung.Text = rs.Fields(0).Value
    dcKelasDitanggung.BoundText = rs.Fields(1).Value
    Set rs = Nothing
    txtNoKartuPA = StrConv(txtNoKartuPA, vbProperCase)

    Exit Sub
errLoad:
End Sub

Private Sub txtNoKunjungan_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtNoRujukan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcNamaAsalRujukan.SetFocus
    If KeyAscii >= 65 And KeyAscii <= 90 Then
        Beep
        MsgBox "Harus Diisi Dengan Angka", vbCritical, "Validasi"
        KeyAscii = 0
    End If
End Sub


Private Sub txtNoSJP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dtpTglSJP.SetFocus
    End If
End Sub

Private Sub txtNoSJP_LostFocus()
    txtNoSJP = StrConv(txtNoSJP, vbProperCase)
End Sub

Private Function sp_AmbulNoKunjungan() As Boolean
    On Error GoTo errLoad
    sp_AmbulNoKunjungan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, dcPenjamin.BoundText)
        .Parameters.Append .CreateParameter("IdAsuransi", adChar, adParamInput, 15, Trim(txtNoKartuPA.Text))
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("KunjunganKe", adInteger, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("TglRujukanOut", adDate, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("TglPendaftaran", adDate, adParamInput, , Format(txttglpendaftaran.Text, "yyyy/MM/dd hh:mm:ss"))
        .Parameters.Append .CreateParameter("NoSJPRujukan", adVarChar, adParamInput, 30, Trim(txtNoSJP.Text))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Check_NoRujukan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam pengambilan No Kunjungan", vbExclamation, "Validasi"
            sp_AmbulNoKunjungan = False
        Else
            txtNoKunjungan.Text = .Parameters("KunjunganKe").Value
            If txtNoKunjungan.Text = "0" Then
                MsgBox "Masa berlaku No. Rujukan (SP3) sudah HABIS", vbExclamation, "Informasi"
                sp_AmbulNoKunjungan = False
            ElseIf val(txtNoKunjungan.Text) > 3 Then
                MsgBox "Masa kunjungan No. Rujukan (SP3) sudah lebih dari 3 kali", vbExclamation, "Informasi"
                sp_AmbulNoKunjungan = False
            End If
            Call Add_HistoryLoginActivity("Check_NoRujukan")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    Call msubPesanError
    sp_AmbulNoKunjungan = False
End Function

Private Function sp_UpdateJenisPasienUmum(f_KdKelompokPasien As String, f_NoPendaftaran As String) As Boolean
    On Error GoTo errLoad
    sp_UpdateJenisPasienUmum = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKelompokpasien", adChar, adParamInput, 2, f_KdKelompokPasien)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_JenisPasienUmumNew"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_UpdateJenisPasienUmum = False
        Else
            Call Add_HistoryLoginActivity("Update_JenisPasienUmumNew")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    sp_UpdateJenisPasienUmum = False
    Call msubPesanError
End Function

'Store procedure untuk mengisi asuransi pasien
'Private Sub sp_AsuransiPasien(ByVal adoCommand As ADODB.Command)
'Dim xrtSQL As String
'    Set dbcmd = New ADODB.Command
'
'    MousePointer = vbHourglass
'    With dbcmd
'        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
'        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, typAsuransi.strIdPenjamin)
'        .Parameters.Append .CreateParameter("IdAsuransi", adVarChar, adParamInput, 25, typAsuransi.strIdAsuransi)
'        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, mstrNoCM)
'        .Parameters.Append .CreateParameter("NamaPeserta", adVarChar, adParamInput, 50, typAsuransi.strNamaPeserta)
'
'        .Parameters.Append .CreateParameter("IDPeserta", adVarChar, adParamInput, 16, typAsuransi.strIdPeserta)
'        .Parameters.Append .CreateParameter("KdGolongan", adChar, adParamInput, 2, IIf(Len(Trim(typAsuransi.strKdGolongan)) = 0, Null, Trim(typAsuransi.strKdGolongan)))
'        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(typAsuransi.dTglLahir, "yyyy/MM/dd"))
'        .Parameters.Append .CreateParameter("Alamat", adVarChar, adParamInput, 100, typAsuransi.strAlamat)
'        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
'
'        .Parameters.Append .CreateParameter("KdHubungan", adChar, adParamInput, 2, typAsuransi.strHubungan)
'        If typAsuransi.strNoSJP <> "" Then
'            .Parameters.Append .CreateParameter("NoSJP", adVarChar, adParamInput, 30, typAsuransi.strNoSJP)
'        Else
'            .Parameters.Append .CreateParameter("NoSJP", adVarChar, adParamInput, 30, Null)
'        End If
'        .Parameters.Append .CreateParameter("TglSJP", adDate, adParamInput, , Format(typAsuransi.dTglSJP, "yyyy/MM/dd HH:mm:ss"))
'        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
'        .Parameters.Append .CreateParameter("NoBP", adVarChar, adParamInput, 3, IIf(Len(Trim(typAsuransi.strNoBp)) = 0, Null, Trim(typAsuransi.strNoBp)))
'
'        .Parameters.Append .CreateParameter("KunjunganKe", adInteger, adParamInput, , typAsuransi.intNoKunjungan)
'        .Parameters.Append .CreateParameter("OutputNoSJP", adVarChar, adParamOutput, 30, Null)
'        .Parameters.Append .CreateParameter("StatusNoSJP", adChar, adParamInput, 1, typAsuransi.strStatusNoSJP)
'        .Parameters.Append .CreateParameter("AnakKe", adInteger, adParamInput, , typAsuransi.intAnakKe)
'        .Parameters.Append .CreateParameter("UnitBagian", adVarChar, adParamInput, 50, IIf(Len(Trim(typAsuransi.strUnitBagian)) = 0, Null, Trim(typAsuransi.strUnitBagian)))
'
'        .Parameters.Append .CreateParameter("KdPaket", adVarChar, adParamInput, 3, Null)
'        .Parameters.Append .CreateParameter("NoRujukan", adVarChar, adParamInput, 30, typAsuransi.strNoRujukan)
'        .Parameters.Append .CreateParameter("KdRujukanAsal", adChar, adParamInput, 2, typAsuransi.strKdRujukanAsal)
'        .Parameters.Append .CreateParameter("DetailRujukanAsal", adVarChar, adParamInput, 100, typAsuransi.strDetailRujukanAsal)
'        .Parameters.Append .CreateParameter("KdDetailRujukanAsal", adChar, adParamInput, 8, typAsuransi.strKdDetailRujukanAsal)
'
'        .Parameters.Append .CreateParameter("NamaPerujuk", adVarChar, adParamInput, 50, typAsuransi.strNamaPerujuk)
'        .Parameters.Append .CreateParameter("TglDirujuk", adDate, adParamInput, , typAsuransi.dTglDirujuk)
'        .Parameters.Append .CreateParameter("DiagnosaRujukan", adVarChar, adParamInput, 100, typAsuransi.strDiagnosaRujukan)
'        .Parameters.Append .CreateParameter("KdDiagnosa", adVarChar, adParamInput, 7, typAsuransi.strKdDiagnosa)
'
'        '###24-4-2008 by john ----'edit splakuk
'        xrtSQL = "SELECT  KdinstitusiAsal, InstitusiAsal FROM InstitusiAsalPasien WHERE InstitusiAsal LIKE '" & typAsuransi.strPerusahaanPenjamin & "%' or KdInstitusiAsal LIKE '" & typAsuransi.strPerusahaanPenjamin & "' and StatusEnabled='1'"
'        Call msubRecFO(rsx, xrtSQL)
'        .Parameters.Append .CreateParameter("KdInstitusiAsal", adVarChar, adParamInput, 4, IIf(Len(Trim(rsx(0).Value)) = 0, Null, Trim(rsx(0).Value)))
'        .Parameters.Append .CreateParameter("KdKelasDiTanggung", adChar, adParamInput, 2, typAsuransi.strKdKelasDitanggung)
'
'        .ActiveConnection = dbConn
'        .CommandText = "AU_AsuransiPasienJoinProgramAskes"
'        .CommandType = adCmdStoredProc
'        .CommandTimeout = 120
'        .Execute
'
'        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
'            MsgBox "Ada kesalahan dalam pemasukan Asuransi Pasien", vbCritical, "Validasi"
'            mstrNoSJP = typAsuransi.strNoSJP
'        Else
'            mstrNoSJP = typAsuransi.strNoSJP
'            Call Add_HistoryLoginActivity("AU_AsuransiPasienJoinProgramAskes")
'        End If
'        Call deleteADOCommandParameters(dbcmd)
'        Set dbcmd = Nothing
'    End With
'    MousePointer = vbDefault
'    Exit Sub
'End Sub
