VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTagihanPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Tagihan Pasien"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTagihanPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   14895
   Begin VB.CommandButton cmdDiagnosa 
      Caption         =   "Diagnosa"
      Height          =   495
      Left            =   9000
      TabIndex        =   122
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Frame fraPosting 
      Height          =   1575
      Left            =   3240
      TabIndex        =   108
      Top             =   3000
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Timer Timer1 
         Interval        =   5
         Left            =   240
         Top             =   360
      End
      Begin MSComctlLib.ProgressBar pbPosting 
         Height          =   375
         Left            =   200
         TabIndex        =   109
         Top             =   1020
         Width           =   5900
         _ExtentX        =   10398
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Max             =   50
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   5985
         TabIndex        =   110
         Top             =   960
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.Label lblPosting 
         Alignment       =   2  'Center
         Caption         =   "VERIFIKASI HUTANG PENJAMIN DAN TANGGUNGAN RS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   111
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame fraPaketKhususJamsostek 
      Height          =   4815
      Left            =   120
      TabIndex        =   88
      Top             =   8640
      Visible         =   0   'False
      Width           =   14655
      Begin VB.CommandButton cmdaktifsimpan 
         Caption         =   "teretetetete"
         Height          =   255
         Left            =   11280
         TabIndex        =   112
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdHitung 
         Caption         =   "&Hitung"
         Height          =   375
         Left            =   9360
         TabIndex        =   101
         Top             =   4200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtTotalPembagian 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6600
         TabIndex        =   99
         Text            =   "0"
         Top             =   480
         Width           =   2250
      End
      Begin VB.TextBox txtTarifTanggungan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         TabIndex        =   97
         Text            =   "0"
         Top             =   480
         Width           =   2250
      End
      Begin VB.CheckBox chkCheckJamsostek 
         Caption         =   "Check3"
         Height          =   210
         Left            =   240
         TabIndex        =   96
         Top             =   1560
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.CheckBox chkBagiRata 
         Caption         =   "Bagi Rata"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   95
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CheckBox chkDitanggungPenjamin 
         Caption         =   "Ditanggung Penjamin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   94
         Top             =   4080
         Width           =   2415
      End
      Begin VB.CommandButton cmdProses 
         Caption         =   "&Proses"
         Height          =   375
         Left            =   11160
         TabIndex        =   93
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CommandButton cmdSelesai 
         Caption         =   "S&elesai"
         Height          =   375
         Left            =   12975
         TabIndex        =   92
         Top             =   4200
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid hgPaketKhususJamsostek 
         Height          =   2655
         Left            =   120
         TabIndex        =   89
         Top             =   1320
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   4683
         _Version        =   393216
         AllowUserResizing=   1
         Appearance      =   0
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
      Begin MSDataListLib.DataCombo dcPaketKhususJamsostek 
         Height          =   330
         Left            =   240
         TabIndex        =   90
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Total Pembagian"
         Height          =   210
         Left            =   6600
         TabIndex        =   100
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Tarif Paket"
         Height          =   210
         Left            =   3960
         TabIndex        =   98
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Nama Paket"
         Height          =   210
         Left            =   240
         TabIndex        =   91
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.PictureBox picUpdateKomponen 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   3960
      ScaleHeight     =   3825
      ScaleWidth      =   10065
      TabIndex        =   34
      Top             =   3000
      Visible         =   0   'False
      Width           =   10095
      Begin VB.TextBox txtTPembebasanUpdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   7950
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "txtTPembebasanUpdate"
         Top             =   3300
         Width           =   1700
      End
      Begin VB.TextBox txtTTanggunganRSUpdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   6255
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "txtTTanggunganRSUpdate"
         Top             =   3300
         Width           =   1700
      End
      Begin VB.TextBox txtTHutangPenjaminUpdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "txtTHutangPenjaminUpdate"
         Top             =   3300
         Width           =   1700
      End
      Begin VB.TextBox txtBarisKe 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Text            =   "txtBarisKe"
         Top             =   2520
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtIsiUpdate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   0
         MaxLength       =   15
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtKdPelayananUpdate 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Text            =   "txtKdPelayananUpdate"
         Top             =   2040
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtKdRuanganPelayananUpdate 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Text            =   "txtKdRuanganPelayananUpdate"
         Top             =   1560
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton cmdTutupUpdate 
         Caption         =   "Tutup"
         Height          =   495
         Left            =   2055
         TabIndex        =   39
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox txtNamaPelayananUpdate 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   360
         Width           =   5055
      End
      Begin VB.TextBox txtTglPelayananUpdate 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtRuangPelayananUpdate 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   360
         Width           =   2175
      End
      Begin MSFlexGridLib.MSFlexGrid fgUpdateKomponen 
         Height          =   2295
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4048
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pelayanan"
         Height          =   210
         Index           =   2
         Left            =   4680
         TabIndex        =   50
         Top             =   120
         Width           =   1320
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pelayanan"
         Height          =   210
         Index           =   1
         Left            =   2400
         TabIndex        =   49
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Ruang Pelayanan"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   120
         Width           =   1380
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Rekapitulasi Total Tagihan"
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
      TabIndex        =   74
      Top             =   6600
      Width           =   14895
      Begin VB.TextBox txtTotalPenjamin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   10920
         TabIndex        =   114
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox chkTagihanApotik 
         Caption         =   "Tagihan Apotik"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12600
         TabIndex        =   85
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtTRS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5880
         TabIndex        =   79
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtTAsuransi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3480
         TabIndex        =   78
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtTotalBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         TabIndex        =   77
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox chkDetail 
         Caption         =   "Detail"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   76
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtTotalPembebasan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8400
         TabIndex        =   75
         Top             =   480
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo dcRuanganApotik 
         Height          =   330
         Left            =   12600
         TabIndex        =   86
         Top             =   525
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Total Tanggungan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10800
         TabIndex        =   115
         Top             =   240
         Width           =   1725
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tanggungan Rumah Sakit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5880
         TabIndex        =   83
         Top             =   240
         Width           =   2445
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tanggungan Penjamin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3480
         TabIndex        =   82
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   81
         Top             =   240
         Width           =   2130
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Pembebasan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8400
         TabIndex        =   80
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.Frame fraDetailRekap 
      Height          =   2535
      Left            =   0
      TabIndex        =   51
      Top             =   4080
      Visible         =   0   'False
      Width           =   14895
      Begin VB.Frame fraRekapOA 
         Caption         =   "Rekapitulasi Obat && Alkes"
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
         Left            =   120
         TabIndex        =   63
         Top             =   1320
         Width           =   14655
         Begin VB.TextBox txtOA_HrsDibyr 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   12120
            TabIndex        =   68
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtOA_TRS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   7080
            TabIndex        =   67
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtOA_TP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   4560
            TabIndex        =   66
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtOA_TBP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2040
            TabIndex        =   65
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtOAPembebasan 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   9600
            TabIndex        =   64
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Harus Dibayar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   12120
            TabIndex        =   73
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Tanggungan Rumah Sakit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   7080
            TabIndex        =   72
            Top             =   240
            Width           =   2445
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Tanggungan Penjamin"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4560
            TabIndex        =   71
            Top             =   240
            Width           =   2115
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Total Biaya Pelayanan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2040
            TabIndex        =   70
            Top             =   240
            Width           =   2130
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Total Pembebasan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   9600
            TabIndex        =   69
            Top             =   240
            Width           =   1770
         End
      End
      Begin VB.Frame fraRekapTM 
         Caption         =   "Rekapitulasi Tindakan Medis"
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
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   14655
         Begin VB.TextBox txtTM_HrsDibyr 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   12120
            TabIndex        =   57
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtTM_TBP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2040
            TabIndex        =   56
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtTM_TP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   4560
            TabIndex        =   55
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtTM_TRS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   7080
            TabIndex        =   54
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtTMPembebasan 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   9600
            TabIndex        =   53
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Harus Dibayar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   12120
            TabIndex        =   62
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Total Biaya Pelayanan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2040
            TabIndex        =   61
            Top             =   240
            Width           =   2130
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Tanggungan Penjamin"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4560
            TabIndex        =   60
            Top             =   240
            Width           =   2115
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Tanggungan Rumah Sakit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   7080
            TabIndex        =   59
            Top             =   240
            Width           =   2445
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Total Pembebasan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   9600
            TabIndex        =   58
            Top             =   240
            Width           =   1770
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   32
      Top             =   8670
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5212
            Text            =   "Rincian Biaya Sementara (F1)"
            TextSave        =   "Rincian Biaya Sementara (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5212
            Text            =   "Ubah Biaya Pelayanan (F5)"
            TextSave        =   "Ubah Biaya Pelayanan (F5)"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5212
            Text            =   "Tambah Pelayanan (F6)"
            TextSave        =   "Tambah Pelayanan (F6)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5212
            Text            =   "Hapus Pelayanan (F7)"
            TextSave        =   "Hapus Pelayanan (F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5212
            Text            =   "Edit Tarif Tanggungan (K)"
            TextSave        =   "Edit Tarif Tanggungan (K)"
         EndProperty
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
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   0
      TabIndex        =   8
      Top             =   7560
      Width           =   14895
      Begin MSDataListLib.DataCombo dcDokter 
         Height          =   330
         Left            =   5040
         TabIndex        =   121
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CommandButton cmdTotalTanggunganPenjamin2 
         Caption         =   "Hitung Claim Ina Cbg's"
         Height          =   495
         Left            =   7800
         TabIndex        =   119
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtTambahanBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2280
         TabIndex        =   117
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdTambahData 
         Caption         =   "Tambah Data"
         Height          =   495
         Left            =   9000
         TabIndex        =   113
         Top             =   120
         Width           =   1935
      End
      Begin MSComctlLib.ProgressBar pbData 
         Height          =   375
         Left            =   2040
         TabIndex        =   107
         Top             =   960
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.PictureBox pbDataPicture 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   2040
         ScaleHeight     =   465
         ScaleWidth      =   5985
         TabIndex        =   105
         Top             =   960
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.CommandButton cmdPerbaikiData 
         Caption         =   "Perbaiki Data"
         Height          =   495
         Left            =   120
         TabIndex        =   104
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   12870
         TabIndex        =   4
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdBayar 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   10935
         TabIndex        =   3
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdPelayanan 
         Caption         =   "&Tambah Pelayanan"
         Height          =   375
         Left            =   9000
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdValidasi 
         Caption         =   "&Verifikasi"
         Height          =   375
         Left            =   2040
         TabIndex        =   106
         Top             =   360
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdTotalTanggunganPenjamin 
         Caption         =   "Hitung Proporsional Tanggungan Penjamin"
         Height          =   495
         Left            =   9000
         TabIndex        =   116
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "DPJP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5040
         TabIndex        =   120
         Top             =   120
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Tambahan Biaya Naik Kelas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2280
         TabIndex        =   118
         Top             =   120
         Visible         =   0   'False
         Width           =   2445
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Data Detail Pelayanan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   28
      Top             =   2880
      Width           =   14895
      Begin VB.TextBox txtIsi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkCheck 
         Height          =   210
         Left            =   480
         TabIndex        =   29
         Top             =   4000
         Visible         =   0   'False
         Width           =   200
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgTagihanPasien1 
         Height          =   3135
         Left            =   2040
         TabIndex        =   1
         Top             =   -3000
         Visible         =   0   'False
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   5530
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSFlexGridLib.MSFlexGrid hgTagihanPasien 
         Height          =   3135
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   5530
         _Version        =   393216
         AllowUserResizing=   1
         Appearance      =   0
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
      Begin MSDataGridLib.DataGrid dgTagihanPasien 
         Height          =   3135
         Left            =   240
         TabIndex        =   30
         Top             =   3720
         Visible         =   0   'False
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   5530
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
      TabIndex        =   9
      Top             =   1800
      Width           =   14895
      Begin VB.TextBox txtPenjamin 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   12000
         TabIndex        =   27
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtJenisPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   10680
         TabIndex        =   25
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   6960
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         TabIndex        =   23
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   22
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   1335
      End
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
         Left            =   8160
         TabIndex        =   15
         Top             =   350
         Width           =   2415
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
            TabIndex        =   18
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
            TabIndex        =   17
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtHari 
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
            TabIndex        =   16
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Left            =   550
            TabIndex        =   21
            Top             =   277
            Width           =   285
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Left            =   1350
            TabIndex        =   20
            Top             =   277
            Width           =   240
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Left            =   2130
            TabIndex        =   19
            Top             =   270
            Width           =   165
         End
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Penjamin"
         Height          =   210
         Left            =   12000
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pasien"
         Height          =   210
         Left            =   10680
         TabIndex        =   14
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   6960
         TabIndex        =   13
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   3480
         TabIndex        =   12
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1560
         TabIndex        =   11
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Total Tagihan Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   14895
      Begin VB.CheckBox chkPaketJamsostek 
         Caption         =   "Paket Khusus Jamsostek"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   87
         Top             =   300
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   7680
         X2              =   7680
         Y1              =   240
         Y2              =   720
      End
      Begin VB.Label LblTgihanSebelumnya 
         Caption         =   "Rp. 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   5160
         TabIndex        =   103
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label25 
         Caption         =   "Tagihan Sebelumnya (F8) -->"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   102
         Top             =   360
         Width           =   4005
      End
      Begin VB.Label lblTotalTagihan 
         Alignment       =   1  'Right Justify
         Caption         =   "Rp. 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   10920
         TabIndex        =   7
         Top             =   240
         Width           =   3840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Tagihan ->"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7800
         TabIndex        =   6
         Top             =   240
         Width           =   3000
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   84
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
      Picture         =   "frmTagihanPasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   13080
      Picture         =   "frmTagihanPasien.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTagihanPasien.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13455
   End
End
Attribute VB_Name = "frmTagihanPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strKdKelPsn As String
Dim mstrNoCMKu As String
Dim subbolEditTanggungan As Boolean
Dim intPembayaranKe As Integer
Dim mcurTagihansebelumnya As Currency
Dim sfilter As String
Dim intJmlDataDipilih As Integer
Dim curTotalBiayaDipilih As Currency
Dim curTotalTanggunganPaket As Currency
Dim curTarifPelayananSewaKamar As Currency
Dim curTarifPelayananVisitDokter As Currency
Dim curTRS As Currency
Dim curPembebasan As Currency
Dim strStatusData As String
Dim sqlQuery As String
Dim rsQuery As New ADODB.recordset
Dim bolValUpdate As Boolean
Dim bolinacbgs As Boolean
Dim curnilaiproposional As Currency
Dim curtotalbiaya As Currency

'Private Sub HitungClaimInaCbg()
'
'
'    If (Dir("c:\SDK\inacbg\result.tlb") <> "") Then
'        Dim context As BridgingInaCbg.context
'        bolinacbgs = True
'
'        Dim Hasil As String
'        'baru sampai sini
'        strSQL = "select value from SettingGlobal where Prefix='UrlInaCbg'"
'        Call msubRecFO(rs, strSQL)
'        If (rs.EOF = False) Then
'            Set context = New BridgingInaCbg.context
'            context.SetEndpoint (rs(0).Value)
'            strSQL = "select * from pasien where noCm='" & mstrNoCM & "'"
'            Call msubRecFO(rsC, strSQL)
'            strSQL = "SELECT     PeriksaDiagnosa.KdDiagnosa FROM PeriksaDiagnosa INNER JOIN SettingGlobal ON PeriksaDiagnosa.KdJenisDiagnosa = SettingGlobal.Value WHERE     (PeriksaDiagnosa.NoPendaftaran = '" & mstrNoPen & "') AND (SettingGlobal.Prefix = 'KdDiagnosaUtama' )"
'            Call msubRecFO(rsD, strSQL)
'            Dim diagnosa As String
'            diagnosa = ""
'            Dim i As Integer
'            For i = 1 To rsD.RecordCount
'                If (diagnosa = "") Then
'                    diagnosa = rsD(0).Value
'                Else
'                    diagnosa = rsD(0).Value & ";"
'                End If
'                rsD.MoveNext
'            Next i
'
'            strSQL = "SELECT     PeriksaDiagnosa.KdDiagnosa FROM PeriksaDiagnosa INNER JOIN SettingGlobal ON PeriksaDiagnosa.KdJenisDiagnosa = SettingGlobal.Value WHERE     (PeriksaDiagnosa.NoPendaftaran = '" & mstrNoPen & "') AND (SettingGlobal.Prefix = 'KdDiagnosaTambahan' )"
'            Call msubRecFO(rsD, strSQL)
'
'            For i = 1 To rsD.RecordCount
'                diagnosa = diagnosa & ";" & rsD(0).Value
'                rsD.MoveNext
'            Next i
'
'             strSQL = "SELECT     DetailPeriksaDiagnosa.KdDiagnosaTindakan FROM DetailPeriksaDiagnosa INNER JOIN SettingGlobal ON DetailPeriksaDiagnosa.KdJenisDiagnosa = SettingGlobal.Value WHERE     (DetailPeriksaDiagnosa.NoPendaftaran = '" & mstrNoPen & "') AND (SettingGlobal.Prefix = 'KdDiagnosaUtama' )"
'            Call msubRecFO(rsD, strSQL)
'            Dim diagnosaTindakan  As String
'            diagnosaTindakan = ""
'            For i = 1 To rsD.RecordCount
'                diagnosaTindakan = diagnosaTindakan & rsD(0).Value + ";"
'                rsD.MoveNext
'            Next i
'
'            strSQL = "SELECT     DetailPeriksaDiagnosa.KdDiagnosaTindakan FROM DetailPeriksaDiagnosa INNER JOIN SettingGlobal ON DetailPeriksaDiagnosa.KdJenisDiagnosa = SettingGlobal.Value WHERE     (DetailPeriksaDiagnosa.NoPendaftaran = '" & mstrNoPen & "') AND (SettingGlobal.Prefix = 'KdDiagnosaTambahan' )"
'            Call msubRecFO(rsD, strSQL)
'
'            For i = 1 To rsD.RecordCount
'                diagnosaTindakan = diagnosaTindakan & rsD(0).Value + ";"
'                rsD.MoveNext
'            Next i
'
'            strSQL = "select * from RegistrasiRJ where NoPendaftaran='" & mstrNoPen & "' "
'            Call msubRecFO(rsD, strSQL)
'            '
'            strSQL = "select * from pasienDaftar where NoPendaftaran='" & mstrNoPen & "' "
'            Call msubRecFO(rsE, strSQL)
'            strSQL = "select * from KelasPelayanan where KdKelas='" & rsE("KdKelasAkhir") & "' "
'            Call msubRecFO(rsF, strSQL)
'            'context.SimulasiTarif(
'            txtTotalPenjamin.Text = context.SimulasiTarif(IIf(rsC("JenisKelamin").Value = "L", "m", "f"), "5", "-", "-", IIf(rsD.RecordCount <> 0, "rawat jalan", "rawat inap"), rsF("KodeExternal"), Format(rsE("TglPendaftaran").Value, "yyyy-MM-dd"), Format(rsE("TglPulang").Value, "yyyy-MM-dd"), "home", "-", "1000", diagnosa, diagnosaTindakan, "", mstrNoCM, rsC("NamaLengkap").Value, Format(rsC("TglLahir").Value, "yyyy-MM-dd"))
'            Dim cmd   As New ADODB.Command
'            sp_AU_TotalBiayaKlaimBPJS cmd
'            txtTotalPenjamin.Text = FormatCurrency(txtTotalPenjamin.Text)
'
'            'Perhitungan Pembagian Cliam
'
'                   curtotalbiaya = txtTotalBiaya.Text
'                   subbolEditTanggungan = True
'
'                   Call txtNoPendaftaran_KeyPress(13)
'                  ' Call txtIsi_KeyPress(13)
'
'
'        End If
'    End If
'End Sub

Private Sub HitungClaimInaCbg()
'Dim cmd   As New ADODB.Command
'    sp_AU_TotalBiayaKlaimBPJS cmd
'    txtTotalPenjamin.Text = FormatCurrency(txtTotalPenjamin.Text)
'
'    'Perhitungan Pembagian Cliam
'
'    curtotalbiaya = txtTotalBiaya.Text
'    subbolEditTanggungan = True
'
'    Call txtNoPendaftaran_KeyPress(13)

'On Error GoTo hell
'Dim strKoefisienTambahanBiayakeVIP As String
'Dim context As BridgingInaCbg.context
'Dim tmpBalikanWSINACBGS As String
'Dim strHasilSync As String
'Dim arrHasilSimulasiTarifGrouper() As String
'Dim i As Long
'
'    If (Dir(strGlobalPathReferenceBridgingINACBG) <> "") Then
'        If (strGlobalUrlINACBG <> "") Then
'            Set context = New BridgingInaCbg.context
'            context.SetEndpoint strGlobalUrlINACBG
'            context.SetKey strGlobalINACBGKeyEnkripDanDekripEklaim
'
'            If strGlobalINACBGVersiEklaim = "5.1" Then
'                If strGlobalINACBGUrlEklaim51 <> "" Then
'                    context.SetEndpoint strGlobalINACBGUrlEklaim51
'                End If
'            End If
'
'            Call sp_DelBiayaPelayananNaikKelas(dbcmd)
'            txtNoPendaftaran_KeyPress 13
'
'        End If
'    End If
End Sub

Private Sub HitungClaimInaCbgNew()
Dim currBiayaTambahanYgHarusDibayarPasien As Currency
Dim currBiayaTambahanYgHarusDibayarPasienDiKelas1 As Currency
Dim currBiayaTambahanYgHarusDibayarPasienDiKelas2 As Currency
Dim currBiayaTambahanYgHarusDibayarPasienDiKelas3 As Currency
Dim currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya As Currency
Dim currBiayaTambahanDariKelas1KeVIP As Currency
Dim strPersentaseDariKelas1 As String

Dim strNIKPegawai As String

Dim context As BridgingInaCbg.context
Dim tmpBalikanDariWebServiceINACBGS As String
Dim tmpBalikanDariWebServiceINACBGS2 As String
            
      If (Dir("c:\SDK\inacbg\result.tlb") <> "") Then
        
'        strSQL = "select NIKINACBG from INACBGNIKPegawai where IdPegawai='" & strIDPegawaiAktif & "'"
'        Call msubRecFO(rs, strSQL)
'        If rs.RecordCount > 0 Then
'            strNIKPegawai = rs("NIKINACBG")
'        Else
            strNIKPegawai = "9022C01" 'default ke Admin
'        End If
        
        strSQL = "select value from SettingGlobal where Prefix='UrlInaCbg'"
        Call msubRecFO(rs, strSQL)
        If (rs.EOF = False) Then
            
            Set context = New BridgingInaCbg.context
            
            context.SetEndpoint (rs(0).Value)
                        
'            Call sp_DelBiayaPelayananNaikKelas(dbcmd)
            txtNoPendaftaran_KeyPress 13
            
            Dim strHasilSyncron As String
            Dim strKdKelasDitanggung As String
            
'            strHasilSyncron = SyncronINACBGPerPasien(mstrNoPen, txtTotalBiaya.Text)

            strKdKelasDitanggung = ""
            If rs.EOF = False Then
                strKdKelasDitanggung = rs("KdKelasDiTanggung")
            End If
            
            If strHasilSyncron <> "" Then
                MsgBox strHasilSyncron, vbCritical
                Exit Sub
            End If
                    
           ' context.SimulasiTarifGrouper51 (Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""))
'            tmpBalikanDariWebServiceINACBGS = context.SimulasiTarifGrouper51(Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""))
            
'            tmpBalikanDariWebServiceINACBGS2 = context.HapusClaim(Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""), strNIKPegawai)
'            If tmpBalikanDariWebServiceINACBGS = 0 Then MsgBox "Diagnosa salah", vbCritical: Exit Sub


            'Biaya tambahan di kelas haknya
'            Call SyncronINACBGPerPasien(mstrNoPen, txtTotalBiaya.Text, True)
'            Call SyncronINACBGPerPasien(mstrNoPen, txtTotalBiaya.Text, True, strGlobalKelasPerawatan)
'            currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya = context.SimulasiTarifGrouper51(Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""))
'            tmpBalikanDariWebServiceINACBGS2 = context.HapusClaim(Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""), strNIKPegawai)
            
'            'Biaya tambahan di kelas 2
'                    Call SyncronINACBGPerPasien(mstrNoPen, TxtTotalBiaya.Text, True, "2")
'                    currBiayaTambahanYgHarusDibayarPasienDiKelas2 = context.SimulasiTarifGrouper51(Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""))
'                    tmpBalikanDariWebServiceINACBGS2 = context.HapusClaim(Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""), strNIKPegawai)

            'Biaya tambahan di kelas 1
'            Call SyncronINACBGPerPasien(mstrNoPen, TxtTotalBiaya.Text, True, "1")
'            currBiayaTambahanYgHarusDibayarPasienDiKelas1 = context.SimulasiTarifGrouper51(Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""))
'            tmpBalikanDariWebServiceINACBGS2 = context.HapusClaim(Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""), strNIKPegawai)
                        

            If strGlobalNamaKelasTertinggiNaikKelas = "vip" Then
                MsgBox "Total Biaya RS : " & FormatCurrency(txtTotalBiaya.Text) & vbCrLf & "Tanggungan InaCBG Hak Kelas Pasien : " & FormatCurrency(currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya) & vbCrLf & "Tanggungan InaCBG Kelas 1 : " & FormatCurrency(currBiayaTambahanYgHarusDibayarPasienDiKelas1) & vbCrLf & "Tanggungan InaCBG Kelas 2 : " & FormatCurrency(currBiayaTambahanYgHarusDibayarPasienDiKelas2) & vbCrLf & "Selisih Biaya RS - Tanggungan InaCBG : " & FormatCurrency(CCur(txtTotalBiaya.Text) - currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya) & vbCrLf & "Persentase Selisih Biaya RS & Tanggungan InaCBG " & CCur(Abs(((CCur(txtTotalBiaya.Text) - currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya) * 100)) / currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya) & " %"
            Else
                MsgBox "Total Biaya RS : " & FormatCurrency(txtTotalBiaya.Text) & vbCrLf & "Tanggungan InaCBG Hak Kelas Pasien : " & FormatCurrency(currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya) & vbCrLf & "Tanggungan InaCBG Kelas 1 : " & FormatCurrency(currBiayaTambahanYgHarusDibayarPasienDiKelas1) & vbCrLf & "Tanggungan InaCBG Kelas 2 : " & FormatCurrency(currBiayaTambahanYgHarusDibayarPasienDiKelas2)
            End If
            'Biaya tambahan jika ada kenaikan kelas
            If strGlobalAdaKenaikanKelas = "1" Then
                
                
                If strGlobalNamaKelasTertinggiNaikKelas = "vvip" Then
                    
                    'Biaya tambahan naik kelas total
                    currBiayaTambahanYgHarusDibayarPasien = CCur(txtTotalBiaya.Text) - currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya
                    
                ElseIf strGlobalNamaKelasTertinggiNaikKelas = "vip" Then
                        
                    'Matikan kodingan ini jika persentasenya tidak memakai rumus khusus seperti di kanujoso
                    'Rumus persentase ke VIP Madya khusus kanujoso
                    'Begin----------------------------------------------------------------------------------------------------------------------------------------
                    Dim currSelisihTarifRSDenganKelasHaknya, currSelisihTarifINACBGKelas1DenganKelasHaknya, currPenguranganSelisih, currBatasMaksimalTambahanBiaya As Currency
                    Dim dblPersentaseDariKelas1 As Double
                    
                    strsqlx5 = "Select Value From SettingGlobal Where Prefix='KdKelasVIPMadya'"
                    Call msubRecFO(rsJ, strsqlx5)
                    If rsJ.RecordCount > 0 Then
                        If rsJ("Value") = strGlobalKdKelasNaikKelas Then
                            currSelisihTarifRSDenganKelasHaknya = CCur(txtTotalBiaya.Text) - currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya
                            currSelisihTarifINACBGKelas1DenganKelasHaknya = currBiayaTambahanYgHarusDibayarPasienDiKelas1 - currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya
                            currPenguranganSelisih = currSelisihTarifRSDenganKelasHaknya - currSelisihTarifINACBGKelas1DenganKelasHaknya
                            currBatasMaksimalTambahanBiaya = (75 * currBiayaTambahanYgHarusDibayarPasienDiKelas1) / 100
                            If currPenguranganSelisih > currBatasMaksimalTambahanBiaya Then
                                dblPersentaseDariKelas1 = 75
                            Else
                                dblPersentaseDariKelas1 = (currPenguranganSelisih / currBiayaTambahanYgHarusDibayarPasienDiKelas1) * 100
                            End If
                        Else
                            strPersentaseDariKelas1 = InputBox("Masukan Persentase Dari Kelas 1", vbOKCancel)
                            dblPersentaseDariKelas1 = CDbl(Replace(strPersentaseDariKelas1, ".", ","))
                            
                            If dblPersentaseDariKelas1 > 75 Then
                                MsgBox "Persentase tidak boleh lebih dari sama dengan 75", vbCritical
                                tmpBalikanDariWebServiceINACBGS2 = context.HapusClaim(Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""), strNIKPegawai)
                                Exit Sub
                            End If
                            
                            If dblPersentaseDariKelas1 < 0 Then
                                MsgBox "Persentase tidak boleh kurang dari 0", vbCritical
                                tmpBalikanDariWebServiceINACBGS2 = context.HapusClaim(Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""), strNIKPegawai)
                                Exit Sub
                            End If
                            
                        End If
                    Else
                    
                        strPersentaseDariKelas1 = InputBox("Masukan Persentase Dari Kelas 1", vbOKCancel)
                        dblPersentaseDariKelas1 = CDbl(Replace(strPersentaseDariKelas1, ".", ","))
                        
                        If dblPersentaseDariKelas1 > 75 Then
                            MsgBox "Persentase tidak boleh lebih dari sama dengan 75", vbCritical
                            tmpBalikanDariWebServiceINACBGS2 = context.HapusClaim(Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""), strNIKPegawai)
                            Exit Sub
                        End If
                        
                        If dblPersentaseDariKelas1 < 0 Then
                            MsgBox "Persentase tidak boleh kurang dari 0", vbCritical
                            tmpBalikanDariWebServiceINACBGS2 = context.HapusClaim(Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""), strNIKPegawai)
                            Exit Sub
                        End If
                        
                    End If
                    'End-----------------------------------------------------------------------------------------------------------------------------------------
                    
                    
'                    'Hidupkan kodingan ini jika persentasenya tidak memakai rumus khusus seperti di kanujoso
'                    strPersentaseDariKelas1 = InputBox("Masukan Persentase Biaya Naik Kelas Ke Kelas VIP", vbOKCancel)
'                    dblPersentaseDariKelas1 = CDbl(strPersentaseDariKelas1)
                    
                    
                    
                    currBiayaTambahanDariKelas1KeVIP = (dblPersentaseDariKelas1 * currBiayaTambahanYgHarusDibayarPasienDiKelas1) / 100
                    
                    
                    'Kelas 1 ke VIP
                    If strGlobalKelasPerawatan = "1" Then
                        currBiayaTambahanYgHarusDibayarPasien = currBiayaTambahanDariKelas1KeVIP
                        
                    Else 'Kelas 2/3 ke VIP
                       
                                                        
                        currBiayaTambahanYgHarusDibayarPasien = currBiayaTambahanDariKelas1KeVIP + (currBiayaTambahanYgHarusDibayarPasienDiKelas1 - currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya)
                        
                    End If
                
                ElseIf strGlobalNamaKelasTertinggiNaikKelas = "kelas_1" Then   'Kelas 2/3 ke kelas 1
                        
                                                    
                    currBiayaTambahanYgHarusDibayarPasien = currBiayaTambahanYgHarusDibayarPasienDiKelas1 - currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya
                    
                ElseIf strGlobalNamaKelasTertinggiNaikKelas = "kelas_2" Then   'kelas 3 ke kelas 2
                                        
'                    'Biaya tambahan di kelas 2
'                    Call SyncronINACBGPerPasien(mstrNoPen, txtTotalBiaya.Text, True, "2")
'                    currBiayaTambahanYgHarusDibayarPasienDiKelas2 = context.SimulasiTarifGrouper51(Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""))

                    
                    currBiayaTambahanYgHarusDibayarPasien = currBiayaTambahanYgHarusDibayarPasienDiKelas2 - currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya
                                                                                                                                                              
                End If
            Else
                Call SyncronINACBGPerPasien(mstrNoPen, txtTotalBiaya.Text)
'                strTmpBalikanDariWebServiceINACBGS = context.SimulasiTarifGrouper(Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""))
                
            End If
            
            
            txtTambahanBiaya.Text = FormatCurrency("0")
            If strGlobalAdaKenaikanKelas = "1" Then
                '######################################################################################
                'khusus ponorogo
                If strGlobalNamaKelasTertinggiNaikKelas = "vip" Then
                
                    If currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya > CCur(txtTotalBiaya) And strKdKelasDitanggung = "03" Then ' tanggungan kelas1 ke vip/utama, klo biaya rs<dari tanggungannya maka biaya naik kelasnya 0
                        currBiayaTambahanYgHarusDibayarPasien = 0
'                    'ElseIf ((CCur(txtTotalBiaya.Text) - currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya) * 100) / currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya > 75 Then
'                        'currBiayaTambahanYgHarusDibayarPasien = currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya * 0.75
'                    ElseIf ((CCur(txtTotalBiaya.Text) - currBiayaTambahanYgHarusDibayarPasienDiKelas1) * 100) / currBiayaTambahanYgHarusDibayarPasienDiKelas1 > 75 Then
                    ElseIf ((CCur(txtTotalBiaya.Text) - currBiayaTambahanYgHarusDibayarPasienDiKelas1) * 100) / currBiayaTambahanYgHarusDibayarPasienDiKelas1 > 75 Then
                        currBiayaTambahanYgHarusDibayarPasien = (currBiayaTambahanYgHarusDibayarPasienDiKelas1 - currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya) + (currBiayaTambahanYgHarusDibayarPasienDiKelas1 * 0.75)
                    ElseIf CCur(txtTotalBiaya.Text) < currBiayaTambahanYgHarusDibayarPasienDiKelas1 Or (CCur(txtTotalBiaya.Text) - currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya) <= 0 Then
                        currBiayaTambahanYgHarusDibayarPasien = currBiayaTambahanYgHarusDibayarPasienDiKelas1 - currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya
                    Else 'selisih biayars-tanggungan kelas 1 <75%
                        currBiayaTambahanYgHarusDibayarPasien = CCur(txtTotalBiaya.Text) - currBiayaTambahanYgHarusDibayarPasienDiKelasHaknya
                    End If
                End If
                '######################################################################################
                txtTambahanBiaya.Text = FormatCurrency(currBiayaTambahanYgHarusDibayarPasien)
            End If
            
            
'            tmpBalikanDariWebServiceINACBGS2 = context.HapusClaim(Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""), strNIKPegawai)
            
'            Call SyncronINACBGPerPasien(mstrNoPen, txtTotalBiaya.Text)
'            tmpBalikanDariWebServiceINACBGS = context.SimulasiTarifGrouper51(Replace(Replace(rs("NoSJP"), Chr(13), ""), Chr(10), ""))

            
            If CCur(txtTambahanBiaya.Text) > 0 Then

                Call sp_BiayaPelayananNaikKelas(dbcmd, txtTambahanBiaya.Text)
                txtNoPendaftaran_KeyPress 13
                
                    
            End If
            
            
            
            txtTotalPenjamin.Text = tmpBalikanDariWebServiceINACBGS
            'MsgBox "Tarif Total InaCbg " & FormatCurrency(txtTotalPenjamin.Text), vbInformation, "Informasi"
            MsgBox "Biaya Naik Kelas " & FormatCurrency(txtTambahanBiaya.Text), vbInformation, "Informasi"
            
            
            Dim cmd   As New ADODB.Command
'            sp_AU_TotalBiayaKlaimBPJS cmd
            
'            txtCostSharing.Text = tmpBalikanDariWebServiceINACBGS
            txtTotalPenjamin.Text = tmpBalikanDariWebServiceINACBGS
 '           txtCostSharing_KeyPress (13)
            
            txtTAsuransi.BackColor = vbWhite
            If CCur(txtTAsuransi.Text) < CCur(txtTotalBiaya.Text) Then
                txtTAsuransi.BackColor = vbYellow
            End If
            
           ' cmdBayar = True
            
'            txtCostSharing.Text = ""
            txtTotalPenjamin.Text = ""
            
        End If
    End If

    strGlobalAdaKenaikanKelas = ""
    strGlobalKelasPerawatan = ""
    strGlobalNamaKelasTertinggiNaikKelas = ""



Exit Sub

duaTambahDuaSamaDenganLima:

Call msubPesanError
'Resume 0

End Sub

Private Sub sp_AU_TotalBiayaKlaimBPJS(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("TotalKlaim", adDouble, adParamInput, 3, txtTotalPenjamin.Text)
        .Parameters.Append .CreateParameter("TotalYangDibayar", adDouble, adParamInput, , -1)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AU_TotalBiayaKlaimBPJS"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam Penginputan Total Biaya", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("AU_TotalBiayaKlaimBPJS")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

Private Sub subHitungTotalUpdateKomponen()
    On Error GoTo errLoad
    Dim i As Integer
    txtTHutangPenjaminUpdate.Text = 0: txtTTanggunganRSUpdate.Text = 0: txtTPembebasanUpdate.Text = 0
    For i = 1 To fgUpdateKomponen.Rows - 2
        txtTHutangPenjaminUpdate.Text = CCur(txtTHutangPenjaminUpdate.Text) + CCur(fgUpdateKomponen.TextMatrix(i, 3))
        txtTTanggunganRSUpdate.Text = CCur(txtTTanggunganRSUpdate.Text) + CCur(fgUpdateKomponen.TextMatrix(i, 4))
        txtTPembebasanUpdate.Text = CCur(txtTPembebasanUpdate.Text) + CCur(fgUpdateKomponen.TextMatrix(i, 5))
    Next i
    txtTHutangPenjaminUpdate.Text = Format(txtTHutangPenjaminUpdate.Text, "#,###.00")
    txtTTanggunganRSUpdate.Text = Format(txtTTanggunganRSUpdate.Text, "#,###.00")
    txtTPembebasanUpdate.Text = Format(txtTPembebasanUpdate.Text, "#,###.00")
    Exit Sub
errLoad:
    Call msubPesanError("subHitungTotalUpdateKomponen")
End Sub

Public Sub subLoadDataKomponenPel()
    On Error GoTo errLoad
    Dim i As Integer

    If LCase(hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 25)) = "oa" Then Exit Sub
    strSQL = "SELECT KomponenTarif.NamaKomponen, TempHargaKomponen.JmlPelayanan, TempHargaKomponen.Harga, TempHargaKomponen.JmlHutangPenjamin, TempHargaKomponen.JmlTanggunganRS , TempHargaKomponen.JmlPembebasan, TempHargaKomponen.KdKomponen" & _
    " FROM TempHargaKomponen INNER JOIN KomponenTarif ON TempHargaKomponen.KdKomponen = KomponenTarif.KdKomponen " & _
    " WHERE (TempHargaKomponen.NoPendaftaran = '" & txtNoPendaftaran.Text & "') " & _
    " AND (TempHargaKomponen.KdRuangan = '" & hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 20) & "') " & _
    " AND (TempHargaKomponen.TglPelayanan = '" & Format(hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 8), "yyyy/MM/dd HH:mm:ss") & "') " & _
    " AND (TempHargaKomponen.KdPelayananRS = '" & Trim(hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 3)) & "')" & _
    " AND (TempHargaKomponen.KdKomponen <> '12')"
    Call msubRecFO(rs, strSQL)
    Call subSetGridUpdatekomponen
    If rs.EOF = True Then
        txtRuangPelayananUpdate.Text = ""
        txtTglPelayananUpdate.Text = ""
        txtNamaPelayananUpdate.Text = ""
        txtKdRuanganPelayananUpdate.Text = ""
        txtKdPelayananUpdate.Text = ""
        txtBarisKe.Text = ""
    Else
        picUpdateKomponen.Left = (frmTagihanPasien.Width - picUpdateKomponen.Width) / 2
        picUpdateKomponen.Top = Frame6.Top
        picUpdateKomponen.Visible = True
        txtRuangPelayananUpdate.Text = hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 10)
        txtTglPelayananUpdate.Text = hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 8)
        txtNamaPelayananUpdate.Text = hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 4)
        txtKdRuanganPelayananUpdate.Text = hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 20)
        txtKdPelayananUpdate.Text = hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 3)
        txtBarisKe.Text = hgTagihanPasien.Row
        cmdUpdate.Enabled = True
        For i = 1 To rs.RecordCount
            fgUpdateKomponen.TextMatrix(i, 0) = rs("NamaKomponen")
            fgUpdateKomponen.TextMatrix(i, 1) = rs("JmlPelayanan")
            fgUpdateKomponen.TextMatrix(i, 2) = Format(rs("Harga"), "#,###.00")
            fgUpdateKomponen.TextMatrix(i, 3) = rs("JmlHutangPenjamin")
            fgUpdateKomponen.TextMatrix(i, 4) = rs("JmlTanggunganRS")
            fgUpdateKomponen.TextMatrix(i, 5) = rs("JmlPembebasan")
            fgUpdateKomponen.TextMatrix(i, 6) = rs("KdKomponen")
            fgUpdateKomponen.Rows = fgUpdateKomponen.Rows + 1
            rs.MoveNext
        Next i
        Call subHitungTotalUpdateKomponen
    End If
    Exit Sub
errLoad:
    Call msubPesanError("subLoadDataKomponenPel")
End Sub

Private Sub subSetGridUpdatekomponen()
    With fgUpdateKomponen
        .Cols = 7
        .Rows = 2
        .FixedCols = 0

        .RowHeight(0) = 400
        .TextMatrix(0, 0) = "Komponen Tarif"
        .TextMatrix(0, 1) = "Jumlah"
        .TextMatrix(0, 2) = "Harga"
        .TextMatrix(0, 3) = "JmlHutangPenjamin"
        .TextMatrix(0, 4) = "JmlTanggunganRS"
        .TextMatrix(0, 5) = "JmlPembebasan"
        .TextMatrix(0, 6) = "KdKomponen"

        .ColWidth(0) = 2000
        .ColWidth(1) = 700
        .ColWidth(2) = 1700
        .ColWidth(3) = 1700
        .ColWidth(4) = 1700
        .ColWidth(5) = 1700
        .ColWidth(6) = 0
    End With
End Sub

Private Sub chkBagiRata_Click()
    On Error GoTo hell
    Dim dblRata2 As Double
    Dim i As Integer
    If chkBagiRata.Value = Checked Then
        If CCur(txtTarifTanggungan.Text) = 0 Then
            MsgBox "Nama Paket kosong", vbExclamation, "Validasi"
            chkBagiRata.Value = Unchecked
            Exit Sub
        End If
        If txtTotalBiaya.Text = "" Then
            MsgBox "Total Biaya Pelayanan kosong, Hubungi administrator", vbExclamation, "Validasi"
            chkBagiRata.Value = Unchecked
            Exit Sub
        End If

        txtTotalPembagian.Text = "0"
        With hgPaketKhususJamsostek
            For i = 1 To .Rows - 1
                dblRata2 = CDec(CCur(.TextMatrix(i, 7)) / CCur(txtTotalBiaya.Text)) * CCur(txtTarifTanggungan.Text)
                .TextMatrix(i, 21) = Format(dblRata2, "###,###.###")
                txtTotalPembagian.Text = Format(CCur(txtTotalPembagian.Text) + CCur(dblRata2), "###,###")
                hgPaketKhususJamsostek.TextMatrix(i, 1) = Chr$(187)
                hgPaketKhususJamsostek.TextMatrix(i, 34) = 1
            Next i
        End With
    Else
        Call chkPaketJamsostek_Click
    End If

    Exit Sub
hell:
    msubPesanError
    chkBagiRata.Value = Unchecked
End Sub

Private Sub chkCheck_Click()
    On Error GoTo errLoad

    If chkCheck.Value = vbChecked Then
        hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, hgTagihanPasien.Col) = Chr$(187)
        hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 34) = 1
    Else
        hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, hgTagihanPasien.Col) = ""
        hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 34) = 0
    End If
    Call subHitungTotal

    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub chkCheck_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkCheck.Visible = False
        Call chkCheck_Click
        hgTagihanPasien.SetFocus
    End If
End Sub

Private Sub chkCheck_LostFocus()
    chkCheck.Visible = False
End Sub

Private Sub chkCheckJamsostek_Click()
    If chkCheckJamsostek.Value = vbChecked Then
        hgPaketKhususJamsostek.TextMatrix(hgPaketKhususJamsostek.Row, hgPaketKhususJamsostek.Col) = Chr$(187)
        hgPaketKhususJamsostek.TextMatrix(hgPaketKhususJamsostek.Row, 34) = 1
    Else
        hgPaketKhususJamsostek.TextMatrix(hgPaketKhususJamsostek.Row, hgPaketKhususJamsostek.Col) = ""
        hgPaketKhususJamsostek.TextMatrix(hgPaketKhususJamsostek.Row, 34) = 0
    End If
    Call subHitungTotalPembagianPaketKhusus
End Sub

Private Sub chkCheckJamsostek_LostFocus()
    chkCheckJamsostek.Visible = False
End Sub

Private Sub chkDetail_Click()
    If chkDetail.Value = 1 Then
        fraDetailRekap.Visible = True
    Else
        fraDetailRekap.Visible = False
    End If
End Sub

Private Sub chkDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdBayar.SetFocus
End Sub

Private Sub chkPaketJamsostek_Click()
    On Error GoTo hell
    Dim i As Integer
    Dim j As Integer
    Dim rsAsktek As New ADODB.recordset
    If chkPaketJamsostek.Value = Checked Then
        chkPaketJamsostek.Enabled = False
        dcPaketKhususJamsostek.BoundText = ""
        dcPaketKhususJamsostek.Text = ""
        chkBagiRata.Value = Unchecked
        chkDitanggungPenjamin.Value = Unchecked
        txtTarifTanggungan.Text = "0"
        txtTotalPembagian.Text = "0"
        curTotalBiayaDipilih = 0

        chkDetail.Enabled = False
        chkTagihanApotik.Enabled = False
        cmdBayar.Enabled = False
        cmdTutup.Enabled = False
        cmdPelayanan.Enabled = False

        fraPaketKhususJamsostek.Top = 1800
        fraPaketKhususJamsostek.Left = 120
        fraPaketKhususJamsostek.Visible = True
        Call subSetGridPaketKhususJamsostek
        Set rs = Nothing
        Call msubDcSource(dcPaketKhususJamsostek, rs, "Select KdPaket,NamaPaket From PaketAsuransi order by KdPaket")

        strSQL = ""
        strSQL = "Select * from V_RincianTotalDetailBiayaPelayanan WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "'" & sfilter
        Set rs = Nothing
        Call msubRecFO(rs, strSQL)

        If rs.RecordCount <> 0 Then
            hgPaketKhususJamsostek.Clear
            hgPaketKhususJamsostek.Rows = rs.RecordCount + 1
            For i = 1 To rs.RecordCount
                For j = 1 To 32
                    If j = 7 Or j = 21 Then
                        hgPaketKhususJamsostek.Row = i: hgPaketKhususJamsostek.Col = j: hgPaketKhususJamsostek.CellForeColor = vbBlue
                        If rs(j - 1).Value = 0 Then
                            hgPaketKhususJamsostek.TextMatrix(i, j) = "" & rs(j - 1).Value
                        Else
                            hgPaketKhususJamsostek.TextMatrix(i, j) = "" & Format(rs(j - 1).Value, "#,###")
                        End If
                    Else
                        If j = 31 Or j = 32 Then
                            If hgPaketKhususJamsostek.TextMatrix(i, 25) = "OA" Then
                                strSQL = ""
                                strSQL = "SELECT BiayaAdministrasi, TarifService, HargaSatuan From DetailPemakaianAlkes" & _
                                " WHERE (NoPendaftaran = '" & txtNoPendaftaran.Text & "' ) AND (KdRuangan = '" & hgPaketKhususJamsostek.TextMatrix(i, 20) & "') AND (KdBarang = '" & hgPaketKhususJamsostek.TextMatrix(i, 3) & "') AND (KdAsal = '" & hgPaketKhususJamsostek.TextMatrix(i, 17) & "') AND (TglPelayanan = '" & Format(hgPaketKhususJamsostek.TextMatrix(i, 8), "yyyy/MM/dd HH:mm:ss") & "') AND (SatuanJml = '" & hgPaketKhususJamsostek.TextMatrix(i, 30) & "')"
                                Set rsAsktek = Nothing
                                Call msubRecFO(rsAsktek, strSQL)
                                If rsAsktek.EOF = False Then
                                    hgPaketKhususJamsostek.TextMatrix(i, 31) = rsAsktek("BiayaAdministrasi")
                                    hgPaketKhususJamsostek.TextMatrix(i, 32) = rsAsktek("TarifService")
                                    hgPaketKhususJamsostek.TextMatrix(i, 33) = rsAsktek("HargaSatuan")
                                Else
                                    hgPaketKhususJamsostek.TextMatrix(i, 31) = "0"
                                    hgPaketKhususJamsostek.TextMatrix(i, 32) = "0"
                                    hgPaketKhususJamsostek.TextMatrix(i, 33) = hgPaketKhususJamsostek.TextMatrix(i, 6)
                                End If

                            Else
                                hgPaketKhususJamsostek.TextMatrix(i, 31) = "0"
                                hgPaketKhususJamsostek.TextMatrix(i, 32) = "0"
                                hgPaketKhususJamsostek.TextMatrix(i, 33) = hgPaketKhususJamsostek.TextMatrix(i, 6)
                            End If
                        Else
                            hgPaketKhususJamsostek.TextMatrix(i, j) = "" & rs(j - 1).Value
                        End If
                    End If
                    If j = 1 Then hgPaketKhususJamsostek.TextMatrix(i, j) = "" 'Chr$(187)
                    If j = 21 Then
                        If val(hgPaketKhususJamsostek.TextMatrix(i, 7)) <> val(hgPaketKhususJamsostek.TextMatrix(i, j)) Then hgPaketKhususJamsostek.CellForeColor = vbRed
                        If val(hgPaketKhususJamsostek.TextMatrix(i, j)) = 0 Then hgPaketKhususJamsostek.CellForeColor = vbBlack
                    End If
                Next j
                rs.MoveNext
            Next i

            For i = 1 To rs.RecordCount
                If hgPaketKhususJamsostek.TextMatrix(i, 1) <> Chr$(187) Then
                    hgPaketKhususJamsostek.TextMatrix(i, 34) = 0
                Else
                    hgPaketKhususJamsostek.TextMatrix(i, 34) = 1
                End If
            Next i
            Call setJudulPaketKhususJamsostek
        End If

    Else
        chkPaketJamsostek.Enabled = True
        fraPaketKhususJamsostek.Visible = False
    End If
    Exit Sub
hell:
    Call msubPesanError
    fraPaketKhususJamsostek.Visible = False
End Sub

Private Sub chkTagihanApotik_Click()
    Dim rsAP As New ADODB.recordset
    If chkTagihanApotik.Value = vbChecked Then
        sfilter = ""
        sfilter = " AND KdRuangan In ('702','703','704','705','706')"
        Call txtNoPendaftaran_KeyPress(13)
        dcRuanganApotik.Enabled = True
        Set rsAP = Nothing
        Call msubDcSource(dcRuanganApotik, rsAP, "Select KdRuangan,NamaRuangan From Ruangan Where KdInstalasi = '07' AND KdRuangan <> '701'")
    Else
        dcRuanganApotik.Enabled = False
        dcRuanganApotik.BoundText = ""
        dcRuanganApotik.Text = ""
        sfilter = ""
        Call txtNoPendaftaran_KeyPress(13)
    End If
End Sub

Private Sub cmdaktifsimpan_Click()
    cmdBayar.Enabled = True
End Sub

Private Sub cmdBayar_Click()
    Dim jDataCheck As Integer
    'Proses Pengecekan Jika ada transaksi tambahan
'    strSQL = "Select * from V_RincianTotalDetailBiayaPelayanan WHERE NoPendaftaran='" & mstrNoPen & "'"
'    Call msubRecFO(rs, strSQL)
'    jDataCheck = rs.RecordCount
'    If jData <> jDataCheck Then
'        BoolKembali = True
'    Else
'        BoolKembali = False
'    End If
'    If BoolKembali = True Then
'        MsgBox "ada tindakan yang baru diinput di ruangan, ulangi penyimpanan transaksi", vbInformation + vbOKOnly, "informas"
'        txtNoPendaftaran_KeyPress (13)
'        Exit Sub
'    End If
    If funcCekValidasi = False Then Exit Sub
    If subbolEditTanggungan = True Then
        If MsgBox("Anda telah merubah besar tanggungan pasien, " & vbNewLine & "pilih Yes untuk meneruskan transaksi", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
        If UpdateDetailBiayaPelayanan = False Then Exit Sub
        Exit Sub
    End If
    
    mstrJenisPasien = txtJenisPasien.Text

    blnValidasiInput = True    'agar validasi jika ada input baru saat bayar
    
    With frmBayar
        .Show
        Me.Enabled = False

        mcurBayar = FormatCurrency(CCur(lblTotalTagihan.Caption), 4)
        mcurAll_TBP = FormatCurrency(CCur(txtTotalBiaya.Text), 4)
        If txtTAsuransi.Text = "" Then
            mcurAll_TP = 0
        Else
            mcurAll_TP = CCur(txtTAsuransi.Text)
        End If
        mcurAll_TP = FormatCurrency(mcurAll_TP, 4)
        mcurAll_TRS = FormatCurrency(CCur(txtTRS.Text), 4)
        mcurAll_HrsDibyr = FormatCurrency(CCur(lblTotalTagihan.Caption), 4)
        mcurTM_Pemb = FormatCurrency(CCur(txtTMPembebasan.Text), 4)
        mcurOA_Pemb = FormatCurrency(CCur(txtOAPembebasan.Text), 4)

        .txtTotalBiaya.Text = txtTotalBiaya.Text
        .txtTAsuransi.Text = txtTAsuransi.Text
        .txtTRS.Text = txtTRS.Text
        .txtTPembebasan.Text = txtTotalPembebasan.Text
        .txtTotalBayar.Text = lblTotalTagihan.Caption
        .txtNamaFormPengirim.Text = Me.Name

        Set rs = Nothing
        strSQL = "SELECT MAX(PembayaranKe) AS PembayaranKe FROM dbo.PembayaranDepositBiayaPerawatan where nopendaftaran = '" & mstrNoPen & "'"
        Call msubRecFO(rs, strSQL)
        If Not IsNull(rs.Fields(0)) Then intPembayaranKe = rs.Fields(0)

        Set rs = Nothing
        'strSQL = "SELECT * from PembayaranDepositBiayaPerawatan where nopendaftaran = '" & mstrNoPen & "' AND PembayaranKe=" & intPembayaranKe & ""
         strSQL = "SELECT sum(JmlSudahDibayar) as JmlSudahDibayar, NoStruk from PembayaranDepositBiayaPerawatan where nopendaftaran = '" & mstrNoPen & "' Group by NoStruk"
       '  strSQL = "SELECT sum(JmlSudahDibayar) from PembayaranDepositBiayaPerawatan where nopendaftaran = '" & mstrNoPen & "' and nostruk is null "
        Call msubRecFO(rs, strSQL)
               
         If rs.RecordCount <> 0 Then
            If IsNull(rs.Fields(0)) Then
                 .txtDeposit.Text = 0
            Else
                
                If IsNull(rs.Fields(1)) Then
                  .txtDeposit.Text = rs("JmlSudahDibayar")
                Else
                
                  .txtDeposit.Text = 0
                End If
                
            End If
            
        Else
        
            .txtDeposit.Text = 0
        
        End If
         
        .txtDeposit.Text = FormatCurrency(.txtDeposit.Text, 4)
        .txtDiscount.Text = 0
        'Tambahan untuk mengkonversi 2 digit dibelakang Koma..(add by JDR)
        .txtDiscount.Text = FormatCurrency(.txtDiscount.Text, 4)
        .txtBiayaAdministrasi.Text = 0
        'Tambahan untuk mengkonversi 2 digit dibelakang Koma.(add by JDR)
        .txtBiayaAdministrasi.Text = FormatCurrency(.txtBiayaAdministrasi.Text, 4)

        .txtJmlUang.Text = lblTotalTagihan.Caption - CCur(.txtDeposit.Text)
        
        '.txtJmlUang = FormatCurrency(.txtJmlUang, 2)
        .txtJmlUang.Text = lblTotalTagihan.Caption - CCur(.txtDeposit.Text)
        .txtJmlUang = Format(.txtJmlUang, "#,###.0000")
        
        If .txtKembalian.Text = "" Then
            .txtKembalian.Text = FormatCurrency(0, 4)
        Else
            .txtKembalian.Text = FormatCurrency(.txtKembalian.Text, 4)
        End If

        If CCur(.txtDeposit.Text) > 0 Then
            .txtSisaTagihan.Text = FormatCurrency(CCur(.lblTotalTagihan.Caption) - (CCur(.txtJmlUang.Text) + CCur(.txtDeposit.Text)), 4)
        Else
            .txtSisaTagihan.Text = FormatCurrency(CCur(.lblTotalTagihan.Caption) - CCur(.txtJmlUang.Text), 4)
        End If
        If Left(CDec(.txtSisaTagihan.Text), 1) = "-" Then .txtSisaTagihan.Text = FormatCurrency(0, 4)

        If mcurAll_HrsDibyr = 0 Then
            .txtDiscount.Enabled = False
            .txtJmlUang.Enabled = False
        Else
            .txtDiscount.Enabled = True
            .txtJmlUang.Enabled = True
        End If
        mcurAll_HrsDibyr = FormatCurrency(mcurAll_HrsDibyr, 4)
        
        .txtJmlUang.Enabled = True
        .txtJmlUang.SetFocus
        .cmdSimpan.Enabled = False
        
    End With
End Sub

Private Sub cmdDiagnosa_Click()
On Error GoTo errLoad
    Dim subKdDokterTemp As String

'    Me.Enabled = False
    With frmPeriksaDiagnosa
        .Show
        .txtNamaFormPengirim.Text = Me.Name
        .txtNoPendaftaran = txtNoPendaftaran.Text
        .txtNoCM = txtNoCM.Text
        .txtNamaPasien = txtNamaPasien.Text
        .txtSex.Text = txtSex.Text

        .txtThn = txtThn.Text
        .txtBln = txtBln.Text
        .txtHari = txtHari.Text
        
        strsqlD = ""
        If KdInstalasiDiagnosa = "01" Then
            strsqlD = "SELECT KdSubInstalasi, IdDokter, NamaLengkap as DokterPemeriksa FROM RegistrasiIGD " & _
            " LEFT JOIN dbo.DataPegawai ON dbo.RegistrasiIGD.IdDokter = dbo.DataPegawai.IdPegawai WHERE (NoPendaftaran = '" & txtNoPendaftaran.Text & "') "
        ElseIf KdInstalasiDiagnosa = "02" Then
            strsqlD = "SELECT KdSubInstalasi, IdDokter, NamaLengkap as DokterPemeriksa FROM RegistrasiRJ " & _
            " LEFT JOIN dbo.DataPegawai ON dbo.RegistrasiRJ.IdDokter = dbo.DataPegawai.IdPegawai WHERE (NoPendaftaran = '" & txtNoPendaftaran.Text & "') "
        ElseIf KdInstalasiDiagnosa = "03" Then
            strsqlD = "SELECT KdSubInstalasi, IdDokter, NamaLengkap as DokterPemeriksa FROM RegistrasiRI " & _
            " LEFT JOIN dbo.DataPegawai ON dbo.RegistrasiRI.IdDokter = dbo.DataPegawai.IdPegawai WHERE (NoPendaftaran = '" & txtNoPendaftaran.Text & "') "
        End If
        
        Set rsD = Nothing
        Call msubRecFO(rsD, strsqlD)
        
        If rsD.EOF = False Then
            mstrKdSubInstalasi = IIf(IsNull(rsD("KdSubinstalasi")), "", rsD("KdSubinstalasi"))
            subKdDokterTemp = IIf(IsNull(rsD("IdDokter")), "", rsD("IdDokter"))
            .txtDokter = IIf(IsNull(rsD("DokterPemeriksa")), "", rsD("DokterPemeriksa"))
            mstrKdDokter = subKdDokterTemp
        End If
        .fraDokter.Visible = False
        
    End With
    Exit Sub
errLoad:
    Call msubPesanError
    frmPeriksaDiagnosa.Show
End Sub

Private Sub cmdHitung_Click()
    On Error GoTo hell
    Call subHitungTotalPembagianPaketKhusus
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdPelayanan_Click()
    If txtNoCM.Text = "" Then Exit Sub
    mstrNoPen = txtNoPendaftaran.Text

    mblnTindakanKasir = True
    frmPilihSubIns.Show
End Sub

Public Sub cmdPerbaikiData_Click()
    On Error GoTo hell_
    Dim i As Integer
    Dim sKdInstalasi As String
    Dim sKdRuanganAsal As String
    pbData.Value = pbData.Min
    pbData.Max = hgTagihanPasien.Rows - 1
    cmdBayar.Enabled = True
    pbData.Value = 1
    For i = 1 To hgTagihanPasien.Rows - 1
        pbData.Value = i
        DoEvents
        With hgTagihanPasien
            If .TextMatrix(i, 37) = "T" And .TextMatrix(i, 25) = "TM" Then
                Set rs = Nothing
                strSQL = "Select KdInstalasi from Ruangan where KdRuangan='" & .TextMatrix(i, 20) & "'"
                Call msubRecFO(rs, strSQL)
                sKdInstalasi = rs.Fields("KdInstalasi")

                Set rs = Nothing
                strSQL = "select dbo.FB_TakeRuanganAsal('" & mstrNoPen & "','" & .TextMatrix(i, 20) & "', null,'" & Format(.TextMatrix(i, 8), "yyyy/mm/dd HH:mm:ss") & "','TM') "
                Call msubRecFO(rs, strSQL)
                sKdRuanganAsal = rs.Fields(0)

                If sKdInstalasi = "09" Or sKdInstalasi = "10" Or sKdInstalasi = "16" Then
                    sqlQuery = "Delete TempHargaKomponen where NoPendaftaran ='" & mstrNoPen & "' AND KdRuangan ='" & .TextMatrix(i, 20) & "'  and TglPelayanan ='" & Format(.TextMatrix(i, 8), "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS = '" & Trim(.TextMatrix(i, 3)) & "' " ' "
                    Call msubRecFO(rsx, sqlQuery)
                    If frmValidasiData.Add_TempHargaKomponenForPenunjang(.TextMatrix(i, 20), .TextMatrix(i, 8), Trim(.TextMatrix(i, 3)), .TextMatrix(i, 15), .TextMatrix(i, 18), .TextMatrix(i, 29), .TextMatrix(i, 5), IIf(.TextMatrix(i, 29) = 0, 0, 1), "", sKdRuanganAsal) = False Then Exit Sub
                ElseIf sKdInstalasi = "04" Then
                    Set rs = Nothing
                    strSQL = "SELECT NoPendaftaran FROM  DokterPelaksanaOperasi" & _
                    " where NoPendaftaran = '" & mstrNoPen & "' "
                    Call msubRecFO(rs, strSQL)
                    If rs.EOF = False Then
                        sqlQuery = "Delete TempHargaKomponen where NoPendaftaran ='" & mstrNoPen & "' AND KdRuangan ='" & .TextMatrix(i, 20) & "'  and TglPelayanan ='" & Format(.TextMatrix(i, 8), "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS = '" & Trim(.TextMatrix(i, 3)) & "' " ' "
                        Call msubRecFO(rsx, sqlQuery)
                        If frmValidasiData.Add_TempHargaKomponenForIBS_DBNew(.TextMatrix(i, 20), .TextMatrix(i, 8), Left(.TextMatrix(i, 3), 6), .TextMatrix(i, 15), .TextMatrix(i, 18), .TextMatrix(i, 5), sKdRuanganAsal) = False Then Exit Sub
                    Else
                        sqlQuery = "Delete TempHargaKomponen where NoPendaftaran ='" & mstrNoPen & "' AND KdRuangan ='" & .TextMatrix(i, 20) & "'  and TglPelayanan ='" & Format(.TextMatrix(i, 8), "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS = '" & Trim(.TextMatrix(i, 3)) & "' " ' "
                        Call msubRecFO(rsx, sqlQuery)
                        If frmValidasiData.Add_TempHargaKomponenForIBSNew(.TextMatrix(i, 20), .TextMatrix(i, 8), Trim(.TextMatrix(i, 3)), .TextMatrix(i, 15), .TextMatrix(i, 18), .TextMatrix(i, 5), sKdRuanganAsal) = False Then Exit Sub
                    End If
                Else
                    sqlQuery = "Delete TempHargaKomponen where NoPendaftaran ='" & mstrNoPen & "' AND KdRuangan ='" & .TextMatrix(i, 20) & "'  and TglPelayanan ='" & Format(.TextMatrix(i, 8), "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS = '" & Trim(.TextMatrix(i, 3)) & "' " ' "
                    Call msubRecFO(rsx, sqlQuery)
                    
                    strSQL = "SELECT IdPegawai FROM  DataPegawai" & _
                    " where NamaLengkap = '" & .TextMatrix(i, 14) & "'"
                    Call msubRecFO(rs, strSQL)
                    idpegawai = rs.Fields("IdPegawai")
                    If frmValidasiData.Add_TempHargaKomponen(.TextMatrix(i, 20), .TextMatrix(i, 8), Trim(.TextMatrix(i, 3)), .TextMatrix(i, 15), .TextMatrix(i, 18), .TextMatrix(i, 29), .TextMatrix(i, 5), IIf(.TextMatrix(i, 29) = 0, 0, 1), idpegawai, sKdRuanganAsal) = False Then Exit Sub
                End If
                'untuk perbaikai data oa yang tidak valid
            ElseIf .TextMatrix(i, 37) = "T" And .TextMatrix(i, 25) = "OA" Then

'               sqlQuery = "Delete TempHargaKomponenObatAlkes where NoPendaftaran ='" & mstrNoPen & "' AND KdRuangan ='" & .TextMatrix(i, 20) & "'  and KdAsal ='" & .TextMatrix(i, 17) & "' and TglPelayanan ='" & Format(.TextMatrix(i, 8), "yyyy/MM/dd HH:mm:ss") & "' and KdBarang = '" & Trim(.TextMatrix(i, 3)) & "' and NoTerima='" & Trim(.TextMatrix(i, 40)) & "'"
                sqlQuery = "Delete TempHargaKomponenObatAlkes where NoPendaftaran ='" & mstrNoPen & "' AND KdRuangan ='" & .TextMatrix(i, 20) & "'  and KdAsal ='" & .TextMatrix(i, 17) & "' and TglPelayanan ='" & Format(.TextMatrix(i, 8), "yyyy/MM/dd HH:mm:ss") & "' and KdBarang = '" & Trim(.TextMatrix(i, 3)) & "' and NoTerima='" & Trim(.TextMatrix(i, 40)) & "' and ResepKe='" & .TextMatrix(i, 41) & "'"
                Call msubRecFO(rsx, sqlQuery)
'
'               strSQL = "select SatuanJml,JmlService,TarifService,NoResep,HargaBeli,KdJenisObat,BiayaAdministrasi,KdRuanganAsal from PemakaianAlkes where NoPendaftaran='" & mstrNoPen & "' and NoStruk is null"
                strSQL = "select SatuanJml,JmlService,TarifService,NoResep,HargaBeli,KdJenisObat,BiayaAdministrasi,KdRuanganAsal,resepke,KdKelas from PemakaianAlkes where NoPendaftaran='" & mstrNoPen & "' and NoStruk is null"
                               
                Call msubRecFO(rs, strSQL)
                Dim sSatunJml As String
                Dim iJmlServeice As Integer
                Dim cTarifServeice As Currency
                Dim sNoResep As String
                Dim cHargabeli As Currency
                Dim sKdJenisObat As String
                Dim cBiayaAdministrasi As Currency
                Dim sKdRuangAsal As String
                Dim iResepKe As Integer
                
                sSatunJml = IIf(IsNull(rs.Fields("SatuanJml")), "", rs.Fields("SatuanJml"))
                iJmlServeice = IIf(IsNull(rs.Fields("JmlService")), 0, rs.Fields("JmlService"))
                cTarifServeice = IIf(IsNull(rs.Fields("TarifService")), 0, rs.Fields("TarifService"))
                sNoResep = IIf(IsNull(rs.Fields("NoResep")), "", rs.Fields("NoResep"))
                cHargabeli = IIf(IsNull(rs.Fields("HargaBeli")), 0, rs.Fields("HargaBeli"))
                sKdJenisObat = IIf(IsNull(rs.Fields("KdJenisObat")), "", rs.Fields("KdJenisObat"))
                cBiayaAdministrasi = IIf(IsNull(rs.Fields("BiayaAdministrasi")), 0, rs.Fields("BiayaAdministrasi"))
                sKdRuangAsal = IIf(IsNull(rs.Fields("KdRuanganAsal")), "", rs.Fields("KdRuanganAsal"))
'                iResepKe = IIf(IsNull(rs.Fields("resepKe")), 0, rs.Fields("resepKe"))

                Dim sStatusDijamin As String
'                If frmValidasiData.Add_TempHargaKomponenOAResep(mstrNoPen, .TextMatrix(i, 20), .TextMatrix(i, 8), Trim(.TextMatrix(i, 3)), .TextMatrix(i, 17), .TextMatrix(i, 30), .TextMatrix(i, 6), cHargaBeli, .TextMatrix(i, 5), sKdJenisObat, iJmlServeice, cTarifServeice, sNoResep, cBiayaAdministrasi, sKdRuangAsal, sStatusDijamin, .TextMatrix(i, 40), .TextMatrix(i, 41)) = False Then Exit Sub
            If frmValidasiData.Add_TempHargaKomponenOAResep(mstrNoPen, .TextMatrix(i, 20), .TextMatrix(i, 8), Trim(.TextMatrix(i, 3)), _
                .TextMatrix(i, 17), .TextMatrix(i, 30), .TextMatrix(i, 6), cHargabeli, .TextMatrix(i, 5), sKdJenisObat, iJmlServeice, _
                cTarifServeice, sNoResep, cBiayaAdministrasi, sKdRuangAsal, sStatusDijamin, .TextMatrix(i, 40), .TextMatrix(i, 41)) = False Then Exit Sub
            
            End If
        End With
    Next i
    Call txtNoPendaftaran_KeyPress(13)
'    MsgBox "Validasi selesai  ", vbInformation, "Informasi"
    pbData.Value = pbData.Min

    cmdBayar.Enabled = True
    Exit Sub
hell_:
'    msubPesanError
'    Resume 0
End Sub

Private Sub cmdProses_Click()
    On Error GoTo hell
    Dim i As Integer

    subbolEditTanggungan = True
    For i = 1 To hgTagihanPasien.Rows - 1
        If hgPaketKhususJamsostek.TextMatrix(i, 34) = 1 Then
            hgTagihanPasien.TextMatrix(i, 21) = hgPaketKhususJamsostek.TextMatrix(i, 21)
        End If
    Next i

    Call cmdSelesai_Click
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdSelesai_Click()
    fraPaketKhususJamsostek.Visible = False
    chkPaketJamsostek.Enabled = True
    chkPaketJamsostek.Value = Unchecked

    chkDetail.Enabled = True
    chkTagihanApotik.Enabled = True
    cmdBayar.Enabled = True
    cmdTutup.Enabled = True
    cmdPelayanan.Enabled = True
End Sub

Private Sub cmdTambahData_Click()
  frmPelayananTindakandanObatAlkes.Show
'   frmTindakan.Show
End Sub

Private Sub cmdTotalTanggunganPenjamin_Click()
'If txtTotalPenjamin.Text = "" Or txtTotalPenjamin.Text = "0" Then Exit Sub
bolinacbgs = True
Call HitungClaimInaCbg
End Sub

Private Sub cmdTotalTanggunganPenjamins_Click()

End Sub

Private Sub cmdTotalTanggunganPenjamin2_Click()
On Error GoTo hell
If funcCekValidasi = False Then Exit Sub
Set rs3 = Nothing

If txtPenjamin <> "BPJS" Then MsgBox "Hitung Klaim InaCBG hanya untuk Pasien BPJS", vbInformation, "Informasi": Exit Sub

'strsqlD = ""
'If KdInstalasiDokter = "01" Then
'    strsqlD = "SELECT IdDokter FROM RegistrasiIGD WHERE (NoPendaftaran = '" & txtNoPendaftaran.Text & "') "
'ElseIf KdInstalasiDokter = "02" Then
'    strsqlD = "SELECT IdDokter FROM RegistrasiRJ WHERE (NoPendaftaran = '" & txtNoPendaftaran.Text & "') "
'ElseIf KdInstalasiDokter = "03" Then
'    strsqlD = "SELECT IdDokter FROM RegistrasiRI WHERE (NoPendaftaran = '" & txtNoPendaftaran.Text & "') "
'End If
'
'Set rsD = Nothing
'Call msubRecFO(rsD, strsqlD)
'
'If IsNull(rsD!IdDokter) Then
'    If MsgBox("Dokter pemeriksa kosong. Apakah akan diinput?", vbYesNo, "Konfirmasi") = vbYes Then Call cmdUbahDokter_Click
'    Exit Sub
'End If

strsqlx3 = ""
strsqlx3 = "SELECT KdDiagnosa FROM PeriksaDiagnosa WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "'"
Call msubRecFO(rs3, strsqlx3)
If rs3.EOF = True Then
   If MsgBox("Data diagnosa kosong. Apakah akan diinput?", vbYesNo, "Konfirmasi") = vbYes Then
       Call cmdDiagnosa_Click
   End If
   Exit Sub
End If
Call HitungClaimInaCbgNew
Exit Sub
hell:
Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    On Error Resume Next
    Unload Me
    If blnForm = False Then
        frmCariPasien.ChkPasienBelumPulang.Value = vbUnchecked
        frmCariPasien.cmdCari_Click
    End If
End Sub

Private Function UpdateDetailBiayaPelayanan() As Boolean
    On Error GoTo errLoad
    Dim i As Integer

    'u/ intern function
    UpdateDetailBiayaPelayanan = True

    'update ke detail biaya pelayanan beban yang ditanggung penjamin/rs
    For i = 1 To frmTagihanPasien.hgTagihanPasien.Rows - 1
        If hgTagihanPasien.TextMatrix(i, 35) <> "" And hgTagihanPasien.TextMatrix(i, 21) <> hgTagihanPasien.TextMatrix(i, 35) Then
            If CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 5)) <> 0 Then
                If sp_DetailBiayaPelayanan4PasienNU(mstrNoPen, _
                    frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 20), _
                    frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 3), _
                    CDate(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 8)), _
                    frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 17), _
                    ((CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 33)) / CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 7))) * CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 21))) + IIf(CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 21)) = 0, 0, CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 31))) + CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 32)), _
                    ((CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 33)) / CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 7))) * CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 22))) + IIf(CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 22)) = 0, 0, CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 31))) + CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 32)), _
                    ((CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 33)) / CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 7))) * CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 23))) + IIf(CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 23)) = 0, 0, CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 31))) + CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 32)), _
                    frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 25), frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 30)) = False Then Exit Function
                End If
            End If
            If hgTagihanPasien.TextMatrix(i, 36) <> "" And hgTagihanPasien.TextMatrix(i, 22) <> hgTagihanPasien.TextMatrix(i, 36) Then
                If CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 5)) <> 0 Then
                    If sp_DetailBiayaPelayanan4PasienNU(mstrNoPen, _
                        frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 20), _
                        frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 3), _
                        CDate(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 8)), _
                        frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 17), _
                        ((CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 33)) / CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 7))) * CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 21))) + IIf(CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 21)) = 0, 0, CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 31))) + CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 32)), _
                        ((CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 33)) / CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 7))) * CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 22))) + IIf(CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 22)) = 0, 0, CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 31))) + CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 32)), _
                        ((CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 33)) / CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 7))) * CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 23))) + IIf(CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 23)) = 0, 0, CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 31))) + CDbl(frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 32)), _
                        frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 25), frmTagihanPasien.hgTagihanPasien.TextMatrix(i, 30)) = False Then Exit Function
                    End If
                End If
            Next i

            'utk validasi stlh Update_DetailBiayaPelayanan4PasienNU
            bolValUpdate = True

            Call txtNoPendaftaran_KeyPress(13)
            'cmdBayar.SetFocus

            'tanggungan sudah diedit
            subbolEditTanggungan = False
            cmdBayar.Enabled = True
            cmdBayar.SetFocus
            Exit Function
errLoad:
            UpdateDetailBiayaPelayanan = False
End Function

Private Sub cmdTutupUpdate_Click()
    picUpdateKomponen.Visible = False
    If cmdUpdate.Enabled = False Then
        hgTagihanPasien.TextMatrix(txtBarisKe.Text, 21) = CCur(txtTHutangPenjaminUpdate.Text)
        hgTagihanPasien.TextMatrix(txtBarisKe.Text, 22) = CCur(txtTTanggunganRSUpdate.Text)
        hgTagihanPasien.TextMatrix(txtBarisKe.Text, 23) = CCur(txtTPembebasanUpdate.Text)
    End If
    hgTagihanPasien.SetFocus
End Sub

Private Sub cmdUpdate_Click()
    Dim strKomponen12 As String
    On Error GoTo errLoad

    Dim i As Integer
    For i = 1 To fgUpdateKomponen.Rows
        If fgUpdateKomponen.TextMatrix(i, 0) = "" Then Exit For
        If sp_Update_TempHargaKomponen4PasienNU(txtNoPendaftaran.Text, txtKdRuanganPelayananUpdate.Text, txtKdPelayananUpdate.Text, txtTglPelayananUpdate.Text, fgUpdateKomponen.TextMatrix(i, 6), CCur(fgUpdateKomponen.TextMatrix(i, 3)), CCur(fgUpdateKomponen.TextMatrix(i, 4)), CCur(fgUpdateKomponen.TextMatrix(i, 5))) = False Then Exit Sub
    Next i
    subbolEditTanggungan = True
    strKomponen12 = "update tempHargaKomponen set JmlHutangPenjamin = " & msubKonversiKomaTitik(CCur(txtTHutangPenjaminUpdate)) & ",  JmlTanggunganRS = " & msubKonversiKomaTitik(CCur(txtTTanggunganRSUpdate)) & ", JmlPembebasan = " & msubKonversiKomaTitik(CCur(txtTPembebasanUpdate)) & " where NoPendaftaran = '" & Trim(txtNoPendaftaran.Text) & "' and year(TglPelayanan) = '" & Year(txtTglPelayananUpdate.Text) & "'  and month(TglPelayanan) = '" & Month(txtTglPelayananUpdate.Text) & "'  and day(TglPelayanan) = '" & Day(txtTglPelayananUpdate.Text) & "'  and datepart(hh,TglPelayanan) = '" & Hour(txtTglPelayananUpdate.Text) & "'  and datepart(mi,TglPelayanan) = '" & Minute(txtTglPelayananUpdate.Text) & "'  and datepart(ss,TglPelayanan) = '" & Second(txtTglPelayananUpdate.Text) & "' and KdPelayananRS = '" & Trim(txtKdPelayananUpdate.Text) & "' and KdKomponen = '12'"
    dbConn.Execute strKomponen12
    strKomponen12 = "update detailBiayaPelayanan set JmlHutangPenjamin = " & msubKonversiKomaTitik(CCur(txtTHutangPenjaminUpdate)) & ",  JmlTanggunganRS = " & msubKonversiKomaTitik(CCur(txtTTanggunganRSUpdate)) & ", JmlPembebasan = " & msubKonversiKomaTitik(CCur(txtTPembebasanUpdate)) & " where NoPendaftaran = '" & Trim(txtNoPendaftaran.Text) & "' and year(TglPelayanan) = '" & Year(txtTglPelayananUpdate.Text) & "'  and month(TglPelayanan) = '" & Month(txtTglPelayananUpdate.Text) & "'  and day(TglPelayanan) = '" & Day(txtTglPelayananUpdate.Text) & "'  and datepart(hh,TglPelayanan) = '" & Hour(txtTglPelayananUpdate.Text) & "'  and datepart(mi,TglPelayanan) = '" & Minute(txtTglPelayananUpdate.Text) & "'  and datepart(ss,TglPelayanan) = '" & Second(txtTglPelayananUpdate.Text) & "' and KdPelayananRS = '" & Trim(txtKdPelayananUpdate.Text) & "'"
    dbConn.Execute strKomponen12

    cmdUpdate.Enabled = False

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdValidasi_Click()
    frmValidasiData.Show
End Sub

Private Sub dcPaketKhususJamsostek_Change()
    On Error GoTo hell
    Dim rsCurrTP As New ADODB.recordset
    Dim strKdKelas As String

    Set rsCurrTP = Nothing
    rsCurrTP.Open "Select KdKelas From BiayaPelayanan Where NoPendaftaran='" & mstrNoPen & "'", dbConn, adOpenForwardOnly, adLockReadOnly
    If rsCurrTP.EOF = True Then strKdKelas = "" Else strKdKelas = rsCurrTP(0)
    strSQL = ""
    strSQL = "select * from V_M_TanggunganPaketAsuransi WHERE KdPaket = '" & dcPaketKhususJamsostek.BoundText & "' AND KdKelompokPasien = '" & mstrKdJenisPasien & "' AND IdPenjamin = '" & mstrKdPenjamin & "' AND KdKelas='" & strKdKelas & "'"
    Set rsCurrTP = Nothing
    rsCurrTP.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rsCurrTP.EOF = True Then Exit Sub
    txtTarifTanggungan.Text = Format(rsCurrTP("JmlTanggungan"), "####,##")
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcPaketKhususJamsostek_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        strSQL = ""
        strSQL = "select KdPaket, NamaPaket from PaketAsuransi WHERE (NamaPaket LIKE '%" & dcPaketKhususJamsostek.Text & "%')"
        Set rs = Nothing
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcPaketKhususJamsostek.BoundText = rs(0).Value
        dcPaketKhususJamsostek.Text = rs(1).Value

        hgPaketKhususJamsostek.SetFocus
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcRuanganApotik_Change()
    sfilter = " AND KdRuangan Like '%" & dcRuanganApotik.BoundText & "%'"
    Call txtNoPendaftaran_KeyPress(13)
End Sub

Private Sub dgTagihanPasien_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgTagihanPasien
    WheelHook.WheelHook dgTagihanPasien
End Sub

Private Sub fgUpdateKomponen_DblClick()
    Call fgUpdateKomponen_KeyPress(13)
End Sub

Private Sub fgUpdateKomponen_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    If KeyAscii = 13 Then
        Select Case fgUpdateKomponen.Col
            Case 3, 4, 5
                If fgUpdateKomponen.TextMatrix(fgUpdateKomponen.Row, 0) = "" Then Exit Sub
                Call subLoadTextUpdate
                txtIsiUpdate.Text = Trim(fgUpdateKomponen.TextMatrix(fgUpdateKomponen.Row, fgUpdateKomponen.Col))
                txtIsiUpdate.SelStart = 0: txtIsiUpdate.SelLength = Len(txtIsiUpdate.Text)
        End Select
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errLoad

    Select Case KeyCode
        Case vbKeyF1
            If txtNoPendaftaran.Text = "" Then Exit Sub
            If hgTagihanPasien.Rows = 2 And hgTagihanPasien.TextMatrix(1, 4) = "" Then Exit Sub
            mstrNoPen = txtNoPendaftaran.Text
            strjudul = Me.Name
            frm_cetak_RincianBiaya.Show

'        Case vbKeyF5
'            If mblnAdmin = False Then Exit Sub
'            If hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 3) = "" Then Exit Sub
'            strSQL = "SELECT * " & _
'            " FROM V_UbahBiayaPelayanan" & _
'            " WHERE NoPendaftaran = '" & txtNoPendaftaran.Text & "' AND (KdPelayananRS = '" & Trim(hgTagihanPasien.TextMatrix(frmTagihanPasien.hgTagihanPasien.Row, 3)) & "')AND (TglPelayanan = '" & Format(hgTagihanPasien.TextMatrix(frmTagihanPasien.hgTagihanPasien.Row, 8), "yyyy/MM/dd HH:mm:ss") & "')AND(KdRuangan = '" & hgTagihanPasien.TextMatrix(frmTagihanPasien.hgTagihanPasien.Row, 20) & "')"
'            Call msubRecFO(rs, strSQL)
'            If rs.EOF Then Exit Sub
'
'            Me.Enabled = False
'
'            With frmUpdateBiayaPelayanan
'                .txtNoPendaftaran = txtNoPendaftaran.Text
'                Call .txtNoPendaftaran_KeyPress(13)
'            End With
'
'        Case vbKeyF6
'            If txtNoCM.Text = "" Then Exit Sub
'            mblnTindakanKasir = True
'            frmPilihSubIns.Show

        Case vbKeyDelete
            If mblnAdmin = False Then Exit Sub
            If hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 7) = 0 Then
                If MsgBox("Apakah anda yakin akan menghapus pelayanan '" _
                & hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 4) & "'" & vbNewLine _
                & "Dengan tanggal pelayanan '" & hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 8) _
                & "'", vbQuestion + vbYesNo) = vbNo Then Exit Sub

                If sp_DeletePelayanan(hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 20), hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 3), hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 8)) = False Then Exit Sub
                MsgBox "Data berhasil dihapus", vbInformation
                Call txtNoPendaftaran_KeyPress(13)
                chkCheck.Visible = False
            End If
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    
    Set rs = Nothing
    strSQL = "SELECT DataPegawai.IdPegawai, DataPegawai.NamaLengkap " & _
             "FROM DataPegawai INNER JOIN DataCurrentPegawai ON DataPegawai.IdPegawai = DataCurrentPegawai.IdPegawai " & _
             "WHERE (DataPegawai.KdJenisPegawai = '001') AND (DataPegawai.NamaLengkap <> '') AND (DataCurrentPegawai.KdStatus = '01') " & _
             "ORDER BY DataPegawai.NamaLengkap"
    Call msubDcSource(dcDokter, rs, strSQL)
    
    txtNoPendaftaran.Text = Right(Year(Now), 2) & Format(Month(Now), "00") & Format(Day(Now), "00")
    txtNoPendaftaran.SelStart = Len(txtNoPendaftaran.Text)

    StatusBar1.Panels.Item(3).Visible = False 'Tambah pelayanan
    StatusBar1.Panels.Item(5).Visible = False 'Edit Tanggungan
    If mblnAdmin = True Then
        StatusBar1.Panels.Item(2).Visible = False 'Ubah pelayanan
        StatusBar1.Panels.Item(4).Visible = False 'Hapus pelayanan
    Else
        StatusBar1.Panels.Item(2).Visible = True 'Ubah pelayanan
        StatusBar1.Panels.Item(4).Visible = True 'Hapus pelayanan
    End If
    fraDetailRekap.Top = 4080
    fraDetailRekap.Left = 0
    Call setClearGridTagihan

    'utk validasi stlh Update_DetailBiayaPelayanan4PasienNU
    bolValUpdate = False
    bolinacbgs = False
    
    Call inisialisasiObjekContext
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Public Function sp_PostingHutangPenjaminPasien_AU(f_NoPendaftaran As String, f_status As String) As Boolean
    On Error GoTo hell
    sp_PostingHutangPenjaminPasien_AU = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, adInteger, Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "PostingHutangPenjaminPasien_AU"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_PostingHutangPenjaminPasien_AU = False
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Exit Function
hell:
    sp_PostingHutangPenjaminPasien_AU = False
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Call msubPesanError
End Function

'posting hutang penjamin/ TRS = 0 lg
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mstrNoPen <> "" Then
        If mstrKdPenjaminPasien <> "2222222222" Then
            Screen.MousePointer = vbHourglass
            If sp_PostingHutangPenjaminPasien_AU(mstrNoPen, "U") = False Then Exit Sub
            Screen.MousePointer = vbDefault
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If blnForm = False Then
        frmCariPasien.Enabled = True
    End If
    bolinacbgs = False
End Sub

Private Sub hgPaketKhususJamsostek_DblClick()
    On Error GoTo hell
    If hgPaketKhususJamsostek.Rows = 1 Then Exit Sub
    If hgPaketKhususJamsostek.TextMatrix(hgPaketKhususJamsostek.Row, 3) = "" Then Exit Sub

    Select Case hgPaketKhususJamsostek.Col
        Case 1
            If mblnAdmin = False Then Exit Sub

            If CCur(txtTarifTanggungan.Text) = 0 Then
                MsgBox "Nama Paket kosong", vbExclamation, "Validasi"
                chkBagiRata.Value = Unchecked
                Exit Sub
            End If
            If txtTotalBiaya.Text = "" Then
                MsgBox "Total Biaya Pelayanan kosong, Hubungi administrator", vbExclamation, "Validasi"
                chkBagiRata.Value = Unchecked
                Exit Sub
            End If

            chkCheckJamsostek.Visible = True
            chkCheckJamsostek.Top = hgPaketKhususJamsostek.RowPos(hgPaketKhususJamsostek.Row) + 1350
            Dim intA As Integer
            intA = ((hgPaketKhususJamsostek.ColPos(hgPaketKhususJamsostek.Col + 1) - hgPaketKhususJamsostek.ColPos(hgPaketKhususJamsostek.Col)) / 2)
            chkCheckJamsostek.Left = hgPaketKhususJamsostek.ColPos(hgPaketKhususJamsostek.Col) + 40 + intA
            chkCheckJamsostek.SetFocus
            If hgPaketKhususJamsostek.Col = 1 Then
                If hgPaketKhususJamsostek.TextMatrix(hgPaketKhususJamsostek.Row, 1) <> "" Then
                    chkCheckJamsostek.Value = 1
                Else
                    chkCheckJamsostek.Value = 0
                End If
            End If

    End Select
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub hgTagihanPasien_DblClick()
    On Error GoTo errLoad

    If hgTagihanPasien.Rows = 1 Then Exit Sub
    If hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 3) = "" Then Exit Sub
    chkCheck.Visible = False

    Select Case hgTagihanPasien.Col
        Case 1
            Set rs = Nothing
            strSQL = "SELECT KdKategoryUser FROM Login WHERE IdPegawai = '" & strIDPegawai & "'"
            Call msubRecFO(rs, strSQL)

            If rs.Fields(0).Value <> "01" And rs.Fields(0).Value <> "03" Then
                MsgBox "Anda tidak punya Akses untuk mengedit", vbCritical
                chkCheck.Visible = False
                Exit Sub
            End If

            chkCheck.Visible = True
            chkCheck.Top = hgTagihanPasien.RowPos(hgTagihanPasien.Row) + 390
            Dim intA As Integer
            intA = ((hgTagihanPasien.ColPos(hgTagihanPasien.Col + 1) - hgTagihanPasien.ColPos(hgTagihanPasien.Col)) / 2)
            chkCheck.Left = hgTagihanPasien.ColPos(hgTagihanPasien.Col) + 50 + intA '160 + intA
            chkCheck.SetFocus
            If hgTagihanPasien.Col = 1 Then
                If hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 1) <> "" Then
                    chkCheck.Value = 1
                Else
                    chkCheck.Value = 0
                End If
            End If

        Case 21, 22, 23
            If mblnAdmin = False Or mstrKdJenisPasien = "01" Then Exit Sub
            subbolEditTanggungan = True
            Call subLoadText
            txtIsi.Text = Trim(hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, hgTagihanPasien.Col))
            txtIsi.SelStart = 0: txtIsi.SelLength = Len(txtIsi.Text)
            If hgTagihanPasien.Col = 21 Then
                hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 35) = Trim(hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, hgTagihanPasien.Col))
            ElseIf hgTagihanPasien.Col = 22 Then
                hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 36) = Trim(hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, hgTagihanPasien.Col))
            End If
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub hgTagihanPasien_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            If hgTagihanPasien.Col = 0 Or hgTagihanPasien.Col = 21 Or hgTagihanPasien.Col = 22 Then Call hgTagihanPasien_DblClick
    End Select
End Sub

Private Sub hgTagihanPasien_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errLoad
    If hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 2) = "" Then Exit Sub
    If Button = vbLeftButton Then Exit Sub
    If strIDPegawaiAktif <> "8888888888" Then Exit Sub
    PopupMenu MDIUtama.MEditKomponen
    Exit Sub
errLoad:
    Call msubPesanError
End Sub


Private Sub Timer1_Timer()
    Static i As Integer
    pbPosting.Value = i
    i = i + 1
    If i = pbPosting.Max Then
        Timer1.Enabled = False
        fraPosting.Visible = False
        i = 0
        Exit Sub
    End If
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If KeyAscii = 13 Then
        hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, hgTagihanPasien.Col) = val(txtIsi.Text)
        txtIsi.Visible = False
        Call subHitungTotal

        If hgTagihanPasien.RowPos(hgTagihanPasien.Row) >= hgTagihanPasien.Height - 360 Then
            hgTagihanPasien.SetFocus
            SendKeys "{DOWN}"
            Exit Sub
        End If

        If hgTagihanPasien.Col = 2 Or hgTagihanPasien.Col = 3 Or hgTagihanPasien.Col = 4 Then
            If hgTagihanPasien.Row = hgTagihanPasien.Rows - 1 Then
                If hgTagihanPasien.TextMatrix(hgTagihanPasien.Rows - 1, 2) <> "" And hgTagihanPasien.TextMatrix(hgTagihanPasien.Rows - 1, 3) <> "" And hgTagihanPasien.TextMatrix(hgTagihanPasien.Rows - 1, 4) <> "" Then
                    hgTagihanPasien.Rows = hgTagihanPasien.Rows + 1
                End If
            End If
        End If

        hgTagihanPasien.SetFocus
    ElseIf KeyAscii = 27 Then
        txtIsi.Visible = False
    End If
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

Private Sub txtIsiUpdate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then txtIsiUpdate.Visible = False: fgUpdateKomponen.SetFocus
End Sub

Private Sub txtIsiUpdate_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If KeyAscii = 13 Then
        If fgUpdateKomponen.TextMatrix(fgUpdateKomponen.Row, 0) = "" Then Exit Sub
        fgUpdateKomponen.TextMatrix(fgUpdateKomponen.Row, fgUpdateKomponen.Col) = val(txtIsiUpdate.Text)
        If val(txtIsiUpdate.Text) > CCur(fgUpdateKomponen.TextMatrix(fgUpdateKomponen.Row, 2)) Then
            If MsgBox("Jumlah tanggungan lebih besar dari tarif pelayanan" & vbNewLine & "Yes untuk meneruskan, No untuk batal", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then
                fgUpdateKomponen.TextMatrix(fgUpdateKomponen.Row, fgUpdateKomponen.Col) = 0
                Exit Sub
            End If
        End If
        Call subHitungTotalUpdateKomponen
        txtIsiUpdate.Visible = False

        If fgUpdateKomponen.RowPos(fgUpdateKomponen.Row) >= fgUpdateKomponen.Height - 360 Then
            fgUpdateKomponen.SetFocus
            SendKeys "{DOWN}"
            Exit Sub
        End If

        If fgUpdateKomponen.Col = 2 Or fgUpdateKomponen.Col = 3 Or fgUpdateKomponen.Col = 4 Then
            If fgUpdateKomponen.Row = fgUpdateKomponen.Rows - 1 Then
                If fgUpdateKomponen.TextMatrix(fgUpdateKomponen.Rows - 1, 2) <> "" And fgUpdateKomponen.TextMatrix(fgUpdateKomponen.Rows - 1, 3) <> "" And fgUpdateKomponen.TextMatrix(fgUpdateKomponen.Rows - 1, 4) <> "" Then
                    fgUpdateKomponen.Rows = fgUpdateKomponen.Rows + 1
                End If
            End If
        End If
        fgUpdateKomponen.SetFocus

    ElseIf KeyAscii = 27 Then
        txtIsi.Visible = False
    End If
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtIsiUpdate_LostFocus()
    txtIsiUpdate.Visible = False
End Sub

Public Sub txtNoPendaftaran_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    Dim i As Integer
    Dim j As Integer

    If statusTagihan <> "OA" Then
            If KeyAscii = 13 Then
                bolDatavalid = False
                Call subClearData
                strSQL = "SELECT * FROM RegistrasiIGD WHERE NoPendaftaran = '" & txtNoPendaftaran.Text & "' AND StatusPulang = 'T'"
                Call msubRecFO(rs, strSQL)
                If rs.RecordCount <> 0 Then
                    MsgBox "Pasien belum keluar dari IGD", vbCritical
                    Exit Sub
                End If
        
                strSQL = "SELECT * FROM v_PasienAktifPakaiKamar WHERE NoPendaftaran='" _
                & txtNoPendaftaran.Text & "'"
                Set rs = Nothing
                rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
                If rs.RecordCount <> 0 Then
                    MsgBox "Pasien belum keluar dari Rawat Inap", vbCritical
                    Exit Sub
                End If
        
                strSQL = "SELECT NoCM FROM V_DaftarPasienSudahBayar WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "'"
                Set rs = Nothing:
                rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
                If rs.RecordCount <> 0 Then
                    mstrNoCMKu = rs.Fields(0).Value
                    strSQL = "SELECT NoPendaftaran FROM PasienBelumBayar WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "'"
                    Set rs = Nothing
                    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
                    If rs.RecordCount = 0 Then
                        strSQL = "INSERT INTO PasienBelumBayar VALUES ('" & txtNoPendaftaran.Text & "','" & mstrNoCMKu & "')"
                        dbConn.Execute strSQL
                    End If
                End If
        
                strSQL = "Select * from V_DaftarPasienBelumBayar_new WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "'"
                
                Set rs = Nothing
                rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
                If rs.RecordCount <> 0 Then
                    strSQL = "SELECT * FROM PasienDaftar WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "'"
                    Call msubRecFO(rsB, strSQL)
        
                    strKdKelPsn = rsB("KdKelompokPasien").Value
                    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & txtNoPendaftaran.Text & "')"
                    Call msubRecFO(rsB, strSQL)
                    If rsB.EOF = False Then
                        mstrKdJenisPasien = rsB("KdKelompokPasien").Value
                        mstrKdPenjaminPasien = IIf(IsNull(rsB("IdPenjamin")), "2222222222", rsB("IdPenjamin"))
                    End If
        
                    If mstrKdPenjaminPasien <> "2222222222" Then
                        strSQL = "SELECT * FROM PemakaianAsuransi WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "'"
                        Call msubRecFO(rsB, strSQL)
                        If rsB.RecordCount = 0 Then
                            MsgBox "Lengkapi dahulu data penjamin pasien", vbCritical, "Validasi"
                            Call subLoadFormJP
                            Exit Sub
                        End If
                    End If
        
                    mstrNoPen = txtNoPendaftaran.Text
                    txtNoCM.Text = rs.Fields(1).Value
                    mstrNoCM = txtNoCM.Text
                    txtNamaPasien.Text = rs.Fields(2).Value
                    txtSex.Text = IIf(rs.Fields(3).Value = "P", "Perempuan", "Laki-Laki")
                    txtThn.Text = rs("UmurTahun")
                    txtBln.Text = rs("UmurBulan")
                    txtHari.Text = rs("UmurHari")
                    txtJenisPasien.Text = rs.Fields(5).Value
        
                    If mstrKdPenjaminPasien <> "2222222222" Then
                        strSQL = "Select * from V_PerusahaanPenjaminPasien WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "'"
                        Call msubRecFO(rs, strSQL)
                        txtPenjamin.Text = rs.Fields(1).Value
                        mstrKdPenjamin = rs.Fields(2).Value
                    Else
                        txtPenjamin.Text = "Bayar Sendiri"
                        mstrKdPenjamin = mstrKdPenjaminPasien
                    End If
        
                    cmdBayar.Enabled = False
        
                    'utk looping add hutang penjamin dan tanggungan RS, yg sebelumnya diset 0 (blm diketahui)
                    If bolValUpdate = False Then
                        If mstrKdPenjaminPasien <> "2222222222" Then
                            fraPosting.Visible = True
                            Screen.MousePointer = vbHourglass
                            Me.Timer1.Enabled = True
                            If sp_PostingHutangPenjaminPasien_AU(mstrNoPen, "A") = False Then Exit Sub
                            Screen.MousePointer = vbDefault
                        End If
                    End If
        
                    strSQL = "Select * from V_RincianTotalDetailBiayaPelayanan WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "' and NoStruk Is Null" & sfilter & " ORDER BY TglPelayanan"
                    Call msubRecFO(rs, strSQL)
                    intJmlPelayanan = rs.RecordCount
                    jData = intJmlPelayanan
                    If rs.RecordCount <> 0 Then
                        hgTagihanPasien.Clear
                        hgTagihanPasien.Rows = rs.RecordCount + 1
                        For i = 1 To rs.RecordCount
                            For j = 1 To 32
                                If j = 7 Or j = 21 Then
                                    hgTagihanPasien.Row = i: hgTagihanPasien.Col = j: hgTagihanPasien.CellForeColor = vbBlue
                                    If rs(j - 1).Value = 0 Then
                                        hgTagihanPasien.TextMatrix(i, j) = "" & rs(j - 1).Value
                                    Else
                                        hgTagihanPasien.TextMatrix(i, j) = "" & rs(j - 1).Value
                                    End If
                                Else
                                    If j = 31 Or j = 32 Then
                                        If hgTagihanPasien.TextMatrix(i, 25) = "OA" Then
                                            strSQL = "SELECT BiayaAdministrasi, TarifService, HargaSatuan From DetailPemakaianAlkes" & _
                                            " WHERE (NoPendaftaran = '" & txtNoPendaftaran.Text & "' ) AND (KdRuangan = '" & hgTagihanPasien.TextMatrix(i, 20) & "') AND (KdBarang = '" & hgTagihanPasien.TextMatrix(i, 3) & "') AND (KdAsal = '" & hgTagihanPasien.TextMatrix(i, 17) & "') AND (TglPelayanan = '" & Format(hgTagihanPasien.TextMatrix(i, 8), "yyyy/MM/dd HH:mm:ss") & "') AND (SatuanJml = '" & hgTagihanPasien.TextMatrix(i, 30) & "')"
                                            Call msubRecFO(rsB, strSQL)
                                            If rsB.EOF = False Then
                                                hgTagihanPasien.TextMatrix(i, 31) = rsB("BiayaAdministrasi")
                                                hgTagihanPasien.TextMatrix(i, 32) = rsB("TarifService")
                                                hgTagihanPasien.TextMatrix(i, 33) = rsB("HargaSatuan")
                                            Else
                                                hgTagihanPasien.TextMatrix(i, 31) = "0"
                                                hgTagihanPasien.TextMatrix(i, 32) = "0"
                                                hgTagihanPasien.TextMatrix(i, 33) = hgTagihanPasien.TextMatrix(i, 6)
                                            End If
                                        Else
                                            hgTagihanPasien.TextMatrix(i, 31) = "0"
                                            hgTagihanPasien.TextMatrix(i, 32) = "0"
                                            hgTagihanPasien.TextMatrix(i, 33) = hgTagihanPasien.TextMatrix(i, 6)
                                        End If
                                    Else
                                        hgTagihanPasien.TextMatrix(i, j) = "" & rs(j - 1).Value
                                    End If
                                End If
                                If j = 1 Then hgTagihanPasien.TextMatrix(i, j) = Chr$(187)
                                If j = 21 Then
                                    If val(hgTagihanPasien.TextMatrix(i, 7)) <> val(hgTagihanPasien.TextMatrix(i, j)) Then hgTagihanPasien.CellForeColor = vbRed
                                    If val(hgTagihanPasien.TextMatrix(i, j)) = 0 Then hgTagihanPasien.CellForeColor = vbBlack
                                End If
                            Next j
                            ' ' untuk status verifikasi
                            
                   
                               curnilaiproposional = 0
                            If bolinacbgs = True Then
                                curnilaiproposional = (txtTotalPenjamin.Text / curtotalbiaya) * hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 7)
                                 hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 35) = Trim(hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, hgTagihanPasien.Col))
                                hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 21) = curnilaiproposional
                                     
                            End If
                            
                            
                            hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 38) = rs.Fields("TotalHutangPenjamin") 'rs.Fields("JmlHutangPenjamin")
                            hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 39) = rs.Fields("TotalTanggunganRS") 'rs.Fields("JmlTanggunganRS")
                            hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 40) = rs.Fields("NoTerima")
                            hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 41) = rs.Fields("No_Item")
                            rs.MoveNext
                            ' untuk status verifikasi
                            With hgTagihanPasien
                                If .TextMatrix(i, 25) = "TM" Then
                                    Set rsQuery = Nothing
        '                           sqlQuery = "SELECT dbo.FB_TakeStatusDataValid('" & txtNoPendaftaran.Text & "','" & .TextMatrix(i, 20) & "','" & Trim(.TextMatrix(i, 3)) & "', '" & Format(.TextMatrix(i, 8), "yyyy/mm/dd HH:mm:ss") & "', " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 6))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 29))) & " , " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 38))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 39))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 23))) & " )  as statusData"
                                    
                                    sqlQuery = "SELECT dbo.FB_TakeStatusDataValidOA('" & txtNoPendaftaran.Text & "','" & .TextMatrix(i, 20) & "','" & Trim(.TextMatrix(i, 3)) & "', '" & Format(.TextMatrix(i, 8), "yyyy/mm/dd HH:mm:ss") & "', " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 6))) & " , " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 38))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 39))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 23))) & ",'" & .TextMatrix(i, 40) & "', '" & .TextMatrix(i, 41) & "')  as statusData"
                                    Call msubRecFO(rsQuery, sqlQuery)
                                    strStatusData = rsQuery.Fields(0).Value
                                    .TextMatrix(i, 37) = strStatusData
                                ElseIf .TextMatrix(i, 25) = "OA" Then
                                    Set rsQuery = Nothing
        '                           sqlQuery = "SELECT dbo.FB_TakeStatusDataValidOA('" & txtNoPendaftaran.Text & "','" & .TextMatrix(i, 20) & "','" & Trim(.TextMatrix(i, 3)) & "', '" & Format(.TextMatrix(i, 8), "yyyy/mm/dd HH:mm:ss") & "', " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 6))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 38))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 39))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 23))) & ",'" & .TextMatrix(i, 40) & "')  as statusData"
                                    sqlQuery = "SELECT dbo.FB_TakeStatusDataValidOA('" & txtNoPendaftaran.Text & "','" & .TextMatrix(i, 20) & "','" & Trim(.TextMatrix(i, 3)) & "', '" & Format(.TextMatrix(i, 8), "yyyy/mm/dd HH:mm:ss") & "', " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 6))) & " , " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 38))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 39))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 23))) & ",'" & .TextMatrix(i, 40) & "', '" & .TextMatrix(i, 41) & "')  as statusData"
                                    Call msubRecFO(rsQuery, sqlQuery)
                                    strStatusData = rsQuery.Fields(0).Value
                                    .TextMatrix(i, 37) = strStatusData
                                End If
                            End With
                        Next i
        
                        ' untuk status verifikasi
                        For i = 1 To rs.RecordCount
                            If hgTagihanPasien.TextMatrix(i, 37) = "T" Then
                                Dim k As Integer
                                With hgTagihanPasien
                                    .Row = i
                                    For k = 0 To .Cols - 1
                                        .Col = k
        '                                .CellBackColor = vbRed
        '                                .CellForeColor = vbWhite
                                        bolDatavalid = True
                                    Next
                                End With
                                cmdBayar.Enabled = False
                            End If
                        Next i
        
                        For i = 1 To rs.RecordCount
                            If hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, hgTagihanPasien.Col) = Chr$(187) Then
                                hgTagihanPasien.TextMatrix(i, 34) = 0
                            Else
                                hgTagihanPasien.TextMatrix(i, 34) = 1
                            End If
                        Next i
        
                        Call setJudulTagihan
                    Else
                        If chkTagihanApotik.Value = Unchecked Then
        '                    dbConn.Execute "DELETE FROM PasienBelumBayar WHERE NoPendaftaran = '" & txtNoPendaftaran.Text & "' AND NoCM = '" & txtNoCM.Text & "'"
                            If sTempGrup = "frmCariPasien" Then
                                Call frmCariPasien.cmdCari_Click
                                sTempGrup = ""
                                frmCariPasien.Enabled = True
                                Unload Me
                                Exit Sub
                            End If
                        End If
                    End If
                    If mStatMulti = True Then
                        Call subHitungTotalMulti
                    Else
                        Call subHitungTotal
                    End If
                Else
                    Call subClearData
                End If
        
                If lblTotalTagihan.Caption = "" Or lblTotalTagihan.Caption = "Rp. 0" Then
                    mcurBayar = 0
                Else
                    mcurBayar = CCur(lblTotalTagihan.Caption)
                End If
            End If
        
            strSQL = "Select * from V_DaftarPasienYgBayarKredit WHERE NoCM='" & txtNoCM.Text & "' order by Pembayaranke desc"
            Call msubRecFO(rs, strSQL)
        
            If rs.RecordCount = 0 Then
                mcurTagihansebelumnya = 0
            Else
                LblTgihanSebelumnya.Caption = rs("SisaTagihan")
                mcurTagihansebelumnya = CCur(LblTgihanSebelumnya.Caption)
                LblTgihanSebelumnya.Caption = FormatCurrency(LblTgihanSebelumnya, 2)
            End If
        
            If bolValUpdate = True Then cmdBayar.Enabled = True
            If bolDatavalid = True Then
                cmdBayar.Enabled = True
            Else
                cmdBayar.Enabled = True
                cmdPerbaikiData.Enabled = False
            End If
    Else
        Call TagihanObat
    End If
    Exit Sub
errLoad:
    Call msubPesanError
'    Resume 0
End Sub

Private Sub TagihanObat()
    Dim i As Integer
    Dim j As Integer

On Error GoTo hell
        bolDatavalid = False
        Call subClearData
'        strSQL = "SELECT * FROM RegistrasiIGD WHERE NoPendaftaran = '" & txtNoPendaftaran.Text & "' AND StatusPulang = 'T'"
'        Call msubRecFO(rs, strSQL)
'        If rs.RecordCount <> 0 Then
'            MsgBox "Pasien belum keluar dari IGD", vbCritical
'            Exit Sub
'        End If
'
'        strSQL = "SELECT * FROM v_PasienAktifPakaiKamar WHERE NoPendaftaran='" _
'        & txtNoPendaftaran.Text & "'"
'        Set rs = Nothing
'        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
'        If rs.RecordCount <> 0 Then
'            MsgBox "Pasien belum keluar dari Rawat Inap", vbCritical
'            Exit Sub
'        End If

        strSQL = "SELECT NoCM FROM V_DaftarPasienSudahBayar WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "'"
        Set rs = Nothing:
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <> 0 Then
            mstrNoCMKu = rs.Fields(0).Value
            strSQL = "SELECT NoPendaftaran FROM PasienBelumBayar WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "'"
            Set rs = Nothing
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount = 0 Then
                strSQL = "INSERT INTO PasienBelumBayar VALUES ('" & txtNoPendaftaran.Text & "','" & mstrNoCMKu & "')"
                dbConn.Execute strSQL
            End If
        End If

        strSQL = "Select * from V_DaftarPasienBelumBayar_new_Fixed WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "'"
        
        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <> 0 Then
            strSQL = "SELECT * FROM PasienDaftar WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "'"
            Call msubRecFO(rsB, strSQL)

            strKdKelPsn = rsB("KdKelompokPasien").Value
            strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & txtNoPendaftaran.Text & "')"
            Call msubRecFO(rsB, strSQL)
            If rsB.EOF = False Then
                mstrKdJenisPasien = rsB("KdKelompokPasien").Value
                mstrKdPenjaminPasien = IIf(IsNull(rsB("IdPenjamin")), "2222222222", rsB("IdPenjamin"))
            End If

            If mstrKdPenjaminPasien <> "2222222222" Then
                strSQL = "SELECT * FROM PemakaianAsuransi WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "'"
                Call msubRecFO(rsB, strSQL)
                If rsB.RecordCount = 0 Then
                    MsgBox "Lengkapi dahulu data penjamin pasien", vbCritical, "Validasi"
                    Call subLoadFormJP
                    Exit Sub
                End If
            End If

            mstrNoPen = txtNoPendaftaran.Text
            txtNoCM.Text = rs.Fields(1).Value
            mstrNoCM = txtNoCM.Text
            txtNamaPasien.Text = rs.Fields(2).Value
            txtSex.Text = IIf(rs.Fields(3).Value = "P", "Perempuan", "Laki-Laki")
            txtThn.Text = rs("UmurTahun")
            txtBln.Text = rs("UmurBulan")
            txtHari.Text = rs("UmurHari")
            txtJenisPasien.Text = rs.Fields(5).Value

            If mstrKdPenjaminPasien <> "2222222222" Then
                strSQL = "Select * from V_PerusahaanPenjaminPasien WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "'"
                Call msubRecFO(rs, strSQL)
                txtPenjamin.Text = rs.Fields(1).Value
                mstrKdPenjamin = rs.Fields(2).Value
            Else
                txtPenjamin.Text = "Bayar Sendiri"
                mstrKdPenjamin = mstrKdPenjaminPasien
            End If

            cmdBayar.Enabled = False

            'utk looping add hutang penjamin dan tanggungan RS, yg sebelumnya diset 0 (blm diketahui)
            If bolValUpdate = False Then
                If mstrKdPenjaminPasien <> "2222222222" Then
                    fraPosting.Visible = True
                    Screen.MousePointer = vbHourglass
                    Me.Timer1.Enabled = True
                    If sp_PostingHutangPenjaminPasien_AU(mstrNoPen, "A") = False Then Exit Sub
                    Screen.MousePointer = vbDefault
                End If
            End If

            strSQL = "Select * from V_RincianTotalDetailBiayaOA WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "' and NoStruk Is Null" & sfilter & " ORDER BY TglPelayanan"
            Call msubRecFO(rs, strSQL)
            intJmlPelayanan = rs.RecordCount
            jData = intJmlPelayanan
            If rs.RecordCount <> 0 Then
                hgTagihanPasien.Clear
                hgTagihanPasien.Rows = rs.RecordCount + 1
                For i = 1 To rs.RecordCount
                    For j = 1 To 32
                        If j = 7 Or j = 21 Then
                            hgTagihanPasien.Row = i: hgTagihanPasien.Col = j: hgTagihanPasien.CellForeColor = vbBlue
                            If rs(j - 1).Value = 0 Then
                                hgTagihanPasien.TextMatrix(i, j) = "" & rs(j - 1).Value
                            Else
                                hgTagihanPasien.TextMatrix(i, j) = "" & rs(j - 1).Value
                            End If
                        Else
                            If j = 31 Or j = 32 Then
                                If hgTagihanPasien.TextMatrix(i, 25) = "OA" Then
                                    strSQL = "SELECT BiayaAdministrasi, TarifService, HargaSatuan From DetailPemakaianAlkes" & _
                                    " WHERE (NoPendaftaran = '" & txtNoPendaftaran.Text & "' ) AND (KdRuangan = '" & hgTagihanPasien.TextMatrix(i, 20) & "') AND (KdBarang = '" & hgTagihanPasien.TextMatrix(i, 3) & "') AND (KdAsal = '" & hgTagihanPasien.TextMatrix(i, 17) & "') AND (TglPelayanan = '" & Format(hgTagihanPasien.TextMatrix(i, 8), "yyyy/MM/dd HH:mm:ss") & "') AND (SatuanJml = '" & hgTagihanPasien.TextMatrix(i, 30) & "')"
                                    Call msubRecFO(rsB, strSQL)
                                    If rsB.EOF = False Then
                                        hgTagihanPasien.TextMatrix(i, 31) = rsB("BiayaAdministrasi")
                                        hgTagihanPasien.TextMatrix(i, 32) = rsB("TarifService")
                                        hgTagihanPasien.TextMatrix(i, 33) = rsB("HargaSatuan")
                                    Else
                                        hgTagihanPasien.TextMatrix(i, 31) = "0"
                                        hgTagihanPasien.TextMatrix(i, 32) = "0"
                                        hgTagihanPasien.TextMatrix(i, 33) = hgTagihanPasien.TextMatrix(i, 6)
                                    End If
                                Else
                                    hgTagihanPasien.TextMatrix(i, 31) = "0"
                                    hgTagihanPasien.TextMatrix(i, 32) = "0"
                                    hgTagihanPasien.TextMatrix(i, 33) = hgTagihanPasien.TextMatrix(i, 6)
                                End If
                            Else
                                hgTagihanPasien.TextMatrix(i, j) = "" & rs(j - 1).Value
                            End If
                        End If
                        If j = 1 Then hgTagihanPasien.TextMatrix(i, j) = Chr$(187)
                        If j = 21 Then
                            If val(hgTagihanPasien.TextMatrix(i, 7)) <> val(hgTagihanPasien.TextMatrix(i, j)) Then hgTagihanPasien.CellForeColor = vbRed
                            If val(hgTagihanPasien.TextMatrix(i, j)) = 0 Then hgTagihanPasien.CellForeColor = vbBlack
                        End If
                    Next j
                    ' ' untuk status verifikasi
                    
           
                       curnilaiproposional = 0
                    If bolinacbgs = True Then
                        curnilaiproposional = (txtTotalPenjamin.Text / curtotalbiaya) * hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 7)
                         hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 35) = Trim(hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, hgTagihanPasien.Col))
                        hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 21) = curnilaiproposional
                             
                    End If
                    
                    
                    hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 38) = rs.Fields("TotalHutangPenjamin") 'rs.Fields("JmlHutangPenjamin")
                    hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 39) = rs.Fields("TotalTanggunganRS") 'rs.Fields("JmlTanggunganRS")
                    hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 40) = rs.Fields("NoTerima")
                    hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 41) = rs.Fields("No_Item")
                    rs.MoveNext
                    ' untuk status verifikasi
                    With hgTagihanPasien
                        If .TextMatrix(i, 25) = "TM" Then
                            Set rsQuery = Nothing
'                           sqlQuery = "SELECT dbo.FB_TakeStatusDataValid('" & txtNoPendaftaran.Text & "','" & .TextMatrix(i, 20) & "','" & Trim(.TextMatrix(i, 3)) & "', '" & Format(.TextMatrix(i, 8), "yyyy/mm/dd HH:mm:ss") & "', " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 6))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 29))) & " , " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 38))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 39))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 23))) & " )  as statusData"
                            
                            sqlQuery = "SELECT dbo.FB_TakeStatusDataValidOA('" & txtNoPendaftaran.Text & "','" & .TextMatrix(i, 20) & "','" & Trim(.TextMatrix(i, 3)) & "', '" & Format(.TextMatrix(i, 8), "yyyy/mm/dd HH:mm:ss") & "', " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 6))) & " , " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 38))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 39))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 23))) & ",'" & .TextMatrix(i, 40) & "', '" & .TextMatrix(i, 41) & "')  as statusData"
                            Call msubRecFO(rsQuery, sqlQuery)
                            strStatusData = rsQuery.Fields(0).Value
                            .TextMatrix(i, 37) = strStatusData
                        ElseIf .TextMatrix(i, 25) = "OA" Then
                            Set rsQuery = Nothing
'                           sqlQuery = "SELECT dbo.FB_TakeStatusDataValidOA('" & txtNoPendaftaran.Text & "','" & .TextMatrix(i, 20) & "','" & Trim(.TextMatrix(i, 3)) & "', '" & Format(.TextMatrix(i, 8), "yyyy/mm/dd HH:mm:ss") & "', " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 6))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 38))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 39))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 23))) & ",'" & .TextMatrix(i, 40) & "')  as statusData"
                            sqlQuery = "SELECT dbo.FB_TakeStatusDataValidOA('" & txtNoPendaftaran.Text & "','" & .TextMatrix(i, 20) & "','" & Trim(.TextMatrix(i, 3)) & "', '" & Format(.TextMatrix(i, 8), "yyyy/mm/dd HH:mm:ss") & "', " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 6))) & " , " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 38))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 39))) & ", " & msubKonversiKomaTitik(CCur(.TextMatrix(i, 23))) & ",'" & .TextMatrix(i, 40) & "', '" & .TextMatrix(i, 41) & "')  as statusData"
                            Call msubRecFO(rsQuery, sqlQuery)
                            strStatusData = rsQuery.Fields(0).Value
                            .TextMatrix(i, 37) = strStatusData
                        End If
                    End With
                Next i

                ' untuk status verifikasi
                For i = 1 To rs.RecordCount
                    If hgTagihanPasien.TextMatrix(i, 37) = "T" Then
                        Dim k As Integer
                        With hgTagihanPasien
                            .Row = i
                            For k = 0 To .Cols - 1
                                .Col = k
'                                .CellBackColor = vbRed
'                                .CellForeColor = vbWhite
                                bolDatavalid = True
                            Next
                        End With
                        cmdBayar.Enabled = False
                    End If
                Next i

                For i = 1 To rs.RecordCount
                    If hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, hgTagihanPasien.Col) = Chr$(187) Then
                        hgTagihanPasien.TextMatrix(i, 34) = 0
                    Else
                        hgTagihanPasien.TextMatrix(i, 34) = 1
                    End If
                Next i

                Call setJudulTagihan
            Else
                If chkTagihanApotik.Value = Unchecked Then
'                    dbConn.Execute "DELETE FROM PasienBelumBayar WHERE NoPendaftaran = '" & txtNoPendaftaran.Text & "' AND NoCM = '" & txtNoCM.Text & "'"
                    If sTempGrup = "frmCariPasien" Then
                        Call frmCariPasien.cmdCari_Click
                        sTempGrup = ""
                        frmCariPasien.Enabled = True
                        Unload Me
                        Exit Sub
                    End If
                End If
            End If
            If mStatMulti = True Then
                Call subHitungTotalMulti
            Else
                Call subHitungTotal
            End If
        Else
            Call subClearData
        End If

        If lblTotalTagihan.Caption = "" Or lblTotalTagihan.Caption = "Rp. 0" Then
            mcurBayar = 0
        Else
            mcurBayar = CCur(lblTotalTagihan.Caption)
        End If
    

    strSQL = "Select * from V_DaftarPasienYgBayarKredit WHERE NoCM='" & txtNoCM.Text & "' order by Pembayaranke desc"
    Call msubRecFO(rs, strSQL)

    If rs.RecordCount = 0 Then
        mcurTagihansebelumnya = 0
    Else
        LblTgihanSebelumnya.Caption = rs("SisaTagihan")
        mcurTagihansebelumnya = CCur(LblTgihanSebelumnya.Caption)
        LblTgihanSebelumnya.Caption = FormatCurrency(LblTgihanSebelumnya, 2)
    End If

    If bolValUpdate = True Then cmdBayar.Enabled = True
    If bolDatavalid = True Then
        cmdBayar.Enabled = True
    Else
        cmdBayar.Enabled = True
        cmdPerbaikiData.Enabled = False
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub



'untuk cek validasi
Private Function funcCekValidasi() As Boolean
    Dim i As Integer
    Dim blnDataTerpilih As Boolean
    If txtNoPendaftaran.Text = "" Then
        MsgBox "No Pendaftaran pasien harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        txtNoPendaftaran.SetFocus
        Exit Function
    End If
    blnDataTerpilih = False
    mblnAdaPlynTdkDibyr = False
    With hgTagihanPasien
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) = Chr$(187) Then
                strKdRuanganBayar = .TextMatrix(i, 20)
                blnDataTerpilih = True
            End If
            If .TextMatrix(i, 1) = "" Then mblnAdaPlynTdkDibyr = True
        Next i
    End With
    
    If blnDataTerpilih = False Then
        MsgBox "Pilih tindakan yang hendak dibayar", vbCritical, "Validasi"
        funcCekValidasi = False
        hgTagihanPasien.SetFocus
        Exit Function
    End If
    funcCekValidasi = True
End Function

Private Sub subClearData()
    txtNoCM.Text = ""
    txtNamaPasien.Text = ""
    txtSex.Text = ""
    txtThn.Text = ""
    txtBln.Text = ""
    txtHari.Text = ""
    txtJenisPasien.Text = ""
    txtPenjamin.Text = ""
    Call setClearGridTagihan
    txtTotalBiaya.Text = ""
    txtTAsuransi.Text = ""
    txtTRS.Text = ""
    txtTM_TBP.Text = FormatCurrency(0, 4)
    txtTM_TP.Text = FormatCurrency(0, 4)
    txtTM_TRS.Text = FormatCurrency(0, 4)
    txtTM_HrsDibyr.Text = FormatCurrency(0, 4)
    txtOA_TBP.Text = FormatCurrency(0, 4)
    txtOA_TP.Text = FormatCurrency(0, 4)
    txtOA_TRS.Text = FormatCurrency(0, 4)
    txtOA_HrsDibyr.Text = FormatCurrency(0, 4)
End Sub

Private Sub setClearGridTagihan()
    Dim i As Integer
    With hgTagihanPasien
        .Clear
        .Rows = 2
        .Cols = 42

        .ColWidth(0) = 0 '320
        .ColWidth(1) = 340
        .ColWidth(2) = 2500
        .ColWidth(3) = 0
        .ColWidth(4) = 1800
        .ColWidth(5) = 400
        .ColWidth(6) = 1100
        .ColWidth(7) = 1000
        .ColWidth(8) = 1400
        .ColWidth(9) = 1000
        .ColWidth(10) = 1600
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        .ColWidth(16) = 0
        .ColWidth(17) = 0
        .ColWidth(18) = 0
        .ColWidth(19) = 0
        .ColWidth(20) = 0
        .ColWidth(21) = 1000
        .ColWidth(22) = 1000
        .ColWidth(23) = 1000
        .ColWidth(24) = 0
        .ColWidth(25) = 0
        .ColWidth(26) = 0
        .ColWidth(27) = 0
        .ColWidth(28) = 0
        .ColWidth(29) = 0
        .ColWidth(30) = 0
        .ColWidth(31) = 0
        .ColWidth(32) = 0
        .ColWidth(33) = 0
        .ColWidth(34) = 0
        .ColWidth(35) = 0
        .ColWidth(36) = 0
        .ColWidth(37) = 0
        .ColWidth(38) = 0
        .ColWidth(39) = 0
        .ColWidth(40) = 0
        .ColWidth(41) = 0
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignLeftCenter
        Call setJudulTagihan
    End With
End Sub

Private Sub subSetGridPaketKhususJamsostek()
    Dim i As Integer
    With hgPaketKhususJamsostek
        .Clear
        .Rows = 2
        .Cols = 35

        .ColWidth(0) = 0 '320
        .ColWidth(1) = 340
        .ColWidth(2) = 2500
        .ColWidth(3) = 0
        .ColWidth(4) = 1800
        .ColWidth(5) = 400
        .ColWidth(6) = 1100
        .ColWidth(7) = 1000
        .ColWidth(8) = 1400
        .ColWidth(9) = 1000
        .ColWidth(10) = 1600
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        .ColWidth(16) = 0
        .ColWidth(17) = 0
        .ColWidth(18) = 0
        .ColWidth(19) = 0
        .ColWidth(20) = 0
        .ColWidth(21) = 1000
        .ColWidth(22) = 1000
        .ColWidth(23) = 1000
        .ColWidth(24) = 0
        .ColWidth(25) = 0
        .ColWidth(26) = 0
        .ColWidth(27) = 0
        .ColWidth(28) = 0
        .ColWidth(29) = 0
        .ColWidth(30) = 0
        .ColWidth(31) = 0
        .ColWidth(32) = 0
        .ColWidth(33) = 0
        .ColWidth(34) = 0

        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignLeftCenter
        Call setJudulPaketKhususJamsostek
    End With
End Sub

Private Sub setJudulTagihan()
    Dim i As Integer
    With hgTagihanPasien
        .TextMatrix(0, 1) = ""
        .TextMatrix(0, 2) = "Jenis_Item"
        .TextMatrix(0, 3) = "Kode_Item"
        .TextMatrix(0, 4) = "Nama_Item"
        .TextMatrix(0, 5) = "Jml_Item"
        .TextMatrix(0, 6) = "Harga_Item"
        .TextMatrix(0, 7) = "SubTotal"
        .TextMatrix(0, 8) = "TglPelayanan"
        .TextMatrix(0, 9) = "Kelas"
        .TextMatrix(0, 10) = "Ruangan"
        .TextMatrix(0, 11) = "SubInstalasi"
        .TextMatrix(0, 12) = "Jenis_Pasien"
        .TextMatrix(0, 13) = "Jenis_Tarif"
        .TextMatrix(0, 14) = "Dokter"
        .TextMatrix(0, 15) = "KdKelas"
        .TextMatrix(0, 16) = "NoStruk"
        .TextMatrix(0, 17) = "KdAsal"
        .TextMatrix(0, 18) = "KdJenisTarif"
        .TextMatrix(0, 19) = "KdSubInstalasi"
        .TextMatrix(0, 20) = "KdRuangan"

        .TextMatrix(0, 21) = "TPenjamin"
        .TextMatrix(0, 22) = "TRS"
        .TextMatrix(0, 23) = "Pembebasan"
        .TextMatrix(0, 24) = "TotalHarusBayar"
        .TextMatrix(0, 25) = "Jenis"
        .TextMatrix(0, 26) = "NoLab_Rad"
        .TextMatrix(0, 27) = "HarusDibayarMinXJumlah"
        .TextMatrix(0, 28) = "Tarif"
        .TextMatrix(0, 29) = "TarifCito"
        .TextMatrix(0, 30) = "Satuan"
        .TextMatrix(0, 31) = "BiayaAdministrasi"
        .TextMatrix(0, 32) = "Service"
        .TextMatrix(0, 33) = "HargaSatuan"
        .TextMatrix(0, 34) = "Status Bayar"

        .TextMatrix(0, 35) = "tempHP"
        .TextMatrix(0, 36) = "tempTRS"
        .TextMatrix(0, 37) = "Status Verifikasi"
        .TextMatrix(0, 38) = "HP-PerPelayanan"
        .TextMatrix(0, 39) = "TRS-PerPelayanan"
        .TextMatrix(0, 40) = "NoTerima"
        .TextMatrix(0, 41) = "ResepKe"
    End With
End Sub

'Store procedure untuk menghapus biaya pelayanan pasien
Private Sub setJudulPaketKhususJamsostek()
    Dim i As Integer
    With hgPaketKhususJamsostek
        .TextMatrix(0, 1) = ""
        .TextMatrix(0, 2) = "Jenis_Item"
        .TextMatrix(0, 3) = "Kode_Item"
        .TextMatrix(0, 4) = "Nama_Item"
        .TextMatrix(0, 5) = "Jml_Item"
        .TextMatrix(0, 6) = "Harga_Item"
        .TextMatrix(0, 7) = "SubTotal"
        .TextMatrix(0, 8) = "TglPelayanan"
        .TextMatrix(0, 9) = "Kelas"
        .TextMatrix(0, 10) = "Ruangan"
        .TextMatrix(0, 11) = "SubInstalasi"
        .TextMatrix(0, 12) = "Jenis_Pasien"
        .TextMatrix(0, 13) = "Jenis_Tarif"
        .TextMatrix(0, 14) = "Dokter"
        .TextMatrix(0, 15) = "KdKelas"
        .TextMatrix(0, 16) = "NoStruk"
        .TextMatrix(0, 17) = "KdAsal"
        .TextMatrix(0, 18) = "KdJenisTarif"
        .TextMatrix(0, 19) = "KdSubInstalasi"
        .TextMatrix(0, 20) = "KdRuangan"

        .TextMatrix(0, 21) = "TPenjamin"
        .TextMatrix(0, 22) = "TRS"
        .TextMatrix(0, 23) = "Pembebasan"
        .TextMatrix(0, 24) = "TotalHarusBayar"
        .TextMatrix(0, 25) = "Jenis"
        .TextMatrix(0, 26) = "NoLab_Rad"
        .TextMatrix(0, 27) = "HarusDibayarMinXJumlah"
        .TextMatrix(0, 28) = "Tarif"
        .TextMatrix(0, 29) = "TarifCito"
        .TextMatrix(0, 30) = "Satuan"
        .TextMatrix(0, 31) = "BiayaAdministrasi"
        .TextMatrix(0, 32) = "Service"
        .TextMatrix(0, 33) = "HargaSatuan"
        .TextMatrix(0, 34) = "Status Bayar"
    End With
End Sub

Private Function sp_DeletePelayanan(strKdRuangan As String, strKdPelayanan As String, dTglPelayanan As Date) As Boolean
    On Error GoTo errLoad
    sp_DeletePelayanan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, strKdRuangan)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, RTrim(strKdPelayanan))
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , dTglPelayanan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Delete_BiayaPelayananNew"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penghapusan Biaya Pelayanan Pasien", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    Call msubPesanError
    sp_DeletePelayanan = False
End Function

Private Sub subHitungTotal()
    On Error GoTo errLoad
    Dim i As Integer

    mcurAll_TBP = 0: mcurAll_TP = 0: mcurAll_TRS = 0: mcurAll_Pemb = 0: mcurAll_HrsDibyr = 0
    mcurTM_TBP = 0: mcurTM_TP = 0: mcurTM_TRS = 0: mcurTM_Pemb = 0: mcurTM_HrsDibyr = 0: mcurTM_JmlByr = 0: mcurTM_ST = 0: mcurTM_HrsDibyrNow = 0
    mcurOA_TBP = 0: mcurOA_TP = 0: mcurOA_TRS = 0: mcurOA_Pemb = 0: mcurOA_HrsDibyr = 0: mcurOA_JmlByr = 0: mcurOA_ST = 0: mcurOA_HrsDibyrNow = 0
    mcurAll_TBP = FormatCurrency(mcurAll_TBP, 2)
    mcurTM_TBP = FormatCurrency(mcurTM_TBP, 2)
    mcurOA_TBP = FormatCurrency(mcurOA_TBP, 2)
    mcurPembebasan = 0
    mcurPembebasan = FormatCurrency(mcurPembebasan, 2)
    mblnTM = False
    mblnOA = False

    txtTotalBiaya.Text = 0: txtTAsuransi.Text = 0: txtTRS.Text = 0: txtTotalPembebasan.Text = 0

    txtTM_TBP.Text = 0: txtTM_TP.Text = 0: txtTM_TRS.Text = 0: txtTM_HrsDibyr.Text = 0: txtTMPembebasan.Text = 0
    txtOA_TBP.Text = 0: txtOA_TP.Text = 0: txtOA_TRS.Text = 0: txtOA_HrsDibyr.Text = 0: txtOAPembebasan.Text = 0
    mcurTM_TBP = FormatCurrency(mcurTM_TBP, 2)
    mcurOA_TBP = FormatCurrency(mcurOA_TBP, 2)

    For i = 1 To hgTagihanPasien.Rows - 1
        If hgTagihanPasien.TextMatrix(i, 1) = Chr$(187) Then
            txtTotalBiaya.Text = txtTotalBiaya.Text + CDbl(hgTagihanPasien.TextMatrix(i, 7))

            txtTAsuransi.Text = txtTAsuransi.Text + (CDbl(hgTagihanPasien.TextMatrix(i, 21)))
            txtTRS.Text = txtTRS.Text + (CDbl(hgTagihanPasien.TextMatrix(i, 22)))
            txtTotalPembebasan.Text = txtTotalPembebasan.Text + (CDbl(hgTagihanPasien.TextMatrix(i, 23)))

            If LCase(hgTagihanPasien.TextMatrix(i, 25)) = "tm" Then
                mblnTM = True
                txtTM_TBP.Text = txtTM_TBP.Text + CDbl(hgTagihanPasien.TextMatrix(i, 7))
                txtTM_TP.Text = txtTM_TP.Text + CDbl(hgTagihanPasien.TextMatrix(i, 21))
                txtTM_TRS.Text = txtTM_TRS.Text + CDbl(hgTagihanPasien.TextMatrix(i, 22))
                txtTMPembebasan.Text = txtTMPembebasan.Text + CDbl(hgTagihanPasien.TextMatrix(i, 23))
                txtTM_HrsDibyr.Text = txtTM_HrsDibyr.Text + CDbl(hgTagihanPasien.TextMatrix(i, 24))
            ElseIf LCase(hgTagihanPasien.TextMatrix(i, 25)) = "oa" Then
                mblnOA = True
                txtOA_TBP.Text = txtOA_TBP.Text + CDbl(hgTagihanPasien.TextMatrix(i, 7))
                txtOA_TP.Text = txtOA_TP.Text + CDbl(hgTagihanPasien.TextMatrix(i, 21))
                txtOA_TRS.Text = txtOA_TRS.Text + CDbl(hgTagihanPasien.TextMatrix(i, 22))
                txtOAPembebasan.Text = txtTMPembebasan.Text + CDbl(hgTagihanPasien.TextMatrix(i, 23))
                txtOA_HrsDibyr.Text = txtOA_HrsDibyr.Text + CDbl(hgTagihanPasien.TextMatrix(i, 24))
            End If
        End If
    Next i

    'Listing Baru Untuk Format Pembulatan 2 digit di belakang Koma
    lblTotalTagihan.Caption = (CCur(txtTM_HrsDibyr.Text) + CCur(txtOA_HrsDibyr.Text))
    lblTotalTagihan.Caption = FormatCurrency(lblTotalTagihan, 2)

    txtTotalBiaya.Text = FormatCurrency(txtTotalBiaya.Text, 2)
    txtTAsuransi.Text = FormatCurrency(txtTAsuransi.Text, 2)
    txtTRS.Text = FormatCurrency(txtTRS.Text, 2)
    txtTotalPembebasan.Text = FormatCurrency(txtTotalPembebasan.Text, 2)

    'TM
    txtTM_TBP.Text = FormatCurrency(txtTM_TBP.Text, 2)
    txtTM_TP.Text = FormatCurrency(txtTM_TP.Text, 2)
    txtTM_TRS.Text = FormatCurrency(txtTM_TRS.Text, 2)
    txtTMPembebasan.Text = FormatCurrency(txtTMPembebasan.Text, 2)
    txtTM_HrsDibyr.Text = FormatCurrency(txtTM_HrsDibyr.Text, 2)

    mcurTM_TBP = FormatCurrency(txtTM_TBP.Text, 2)
    mcurTM_TP = FormatCurrency(txtTM_TP.Text, 2)
    mcurTM_TRS = FormatCurrency(txtTM_TRS.Text, 2)
    mcurTM_Pemb = FormatCurrency(txtTMPembebasan.Text, 2)
    mcurTM_HrsDibyr = FormatCurrency(txtTM_HrsDibyr.Text, 2)
    mcurTM_JmlByr = FormatCurrency(0, 2): mcurTM_ST = FormatCurrency(0, 2)
    mcurTM_HrsDibyrNow = FormatCurrency(mcurTM_HrsDibyr, 2)

    'OA
    txtOA_TBP.Text = FormatCurrency(txtOA_TBP.Text, 2)
    txtOA_TP.Text = FormatCurrency(txtOA_TP.Text, 2)
    txtOA_TRS.Text = FormatCurrency(txtOA_TRS.Text, 2)
    txtOAPembebasan.Text = FormatCurrency(txtOAPembebasan.Text, 2)
    txtOA_HrsDibyr.Text = FormatCurrency(txtOA_HrsDibyr.Text, 2)

    'Format 4 digit
    mcurOA_TBP = FormatCurrency(CCur(txtOA_TBP.Text), 2)
    mcurOA_TP = FormatCurrency(CCur(txtOA_TP.Text), 2)
    mcurOA_TRS = FormatCurrency(CCur(txtOA_TRS.Text), 2)
    mcurOA_Pemb = FormatCurrency(CCur(txtOAPembebasan.Text), 2)
    mcurOA_HrsDibyr = FormatCurrency(CCur(txtOA_HrsDibyr.Text), 2)
    mcurOA_JmlByr = FormatCurrency(0, 2): mcurOA_ST = FormatCurrency(0, 2)
    mcurOA_JmlByr = FormatCurrency(CCur(mcurOA_JmlByr), 2)
    mcurOA_HrsDibyrNow = FormatCurrency(CCur(mcurOA_HrsDibyr), 2)

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subHitungTotalMulti()
    On Error GoTo errLoad
    Dim i As Integer

    mcurAll_TBP_M = 0: mcurAll_TP_M = 0: mcurAll_TRS_M = 0: mcurAll_Pemb_M = 0: mcurAll_HrsDibyr_M = 0
    mcurTM_TBP_M = 0: mcurTM_TP_M = 0: mcurTM_TRS_M = 0: mcurTM_Pemb_M = 0: mcurTM_HrsDibyr_M = 0: mcurTM_JmlByr_M = 0: mcurTM_ST_M = 0: mcurTM_HrsDibyrNow_M = 0
    mcurOA_TBP_M = 0: mcurOA_TP_M = 0: mcurOA_TRS_M = 0: mcurOA_Pemb_M = 0: mcurOA_HrsDibyr_M = 0: mcurOA_JmlByr_M = 0: mcurOA_ST_M = 0: mcurOA_HrsDibyrNow_M = 0

    mcurPembebasan_M = 0
    
    mblnTM = False
    mblnOA = False



    For i = 1 To hgTagihanPasien.Rows - 1
        If hgTagihanPasien.TextMatrix(i, 1) = Chr$(187) Then
            mcurAll_TBP_M = mcurAll_TBP_M + CDbl(hgTagihanPasien.TextMatrix(i, 7))

            mcurAll_TP_M = mcurAll_TP_M + (CDbl(hgTagihanPasien.TextMatrix(i, 21)))
            mcurAll_TRS_M = mcurAll_TRS_M + (CDbl(hgTagihanPasien.TextMatrix(i, 22)))
            mcurAll_Pemb_M = mcurAll_Pemb_M + (CDbl(hgTagihanPasien.TextMatrix(i, 23)))

            If LCase(hgTagihanPasien.TextMatrix(i, 25)) = "tm" Then
                mblnTM = True
                mcurTM_TBP_M = mcurTM_TBP_M + CDbl(hgTagihanPasien.TextMatrix(i, 7))
                mcurTM_TP_M = mcurTM_TP_M + CDbl(hgTagihanPasien.TextMatrix(i, 21))
                mcurTM_TRS_M = mcurTM_TRS_M + CDbl(hgTagihanPasien.TextMatrix(i, 22))
                mcurTM_Pemb_M = mcurTM_Pemb_M + CDbl(hgTagihanPasien.TextMatrix(i, 23))
                mcurTM_HrsDibyr_M = mcurTM_HrsDibyr_M + CDbl(hgTagihanPasien.TextMatrix(i, 24))
            ElseIf LCase(hgTagihanPasien.TextMatrix(i, 25)) = "oa" Then
                mblnOA = True
                mcurOA_TBP_M = mcurOA_TBP_M + CDbl(hgTagihanPasien.TextMatrix(i, 7))
                mcurOA_TP_M = mcurOA_TP_M + CDbl(hgTagihanPasien.TextMatrix(i, 21))
                mcurOA_TRS_M = mcurOA_TRS_M + CDbl(hgTagihanPasien.TextMatrix(i, 22))
                mcurOA_Pemb_M = mcurOA_Pemb_M + CDbl(hgTagihanPasien.TextMatrix(i, 23))
                mcurOA_HrsDibyr_M = mcurOA_HrsDibyr_M + CDbl(hgTagihanPasien.TextMatrix(i, 24))
            End If
        End If
    Next i

    'Listing Baru Untuk Format Pembulatan 2 digit di belakang Koma
'    lblTotalTagihan.Caption = (CCur(txtTM_HrsDibyr.Text) + CCur(txtOA_HrsDibyr.Text))
'    lblTotalTagihan.Caption = FormatCurrency(lblTotalTagihan, 2)
'
'    txtTotalBiaya.Text = FormatCurrency(txtTotalBiaya.Text, 2)
'    txtTAsuransi.Text = FormatCurrency(txtTAsuransi.Text, 2)
'    txtTRS.Text = FormatCurrency(txtTRS.Text, 2)
'    txtTotalPembebasan.Text = FormatCurrency(txtTotalPembebasan.Text, 2)
'
'    'TM
'    txtTM_TBP.Text = FormatCurrency(txtTM_TBP.Text, 2)
'    txtTM_TP.Text = FormatCurrency(txtTM_TP.Text, 2)
'    txtTM_TRS.Text = FormatCurrency(txtTM_TRS.Text, 2)
'    txtTMPembebasan.Text = FormatCurrency(txtTMPembebasan.Text, 2)
'    txtTM_HrsDibyr.Text = FormatCurrency(txtTM_HrsDibyr.Text, 2)
'
'    mcurTM_TBP = FormatCurrency(txtTM_TBP.Text, 2)
'    mcurTM_TP = FormatCurrency(txtTM_TP.Text, 2)
'    mcurTM_TRS = FormatCurrency(txtTM_TRS.Text, 2)
'    mcurTM_Pemb = FormatCurrency(txtTMPembebasan.Text, 2)
'    mcurTM_HrsDibyr = FormatCurrency(txtTM_HrsDibyr.Text, 2)
'    mcurTM_JmlByr = FormatCurrency(0, 2): mcurTM_ST = FormatCurrency(0, 2)
'    mcurTM_HrsDibyrNow = FormatCurrency(mcurTM_HrsDibyr, 2)
'
'    'OA
'    txtOA_TBP.Text = FormatCurrency(txtOA_TBP.Text, 2)
'    txtOA_TP.Text = FormatCurrency(txtOA_TP.Text, 2)
'    txtOA_TRS.Text = FormatCurrency(txtOA_TRS.Text, 2)
'    txtOAPembebasan.Text = FormatCurrency(txtOAPembebasan.Text, 2)
'    txtOA_HrsDibyr.Text = FormatCurrency(txtOA_HrsDibyr.Text, 2)
'
'    'Format 4 digit
'    mcurOA_TBP = FormatCurrency(CCur(txtOA_TBP.Text), 2)
'    mcurOA_TP = FormatCurrency(CCur(txtOA_TP.Text), 2)
'    mcurOA_TRS = FormatCurrency(CCur(txtOA_TRS.Text), 2)
'    mcurOA_Pemb = FormatCurrency(CCur(txtOAPembebasan.Text), 2)
'    mcurOA_HrsDibyr = FormatCurrency(CCur(txtOA_HrsDibyr.Text), 2)
'    mcurOA_JmlByr = FormatCurrency(0, 2): mcurOA_ST = FormatCurrency(0, 2)
'    mcurOA_JmlByr = FormatCurrency(CCur(mcurOA_JmlByr), 2)
'    mcurOA_HrsDibyrNow = FormatCurrency(CCur(mcurOA_HrsDibyr), 2)

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

'paket jamsostek khusus
Private Sub subHitungTotalPembagianPaketKhusus()
    On Error GoTo hell
    Dim i As Integer
    Dim j As Integer
    Dim dblTotalBagi As Double

    curTotalBiayaDipilih = 0
    txtTotalPembagian.Text = "0"

    For j = 1 To hgPaketKhususJamsostek.Rows - 1
        If hgPaketKhususJamsostek.TextMatrix(j, 34) = 1 Then
            curTotalBiayaDipilih = curTotalBiayaDipilih + hgPaketKhususJamsostek.TextMatrix(j, 7)
        End If
    Next j

    For i = 1 To hgPaketKhususJamsostek.Rows - 1
        If hgPaketKhususJamsostek.TextMatrix(i, 34) = 1 Then
            dblTotalBagi = CDec(CDec(hgPaketKhususJamsostek.TextMatrix(i, 7)) / curTotalBiayaDipilih) * CCur(txtTarifTanggungan.Text)
            hgPaketKhususJamsostek.TextMatrix(i, 21) = Format(dblTotalBagi, "###,###.###")
            txtTotalPembagian.Text = Format(CCur(txtTotalPembagian.Text) + CDec(dblTotalBagi), "###,###")
        End If
    Next i
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subLoadText()
    Dim i As Integer
    txtIsi.Left = hgTagihanPasien.Left
    Select Case hgTagihanPasien.Col
        Case 21, 22, 23
        Case Else
            Exit Sub
    End Select

    For i = 0 To hgTagihanPasien.Col - 1
        txtIsi.Left = txtIsi.Left + hgTagihanPasien.ColWidth(i)
    Next i
    txtIsi.Visible = True
    txtIsi.Top = hgTagihanPasien.Top - 7

    For i = 0 To hgTagihanPasien.Row - 1
        txtIsi.Top = txtIsi.Top + hgTagihanPasien.RowHeight(i)
    Next i

    If hgTagihanPasien.TopRow > 1 Then
        txtIsi.Top = txtIsi.Top - ((hgTagihanPasien.TopRow - 1) * hgTagihanPasien.RowHeight(1))
    End If

    txtIsi.Width = hgTagihanPasien.ColWidth(hgTagihanPasien.Col)
    txtIsi.Height = hgTagihanPasien.RowHeight(hgTagihanPasien.Row)

    txtIsi.Visible = True
    txtIsi.SelStart = Len(txtIsi.Text)
    txtIsi.SetFocus
End Sub

Private Sub subLoadTextUpdate()
    Dim i As Integer
    txtIsiUpdate.Left = fgUpdateKomponen.Left

    For i = 0 To fgUpdateKomponen.Col - 1
        txtIsiUpdate.Left = txtIsiUpdate.Left + fgUpdateKomponen.ColWidth(i)
    Next i
    txtIsiUpdate.Visible = True
    txtIsiUpdate.Top = fgUpdateKomponen.Top - 7

    For i = 0 To fgUpdateKomponen.Row - 1
        txtIsiUpdate.Top = txtIsiUpdate.Top + fgUpdateKomponen.RowHeight(i)
    Next i

    If fgUpdateKomponen.TopRow > 1 Then
        txtIsiUpdate.Top = txtIsiUpdate.Top - ((fgUpdateKomponen.TopRow - 1) * fgUpdateKomponen.RowHeight(1))
    End If

    txtIsiUpdate.Width = fgUpdateKomponen.ColWidth(fgUpdateKomponen.Col)
    txtIsiUpdate.Height = fgUpdateKomponen.RowHeight(fgUpdateKomponen.Row)

    txtIsiUpdate.Visible = True
    txtIsiUpdate.SelStart = Len(txtIsiUpdate.Text)
    txtIsiUpdate.SetFocus
End Sub

Private Function sp_Update_TempHargaKomponen4PasienNU(f_NoPendaftaran As String, f_KdRuangan As String, f_KdPelayananRS As String, _
    f_tglPelayanan As Date, f_KdKomponen As String, f_HutangPenjamin As Currency, f_TanggunganRS As Currency, f_JmlPembebasan As Currency) As Boolean
    On Error GoTo errLoad
    sp_Update_TempHargaKomponen4PasienNU = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, adInteger, Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("KdPelayananRS", adVarChar, adParamInput, 9, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_tglPelayanan, "yyyy/MM/dd hh:mm:ss"))
        .Parameters.Append .CreateParameter("KdKomponen", adVarChar, adParamInput, 9, f_KdKomponen)
        .Parameters.Append .CreateParameter("JmlHutangPenjamin", adCurrency, adParamInput, , f_HutangPenjamin)
        .Parameters.Append .CreateParameter("JmlTanggunganRS", adCurrency, adParamInput, , f_TanggunganRS)
        .Parameters.Append .CreateParameter("JmlPembebasan", adCurrency, adParamInput, , f_JmlPembebasan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "K")

        .ActiveConnection = dbConn
        .CommandText = "Update_TempHargaKomponen4PasienNU"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_Update_TempHargaKomponen4PasienNU = False

        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Exit Function
errLoad:
    Call msubPesanError
    sp_Update_TempHargaKomponen4PasienNU = False
End Function

Private Function sp_DetailBiayaPelayanan4PasienNU(f_NoPendaftaran As String, f_KdRuangan As String, f_KdItem As String, _
    f_tglPelayanan As Date, f_KdAsal As String, f_HutangPenjamin As Currency, f_TanggunganRS As Currency, f_JmlPembebasan As Currency, f_Jenis As String, f_Satuan As String) As Boolean
    Dim strKomponen12 As String
    On Error GoTo errLoad
    sp_DetailBiayaPelayanan4PasienNU = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, adInteger, Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("Kode_Item", adVarChar, adParamInput, 9, Trim(f_KdItem))
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_tglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("JmlHutangPenjamin", adCurrency, adParamInput, , f_HutangPenjamin)
        .Parameters.Append .CreateParameter("JmlTanggunganRS", adCurrency, adParamInput, , f_TanggunganRS)
        .Parameters.Append .CreateParameter("JmlPembebasan", adCurrency, adParamInput, , f_JmlPembebasan)
        .Parameters.Append .CreateParameter("Jenis", adChar, adParamInput, 2, f_Jenis)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, f_Satuan)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "T")

        .ActiveConnection = dbConn
        .CommandText = "Update_DetailBiayaPelayanan4PasienNU"
        .CommandType = adCmdStoredProc
        .Execute

        strKomponen12 = "update tempHargaKomponen set JmlHutangPenjamin = " & msubKonversiKomaTitik(CCur(f_HutangPenjamin)) & ",  JmlTanggunganRS = " & msubKonversiKomaTitik(CCur(f_TanggunganRS)) & ", JmlPembebasan = " & msubKonversiKomaTitik(CCur(f_JmlPembebasan)) & " where NoPendaftaran = '" & Trim(f_NoPendaftaran) & "' and year(TglPelayanan) = '" & Year(f_tglPelayanan) & "'  and month(TglPelayanan) = '" & Month(f_tglPelayanan) & "'  and day(TglPelayanan) = '" & Day(f_tglPelayanan) & "'  and datepart(hh,TglPelayanan) = '" & Hour(f_tglPelayanan) & "'  and datepart(mi,TglPelayanan) = '" & Minute(f_tglPelayanan) & "'  and datepart(ss,TglPelayanan) = '" & Second(f_tglPelayanan) & "' and KdPelayananRS = '" & Trim(f_KdItem) & "' and KdKomponen = '12'"
        dbConn.Execute strKomponen12

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_DetailBiayaPelayanan4PasienNU = False
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Exit Function
errLoad:
    Call msubPesanError
    sp_DetailBiayaPelayanan4PasienNU = False
End Function

Private Sub subLoadFormJP()
    On Error GoTo hell
    mstrNoPen = txtNoPendaftaran
    mstrNoCM = txtNoCM
    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If

    With frmUbahJenisPasien
        .Show
        .txtNamaFormPengirim.Text = Me.Name
        .txtNoCM.Text = txtNoCM.Text
        .txtNamaPasien.Text = txtNamaPasien.Text
        .txtJK.Text = txtSex.Text
        .txtThn.Text = txtThn.Text
        .txtBln.Text = txtBln.Text
        .txtHr.Text = txtHari.Text
        .lblNoPendaftaran.Visible = False
        .txtNoPendaftaran.Visible = False
        .dcJenisPasien.BoundText = mstrKdJenisPasien
        .dcPenjamin.BoundText = mstrKdPenjaminPasien
    End With
    Exit Sub
hell:
End Sub

Private Sub sp_DelBiayaPelayananNaikKelas(ByVal adoCommand As ADODB.Command)
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
'        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, strGlobalKdRuanganNaikKelas)
'        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Delete_BiayaPelayananNaikKelas"
        .CommandType = adCmdStoredProc
'        dbConn.BeginTrans
        .Execute
'        dbConn.RollbackTrans

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penghapusan Biaya Pelayanan Pasien", vbCritical, "Validasi"
        Else
'            Call Add_HistoryLoginActivity("Delete_BiayaPelayananNaikKelas")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

Private Function sp_BiayaPelayananNaikKelas(ByVal adoCommand As ADODB.Command, curTarif As Currency) As Boolean
    
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuanganNaikKelas", adChar, adParamInput, 3, strGlobalKdRuanganNaikKelas)
        .Parameters.Append .CreateParameter("KdKelasNaikKelas", adChar, adParamInput, 2, strGlobalKdKelasNaikKelas)
        .Parameters.Append .CreateParameter("Tarif", adInteger, adParamInput, , curTarif)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(Now, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
   
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_BiayaPelayananNaikKelas"
        .CommandType = adCmdStoredProc
        
'        dbConn.BeginTrans
        .Execute
'        dbConn.RollbackTrans
        
        
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Biaya Pelayanan Naik Kelas", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            sp_BiayaPelayananNaikKelas = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
        sp_BiayaPelayananNaikKelas = True
    End With
End Function

