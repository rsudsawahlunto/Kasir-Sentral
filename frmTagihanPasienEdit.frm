VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTagihanPasienEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Edit Tagihan Pasien"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTagihanPasienEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   14925
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   58
      Top             =   8325
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6535
            Text            =   "Rincian Biaya Sementara (F1)"
            TextSave        =   "Rincian Biaya Sementara (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6535
            Text            =   "Ubah Biaya Pelayanan (F5)"
            TextSave        =   "Ubah Biaya Pelayanan (F5)"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6535
            Text            =   "Tambah Pelayanan (F6)"
            TextSave        =   "Tambah Pelayanan (F6)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6535
            Text            =   "Hapus Pelayanan (F7)"
            TextSave        =   "Hapus Pelayanan (F7)"
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
   Begin VB.Frame fraDetailRekap 
      Height          =   2535
      Left            =   0
      TabIndex        =   47
      Top             =   4080
      Visible         =   0   'False
      Width           =   14895
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
         TabIndex        =   53
         Top             =   240
         Width           =   14655
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
            Left            =   5160
            TabIndex        =   11
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
            Left            =   2640
            TabIndex        =   10
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
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   2415
         End
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
            Left            =   7680
            TabIndex        =   12
            Top             =   480
            Width           =   2415
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
            Left            =   5160
            TabIndex        =   57
            Top             =   240
            Width           =   2445
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
            Left            =   2640
            TabIndex        =   56
            Top             =   240
            Width           =   2115
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
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   2130
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
            Left            =   7680
            TabIndex        =   54
            Top             =   240
            Width           =   1380
         End
      End
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
         TabIndex        =   48
         Top             =   1320
         Width           =   14655
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
            Left            =   120
            TabIndex        =   13
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
            Left            =   2640
            TabIndex        =   14
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
            Left            =   5160
            TabIndex        =   15
            Top             =   480
            Width           =   2415
         End
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
            Left            =   7680
            TabIndex        =   16
            Top             =   480
            Width           =   2415
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
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   2130
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
            Left            =   2640
            TabIndex        =   51
            Top             =   240
            Width           =   2115
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
            Left            =   5160
            TabIndex        =   50
            Top             =   240
            Width           =   2445
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
            Left            =   7680
            TabIndex        =   49
            Top             =   240
            Width           =   1380
         End
      End
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
      Height          =   765
      Left            =   0
      TabIndex        =   27
      Top             =   7560
      Width           =   14895
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   2160
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdBayar 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1935
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
      TabIndex        =   39
      Top             =   6600
      Width           =   14895
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
         Left            =   7440
         TabIndex        =   20
         Top             =   480
         Width           =   855
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
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   2175
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
         Left            =   2520
         TabIndex        =   18
         Top             =   480
         Width           =   2175
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
         Left            =   4920
         TabIndex        =   19
         Top             =   480
         Width           =   2415
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
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   2130
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
         Left            =   2520
         TabIndex        =   42
         Top             =   240
         Width           =   2115
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
         Left            =   4920
         TabIndex        =   41
         Top             =   240
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
      TabIndex        =   40
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
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkCheck 
         Height          =   210
         Left            =   480
         TabIndex        =   44
         Top             =   4000
         Visible         =   0   'False
         Width           =   200
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgTagihanPasien1 
         Height          =   3135
         Left            =   2040
         TabIndex        =   23
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
         TabIndex        =   46
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
         TabIndex        =   45
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
      TabIndex        =   28
      Top             =   1800
      Width           =   14895
      Begin VB.TextBox txtPenjamin 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   11520
         TabIndex        =   8
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtJenisPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   9000
         TabIndex        =   7
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   2
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   855
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
         Left            =   6480
         TabIndex        =   34
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
            TabIndex        =   4
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
            TabIndex        =   5
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
            TabIndex        =   6
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Left            =   550
            TabIndex        =   37
            Top             =   277
            Width           =   285
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Left            =   1350
            TabIndex        =   36
            Top             =   277
            Width           =   240
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Left            =   2130
            TabIndex        =   35
            Top             =   270
            Width           =   165
         End
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Penjamin"
         Height          =   210
         Left            =   11520
         TabIndex        =   38
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pasien"
         Height          =   210
         Left            =   9000
         TabIndex        =   33
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   5280
         TabIndex        =   32
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   2520
         TabIndex        =   31
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1560
         TabIndex        =   30
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   120
         TabIndex        =   29
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
      TabIndex        =   24
      Top             =   960
      Width           =   14895
      Begin VB.Label lblTotalTagihan 
         Alignment       =   1  'Right Justify
         Caption         =   "Rp. 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   10320
         TabIndex        =   26
         Top             =   240
         Width           =   4440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Tagihan ->"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   6600
         TabIndex        =   25
         Top             =   240
         Width           =   3690
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   60
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
      Picture         =   "frmTagihanPasienEdit.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   13080
      Picture         =   "frmTagihanPasienEdit.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTagihanPasienEdit.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "frmTagihanPasienEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strKdKelPsn As String
Dim mstrNoCMKu As String
Dim subbolEditTanggungan As Boolean

Private Sub chkCheck_Click()
    On Error GoTo errLoad

    If chkCheck.Value = vbChecked Then
        hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, hgTagihanPasien.Col) = Chr$(187)
    Else
        hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, hgTagihanPasien.Col) = ""
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

Private Sub cmdBayar_Click()
    If funcCekValidasi = False Then Exit Sub
    If subbolEditTanggungan = True Then
        If MsgBox("Anda telah merubah besar tanggungan pasien, " & vbNewLine & "pilih Yes untuk meneruskan transaksi", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
        If UpdateDetailBiayaPelayanan = False Then Exit Sub

        cmdBayar.Enabled = False
        Call setClearGridTagihan
        Exit Sub
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Function UpdateDetailBiayaPelayanan() As Boolean
    On Error GoTo errLoad
    Dim i As Integer

    'u/ intern function
    UpdateDetailBiayaPelayanan = True

    'update ke detail biaya pelayanan beban yang ditanggung penjamin/rs
    For i = 1 To hgTagihanPasien.Rows - 1
        If CDbl(hgTagihanPasien.TextMatrix(i, 5)) <> 0 Then
            If sp_DetailBiayaPelayanan4PasienNU(mstrNoPen, _
                hgTagihanPasien.TextMatrix(i, 20), _
                hgTagihanPasien.TextMatrix(i, 3), _
                CDate(hgTagihanPasien.TextMatrix(i, 8)), _
                hgTagihanPasien.TextMatrix(i, 17), _
                CDbl(hgTagihanPasien.TextMatrix(i, 21)) / CDbl(hgTagihanPasien.TextMatrix(i, 5)), _
                CDbl(hgTagihanPasien.TextMatrix(i, 22)) / CDbl(hgTagihanPasien.TextMatrix(i, 5)), _
                CDbl(hgTagihanPasien.TextMatrix(i, 23)) / CDbl(hgTagihanPasien.TextMatrix(i, 5)), _
                hgTagihanPasien.TextMatrix(i, 25), _
                hgTagihanPasien.TextMatrix(i, 30)) = False Then Exit Function
            End If
        Next i
        Call txtNoPendaftaran_KeyPress(13)
        cmdBayar.SetFocus

        'tanggungan sudah diedit
        subbolEditTanggungan = False

        Exit Function
errLoad:
        UpdateDetailBiayaPelayanan = False
        Call msubPesanError
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errLoad

    Select Case KeyCode
        Case vbKeyF1
            If txtNoPendaftaran.Text = "" Then Exit Sub
            If hgTagihanPasien.Rows = 2 And hgTagihanPasien.TextMatrix(1, 4) = "" Then Exit Sub
            mstrNoPen = txtNoPendaftaran.Text
            frm_cetak_RincianBiaya.Show

        Case vbKeyF5
            If mblnAdmin = False Then Exit Sub
            If hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 3) = "" Then Exit Sub
            strSQL = "SELECT * " & _
            " FROM V_UbahBiayaPelayanan" & _
            " WHERE NoPendaftaran = '" & txtNoPendaftaran.Text & "' AND (KdPelayananRS = '" & Trim(hgTagihanPasien.TextMatrix(frmTagihanPasien.hgTagihanPasien.Row, 3)) & "')AND (TglPelayanan = '" & Format(hgTagihanPasien.TextMatrix(frmTagihanPasien.hgTagihanPasien.Row, 8), "yyyy/MM/dd HH:mm:ss") & "')AND(KdRuangan = '" & hgTagihanPasien.TextMatrix(frmTagihanPasien.hgTagihanPasien.Row, 20) & "')"
            Call msubRecFO(rs, strSQL)
            If rs.EOF Then Exit Sub

            Me.Enabled = False

            With frmUpdateBiayaPelayanan
                .txtNoPendaftaran = txtNoPendaftaran.Text
                Call .txtNoPendaftaran_KeyPress(13)
            End With

        Case vbKeyF6
            If txtNoCM.Text = "" Then Exit Sub
            mblnTindakanKasir = True
            frmPilihSubIns.Show

        Case vbKeyF7
            If mblnAdmin = False Then Exit Sub
            If hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 4) = "" Then Exit Sub
            If MsgBox("Apakah anda yakin akan menghapus pelayanan '" _
            & hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 4) & "'" & vbNewLine _
            & "Dengan tanggal pelayanan '" & hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 8) _
            & "'", vbQuestion + vbYesNo) = vbNo Then Exit Sub

            sp_DelBiayaPelayanan dbcmd
            MsgBox "Data berhasil dihapus", vbInformation
            Call txtNoPendaftaran_KeyPress(13)
            chkCheck.Visible = False
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    txtNoPendaftaran.Text = Right(Year(Now), 2) & Format(Month(Now), "00") & Format(Day(Now), "00")
    txtNoPendaftaran.SelStart = Len(txtNoPendaftaran.Text)

    StatusBar1.Panels.Item(3).Visible = False 'Tambah Pelayanan
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
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If blnFrmCariPasien = True Then
        Call frmCariPasien.cmdCari_Click
        frmCariPasien.Enabled = True
    End If
End Sub

Private Sub hgTagihanPasien_DblClick()
    On Error GoTo errLoad

    If hgTagihanPasien.Rows = 1 Then Exit Sub
    If hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 3) = "" Then Exit Sub
    chkCheck.Visible = False

    Select Case hgTagihanPasien.Col
        Case 1
            If mblnAdmin = False Then Exit Sub

            chkCheck.Visible = True
            chkCheck.Top = hgTagihanPasien.RowPos(hgTagihanPasien.Row) + 390
            Dim intA As Integer
            intA = ((hgTagihanPasien.ColPos(hgTagihanPasien.Col + 1) - hgTagihanPasien.ColPos(hgTagihanPasien.Col)) / 2)
            chkCheck.Left = hgTagihanPasien.ColPos(hgTagihanPasien.Col) + 160 + intA
            chkCheck.SetFocus
            If hgTagihanPasien.Col = 1 Then
                If hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 1) <> "" Then
                    chkCheck.Value = 1
                Else
                    chkCheck.Value = 0
                End If
            End If
        
        Case 21, 22
            If mblnAdmin = False Then Exit Sub
            subbolEditTanggungan = True
            Call subLoadText
            txtIsi.Text = Trim(hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, hgTagihanPasien.Col))
            txtIsi.SelStart = 0: txtIsi.SelLength = Len(txtIsi.Text)
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

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If KeyAscii = 13 Then
        hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, hgTagihanPasien.Col) = val(txtIsi.Text)
        txtIsi.Visible = False

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

Public Sub txtNoPendaftaran_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    Dim i As Integer
    Dim j As Integer

    If KeyAscii = 13 Then
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

        strSQL = "Select * from V_DaftarPasienBelumBayar WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "'"
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
            txtThn.Text = rs.Fields(8).Value
            txtBln.Text = rs.Fields(9).Value
            txtHari.Text = rs.Fields(10).Value
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

            strSQL = "Select * from V_RincianTotalDetailBiayaPelayanan WHERE NoPendaftaran='" & txtNoPendaftaran.Text & "'"
            Call msubRecFO(rs, strSQL)

            If rs.RecordCount <> 0 Then
                hgTagihanPasien.Clear
                hgTagihanPasien.Rows = rs.RecordCount + 1
                For i = 1 To rs.RecordCount
                    For j = 1 To hgTagihanPasien.Cols - 1
                        hgTagihanPasien.TextMatrix(i, j) = "" & rs(j - 1).Value
                        If j = 1 Then hgTagihanPasien.TextMatrix(i, j) = Chr$(187)
                    Next j
                    rs.MoveNext
                Next i
                Call setJudulTagihan
            End If
            Call subHitungTotal
        Else
            Call subClearData
        End If

        If lblTotalTagihan.Caption = "" Or lblTotalTagihan.Caption = "Rp. 0" Then
            mcurBayar = 0
        Else
            mcurBayar = CCur(lblTotalTagihan.Caption)
        End If
    End If

    Exit Sub
errLoad:
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
            If .TextMatrix(i, 1) = Chr$(187) Then blnDataTerpilih = True
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
    txtTM_TBP.Text = FormatCurrency(0, 2)
    txtTM_TP.Text = FormatCurrency(0, 2)
    txtTM_TRS.Text = FormatCurrency(0, 2)
    txtTM_HrsDibyr.Text = FormatCurrency(0, 2)
    txtOA_TBP.Text = FormatCurrency(0, 2)
    txtOA_TP.Text = FormatCurrency(0, 2)
    txtOA_TRS.Text = FormatCurrency(0, 2)
    txtOA_HrsDibyr.Text = FormatCurrency(0, 2)
End Sub

Private Sub setClearGridTagihan()
    Dim i As Integer
    With hgTagihanPasien
        .Clear
        .Rows = 2
        .Cols = 33

        .ColWidth(0) = 320
        .ColWidth(1) = 340
        .ColWidth(2) = 2500
        .ColWidth(3) = 0
        .ColWidth(4) = 1800
        .ColWidth(5) = 800
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
        .ColWidth(21) = 1250
        .ColWidth(22) = 1250
        .ColWidth(23) = 0 'Pembebasan
        .ColWidth(24) = 0
        .ColWidth(25) = 0 'Jenis Transaksi
        .ColWidth(26) = 0
        .ColWidth(27) = 0
        .ColWidth(28) = 0
        .ColWidth(29) = 0
        .ColWidth(30) = 0 'Satuan
        .ColWidth(31) = 0
        .ColWidth(32) = 0
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignLeftCenter
        Call setJudulTagihan
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

        .TextMatrix(0, 21) = "Total TPenjamin"
        .TextMatrix(0, 22) = "Total TRS"
        .TextMatrix(0, 23) = "TotalPembebasan"
        .TextMatrix(0, 24) = "TotalHarusBayar"
        .TextMatrix(0, 25) = "Jenis"

        .TextMatrix(0, 26) = "NoLAB"
        .TextMatrix(0, 27) = "Total TRS"
        .TextMatrix(0, 28) = "HarusDibayar"
        .TextMatrix(0, 29) = "TotalHarusBayar"
        .TextMatrix(0, 30) = "Jenis"
    End With
End Sub

'Store procedure untuk menghapus biaya pelayanan pasien
Private Sub sp_DelBiayaPelayanan(ByVal adoCommand As ADODB.Command)
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 20))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, Trim(hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 3)))
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(hgTagihanPasien.TextMatrix(hgTagihanPasien.Row, 8), "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "Delete_BiayaPelayananNew"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penghapusan Biaya Pelayanan Pasien", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Delete_BiayaPelayanan")

        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

Private Sub subHitungTotal()
    On Error GoTo errLoad
    Dim i As Integer

    mcurAll_TBP = 0: mcurAll_TP = 0: mcurAll_TRS = 0: mcurAll_Pemb = 0: mcurAll_HrsDibyr = 0
    mcurTM_TBP = 0: mcurTM_TP = 0: mcurTM_TRS = 0: mcurTM_Pemb = 0: mcurTM_HrsDibyr = 0: mcurTM_JmlByr = 0: mcurTM_ST = 0: mcurTM_HrsDibyrNow = 0
    mcurOA_TBP = 0: mcurOA_TP = 0: mcurOA_TRS = 0: mcurOA_Pemb = 0: mcurOA_HrsDibyr = 0: mcurOA_JmlByr = 0: mcurOA_ST = 0: mcurOA_HrsDibyrNow = 0
    mcurPembebasan = 0
    mblnTM = False
    mblnOA = False

    txtTotalBiaya.Text = 0: txtTAsuransi.Text = 0: txtTRS.Text = 0

    txtTM_TBP.Text = 0: txtTM_TP.Text = 0: txtTM_TRS.Text = 0: txtTM_HrsDibyr.Text = 0
    txtOA_TBP.Text = 0: txtOA_TP.Text = 0: txtOA_TRS.Text = 0: txtOA_HrsDibyr.Text = 0

    For i = 1 To hgTagihanPasien.Rows - 1
        If hgTagihanPasien.TextMatrix(i, 1) = Chr$(187) Then
            txtTotalBiaya.Text = txtTotalBiaya.Text + CDbl(hgTagihanPasien.TextMatrix(i, 7))
            txtTAsuransi.Text = txtTAsuransi.Text + (CDbl(hgTagihanPasien.TextMatrix(i, 21)))
            txtTRS.Text = txtTRS.Text + (CDbl(hgTagihanPasien.TextMatrix(i, 22)))

            If LCase(hgTagihanPasien.TextMatrix(i, 24)) = "tm" Then
                mblnTM = True
                txtTM_TBP.Text = txtTM_TBP.Text + CDbl(hgTagihanPasien.TextMatrix(i, 7))
                txtTM_TP.Text = txtTM_TP.Text + CDbl(hgTagihanPasien.TextMatrix(i, 21))
                txtTM_TRS.Text = txtTM_TRS.Text + CDbl(hgTagihanPasien.TextMatrix(i, 22))
                txtTM_HrsDibyr.Text = txtTM_HrsDibyr.Text + CDbl(hgTagihanPasien.TextMatrix(i, 23))
            ElseIf LCase(hgTagihanPasien.TextMatrix(i, 24)) = "oa" Then
                mblnOA = True
                txtOA_TBP.Text = txtOA_TBP.Text + CDbl(hgTagihanPasien.TextMatrix(i, 7))
                txtOA_TP.Text = txtOA_TP.Text + CDbl(hgTagihanPasien.TextMatrix(i, 21))
                txtOA_TRS.Text = txtOA_TRS.Text + CDbl(hgTagihanPasien.TextMatrix(i, 22))
                txtOA_HrsDibyr.Text = txtOA_HrsDibyr.Text + CDbl(hgTagihanPasien.TextMatrix(i, 23))
            End If
        End If
    Next i

    lblTotalTagihan.Caption = CCur(txtTM_HrsDibyr.Text) + CCur(txtOA_HrsDibyr.Text)

    lblTotalTagihan.Caption = FormatCurrency(lblTotalTagihan, 2)

    txtTotalBiaya.Text = FormatCurrency(txtTotalBiaya.Text, 2)
    txtTAsuransi.Text = FormatCurrency(txtTAsuransi.Text, 2)
    txtTRS.Text = FormatCurrency(txtTRS.Text, 2)

    txtTM_TBP.Text = FormatCurrency(txtTM_TBP.Text, 2)
    txtTM_TP.Text = FormatCurrency(txtTM_TP.Text, 2)
    txtTM_TRS.Text = FormatCurrency(txtTM_TRS.Text, 2)
    txtTM_HrsDibyr.Text = FormatCurrency(txtTM_HrsDibyr.Text, 2)

    mcurTM_TBP = txtTM_TBP.Text
    mcurTM_TP = txtTM_TP.Text
    mcurTM_TRS = txtTM_TRS.Text
    mcurTM_Pemb = 0
    mcurTM_HrsDibyr = txtTM_HrsDibyr.Text
    mcurTM_JmlByr = 0: mcurTM_ST = 0
    mcurTM_HrsDibyrNow = mcurTM_HrsDibyr

    txtOA_TBP.Text = FormatCurrency(txtOA_TBP.Text, 2)
    txtOA_TP.Text = FormatCurrency(txtOA_TP.Text, 2)
    txtOA_TRS.Text = FormatCurrency(txtOA_TRS.Text, 2)
    txtOA_HrsDibyr.Text = FormatCurrency(txtOA_HrsDibyr.Text, 2)

    mcurOA_TBP = txtOA_TBP.Text
    mcurOA_TP = txtOA_TP.Text
    mcurOA_TRS = txtOA_TRS.Text
    mcurOA_Pemb = 0
    mcurOA_HrsDibyr = txtOA_HrsDibyr.Text
    mcurOA_JmlByr = 0: mcurOA_ST = 0
    mcurOA_HrsDibyrNow = mcurOA_HrsDibyr

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadText()
    Dim i As Integer
    txtIsi.Left = hgTagihanPasien.Left
    Select Case hgTagihanPasien.Col
        Case 21, 22
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

Private Function sp_DetailBiayaPelayanan4PasienNU(f_NoPendaftaran As String, f_KdRuangan As String, f_KdItem As String, _
    f_tglPelayanan As Date, f_KdAsal As String, f_HutangPenjamin As Currency, f_TanggunganRS As Currency, f_TotalPembebasan As Currency, f_Jenis As String, f_Satuan As String) As Boolean
    On Error GoTo errLoad
    sp_DetailBiayaPelayanan4PasienNU = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, adInteger, Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("Kode_Item", adVarChar, adParamInput, 9, f_KdItem)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_tglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("JmlHutangPenjamin", adCurrency, adParamInput, , f_HutangPenjamin)
        .Parameters.Append .CreateParameter("JmlTanggunganRS", adCurrency, adParamInput, , f_TanggunganRS)
        .Parameters.Append .CreateParameter("JmlPembebasan", adCurrency, adParamInput, , f_TotalPembebasan)
        .Parameters.Append .CreateParameter("Jenis", adChar, adParamInput, 2, f_Jenis)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, f_Satuan)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "T")
        .ActiveConnection = dbConn
        .CommandText = "Update_DetailBiayaPelayanan4PasienNUNew"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_DetailBiayaPelayanan4PasienNU = False
        Else
            Call Add_HistoryLoginActivity("Update_DetailBiayaPelayanan4PasienNUNew")
        End If
    End With

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
hell:
End Sub

