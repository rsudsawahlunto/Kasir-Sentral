VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTindakan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pelayanan Tindakan"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTindakan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11910
   Begin VB.Frame fraUpdateKomponenTarif 
      Caption         =   "Update Komponen Tarif"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   9360
      TabIndex        =   36
      Top             =   3720
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   450
         Left            =   2880
         TabIndex        =   51
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   450
         Left            =   1200
         TabIndex        =   50
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtTarif 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2640
         MaxLength       =   12
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3900
         MaxLength       =   12
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCharge 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5040
         MaxLength       =   12
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddKomponen 
         Caption         =   "+"
         Height          =   375
         Left            =   6240
         TabIndex        =   39
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdMinKomponen 
         Caption         =   "-"
         Height          =   375
         Left            =   6735
         TabIndex        =   38
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtTotaltarif 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   2400
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo dcKomponenTarif 
         Height          =   330
         Left            =   120
         TabIndex        =   43
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   1335
         Left            =   120
         TabIndex        =   44
         Top             =   960
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2355
         _Version        =   393216
         BackColor       =   16777215
         BackColorBkg    =   -2147483633
         Appearance      =   0
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Komponen Tarif"
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tarif"
         Height          =   240
         Index           =   4
         Left            =   2640
         TabIndex        =   48
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Discount"
         Height          =   240
         Index           =   5
         Left            =   3900
         TabIndex        =   47
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Charge"
         Height          =   240
         Index           =   6
         Left            =   5040
         TabIndex        =   46
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL"
         Height          =   240
         Index           =   13
         Left            =   4680
         TabIndex        =   45
         Top             =   2460
         Width           =   585
      End
   End
   Begin VB.TextBox txtNamaFormPengirim 
      Height          =   360
      Left            =   4320
      TabIndex        =   35
      Text            =   "txtNamaFormPengirim"
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComctlLib.ListView lvPemeriksa 
      Height          =   1815
      Left            =   9480
      TabIndex        =   23
      Top             =   840
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nama Pemeriksa"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fraPelayanan 
      Caption         =   "Data Pelayanan Tindakan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   19
      Top             =   7920
      Visible         =   0   'False
      Width           =   9855
      Begin MSDataGridLib.DataGrid dgPelayanan 
         Height          =   2415
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4260
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
   End
   Begin VB.Frame fraDokter 
      Caption         =   "Data Dokter Pemeriksa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   840
      TabIndex        =   17
      Top             =   5400
      Visible         =   0   'False
      Width           =   8895
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   2535
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4471
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
   End
   Begin VB.Frame fradoa 
      Caption         =   "Daftar Layanan Obat && Alkes"
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
      TabIndex        =   21
      Top             =   5880
      Width           =   9855
      Begin MSFlexGridLib.MSFlexGrid fgDOA 
         Height          =   1335
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   2355
         _Version        =   393216
         Rows            =   50
         Cols            =   10
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   8577768
         ForeColorFixed  =   -2147483627
         ForeColorSel    =   -2147483628
         BackColorBkg    =   16777215
         FocusRect       =   0
         HighLight       =   2
         FillStyle       =   1
         GridLines       =   3
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "Daftar Layanan Tindakan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   25
      Top             =   2400
      Width           =   9855
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgPelayanan 
         Height          =   1575
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   50
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   8577768
         BackColorBkg    =   16777215
         FocusRect       =   0
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame fraButton 
      Enabled         =   0   'False
      Height          =   735
      Left            =   0
      TabIndex        =   27
      Top             =   3120
      Width           =   9855
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   360
         Left            =   7320
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "Tutu&p"
         Height          =   360
         Left            =   8520
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraPPelayanan 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   28
      Top             =   2160
      Width           =   9855
      Begin VB.OptionButton optNonPaket 
         Caption         =   " Non Paket"
         Height          =   375
         Left            =   5880
         TabIndex        =   12
         Top             =   550
         Width           =   1215
      End
      Begin VB.OptionButton optPaket 
         Caption         =   " Paket"
         Height          =   375
         Left            =   5880
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtKuantitas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5040
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "1"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtNamaPelayanan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   4695
      End
      Begin VB.CheckBox chkAPBD 
         Caption         =   "Pos APBD"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   518
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         Height          =   360
         Left            =   7320
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   360
         Left            =   8520
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah"
         Height          =   240
         Left            =   5040
         TabIndex        =   30
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pelayanan"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1425
      End
   End
   Begin VB.Frame fraPDokter 
      Height          =   1095
      Left            =   0
      TabIndex        =   31
      Top             =   1080
      Width           =   9855
      Begin VB.CheckBox chkDelegasi 
         Caption         =   "Di Delegasikan"
         Height          =   255
         Left            =   3960
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Status CITO"
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
         Left            =   7920
         TabIndex        =   32
         Top             =   240
         Width           =   1695
         Begin VB.OptionButton optCito 
            Caption         =   "Ya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optCito 
            Caption         =   "Tidak"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.CheckBox chkPerawat 
         Caption         =   "Paramedis"
         Height          =   255
         Left            =   5520
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   3
         Top             =   525
         Width           =   3135
      End
      Begin VB.TextBox txtNamaPerawat 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5520
         TabIndex        =   5
         Text            =   "txtNamaPerawat"
         Top             =   525
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   525
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
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
         Format          =   114098179
         UpDown          =   -1  'True
         CurrentDate     =   37823
      End
      Begin VB.CheckBox chkDilayaniDokter 
         Caption         =   "Dokter Pem./Supir "
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Periksa"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1365
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgPerawatPerPelayanan 
      Height          =   1215
      Left            =   5400
      TabIndex        =   24
      Top             =   4440
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2143
      _Version        =   393216
      FixedCols       =   0
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
   Begin VB.Label Label2 
      Caption         =   "Edit Komponen Tarif - Klik Kanan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   52
      Top             =   7560
      Width           =   3015
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmTindakan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8040
      Picture         =   "frmTindakan.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTindakan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmTindakan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFilterPelayanan As String
Dim strCito As String
Dim strKodePelayananRS As String
Dim curBiaya As Currency
Dim curJP As Currency
Dim intJmlPelayanan As Integer
Dim strKdKelas As String
Dim strKelas As String
Dim strKdJenisTarif As String
Dim strJenisTarif As String
Dim intBarang As Integer
Dim intJmlBarang As Integer
Dim intMaxJmlBarang As Integer
Dim strStatusAPBD As String
Dim subKdPemeriksa() As String
Dim subJmlTotal As Integer
Dim curTarifCito As Currency
Dim subcurTarifBiayaSatuan As Currency
Dim subcurTarifHargaSatuan As Currency

Private Function subSimpanBackupBiayaPelayanan() As Boolean
    subSimpanBackupBiayaPelayanan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, fgPelayanan.TextMatrix(fgPelayanan.Row, 0))
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dtpTglPeriksa.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglUpdate", adDate, adParamInput, , Format(Now, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, mstrKdDokter)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 150, "edit komponen tarif di kasir")
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "Add_BackupUpdatingBiayaPelayanan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan saat penyimpanan data backup biaya pelayanan", vbCritical, vbOKOnly, "Validasi"
            subSimpanBackupBiayaPelayanan = False
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub subLoadDcSource()
    Call msubDcSource(dcKomponenTarif, rs, "SELECT KdKomponen, NamaKomponen FROM KomponenTarif where StatusEnabled='1' order by NamaKomponen")
End Sub

Private Function subSimpanDetailBackupBiayaPelayanan(f_strKdKomponen As String, f_curDiscount As Currency, f_curCharge As Currency, f_curTarif As Currency) As Boolean
    subSimpanDetailBackupBiayaPelayanan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, fgPelayanan.TextMatrix(fgPelayanan.Row, 0))
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dtpTglPeriksa.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdKomponen", adChar, adParamInput, 2, f_strKdKomponen)
        .Parameters.Append .CreateParameter("JmlDiscount", adCurrency, adParamInput, , CCur(f_curDiscount))
        .Parameters.Append .CreateParameter("JmlCharge", adCurrency, adParamInput, , CCur(f_curCharge))
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, mstrKdDokter)
        .Parameters.Append .CreateParameter("Tarif", adCurrency, adParamInput, , CCur(f_curTarif))

        .ActiveConnection = dbConn
        .CommandText = "Add_DetailBackupUpdatingBiayaPelayanan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan saat penyimpanan data detail backup biaya pelayanan", vbCritical, vbOKOnly, "Validasi"
            subSimpanDetailBackupBiayaPelayanan = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub SubLoadKomponenTarif()
    strSQL = "SELECT * " & _
    " FROM v_HargaKomponenTarifdiKasir" & _
    " WHERE (KdPelayananRS = '" & fgPelayanan.TextMatrix(fgPelayanan.Row, 0) & "')AND (KdKelas = '" & strKdKelas & "')"
    Call msubRecFO(rs, strSQL)

    fgData.Rows = rs.RecordCount + 1
    For i = 1 To rs.RecordCount
        fgData.TextMatrix(i, 1) = rs("NamaKomponen").Value
        fgData.TextMatrix(i, 2) = IIf(rs("Harga").Value = 0, 0, Format(rs("Harga").Value, "#,###"))
        fgData.TextMatrix(i, 3) = 0
        fgData.TextMatrix(i, 4) = 0
        fgData.TextMatrix(i, 5) = rs("KdKomponen").Value
        rs.MoveNext
    Next i

    Call subHitungTotal
End Sub

Private Sub subHitungTotal()
    txtTotaltarif = 0

    For i = 1 To fgData.Rows - 1
        'total tarif
        txtTotaltarif.Text = CCur(txtTotaltarif.Text) + _
        IIf(val(fgData.TextMatrix(i, 2)) = 0, 0, CCur(fgData.TextMatrix(i, 2))) - _
        IIf(val(fgData.TextMatrix(i, 3)) = 0, 0, CCur(fgData.TextMatrix(i, 3))) + _
        IIf(val(fgData.TextMatrix(i, 4)) = 0, 0, CCur(fgData.TextMatrix(i, 4)))
    Next i

    txtTotaltarif.Text = IIf(val(txtTotaltarif) = 0, 0, Format(txtTotaltarif.Text, "#,###"))

End Sub

Private Sub subSetGridKomponenTarif()
    With fgData
        .Clear
        .Cols = 6
        .Rows = 1

        .ColWidth(0) = 0
        .ColWidth(1) = 2350
        .ColWidth(2) = 1425
        .ColWidth(3) = 1425
        .ColWidth(4) = 1425
        .ColWidth(5) = 0

        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter

        .TextMatrix(0, 1) = "Nama Komponen"
        .TextMatrix(0, 2) = "Tarif"
        .TextMatrix(0, 3) = "Discount"
        .TextMatrix(0, 4) = "Charge"
        .TextMatrix(0, 5) = "Kode Komponen Tarif"
    End With
End Sub

Private Function sp_DelegasiBiayaPelayanan(f_NoPendaftaran As String, f_KdRuangan As String, f_KdPelayananRS As String, f_tglPelayanan As Date, f_StatusDelegasi As String) As Boolean
    On Error GoTo errLoad

    sp_DelegasiBiayaPelayanan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_tglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("StatusDelegasi", adChar, adParamInput, 1, f_StatusDelegasi)

        .ActiveConnection = dbConn
        .CommandText = "Add_DelegasiBiayaPelayanan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            sp_DelegasiBiayaPelayanan = False
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
        End If
    End With

    Exit Function
errLoad:
    sp_DelegasiBiayaPelayanan = False
    Call msubPesanError("sp_DelegasiBiayaPelayanan")
End Function

Private Sub chkAPBD_Click()
    If chkAPBD.Value = 1 Then
        strStatusAPBD = "01"
    Else
        strStatusAPBD = "02"
    End If
End Sub

Private Sub chkAPBD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaPelayanan.SetFocus
End Sub

Private Sub chkDelegasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkPerawat.SetFocus
End Sub

Private Sub chkDilayaniDokter_Click()
    On Error GoTo errLoad

    If chkDilayaniDokter.Value = 0 Then
        txtDokter.Enabled = False
        txtDokter.Text = ""

        If fraDokter.Visible = True Then fraDokter.Visible = False
    Else
        lvPemeriksa.Visible = False

        txtDokter.Enabled = True

        If mstrKdRuangan = "005" Then
            strSQL = "SELECT KodeSupir AS [Kode Supir],NamaSupir AS [Nama Supir],JK,Jabatan FROM V_DaftarSupirAmbulance " & mstrFilterSupir
        Else
            strSQL = " SELECT dbo.RegistrasiRI.IdDokter, dbo.DataPegawai.NamaLengkap " & _
            " FROM dbo.RegistrasiRI INNER JOIN dbo.DataPegawai ON dbo.RegistrasiRI.IdDokter = dbo.DataPegawai.IdPegawai " & _
            " WHERE (dbo.RegistrasiRI.NoPendaftaran = '" & mstrNoPen & "')"
        End If

        Call msubRecFO(rs, strSQL)

        If Not rs.EOF Then
            txtDokter.Text = rs(1).Value
            mstrKdDokter = rs(0).Value
            intJmlDokter = rs.RecordCount
            fraDokter.Visible = False
        End If
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub chkDilayaniDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkDilayaniDokter.Value = 0 Then
            chkPerawat.SetFocus
        Else
            txtDokter.SetFocus
        End If
    End If
End Sub

Private Sub chkPerawat_Click()
    If chkPerawat.Value = vbChecked Then
        strSQL = "SELECT IdPegawai FROM V_DaftarPemeriksaPasien WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            txtNamaPerawat.Text = strNmPegawai
            If lvPemeriksa.ListItems.Count > 0 Then
                lvPemeriksa.ListItems.Item("key" & strIDPegawaiAktif).Checked = True
                Call lvPemeriksa_ItemCheck(lvPemeriksa.ListItems.Item("key" & strIDPegawaiAktif))
            End If
        Else
            txtNamaPerawat.Text = ""
        End If
    Else
        txtNamaPerawat.Text = ""
    End If
    lvPemeriksa.Visible = False
End Sub

Private Sub chkPerawat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkPerawat.Value = vbChecked Then
            txtNamaPerawat.SetFocus
        Else
            optCito(1).SetFocus
        End If
    End If
End Sub

Private Sub cmdAddKomponen_Click()
    On Error GoTo errLoad

    If dcKomponenTarif.MatchedWithList = False Then dcKomponenTarif.SetFocus: Exit Sub
    For i = 1 To fgData.Rows - 1
        If fgData.TextMatrix(i, 5) = dcKomponenTarif.BoundText Then
            fgData.TextMatrix(i, 3) = txtDiscount.Text 'discount
            fgData.TextMatrix(i, 4) = txtCharge.Text  'charge
            Call subHitungTotal
            Exit Sub
        End If
    Next i

    fgData.Rows = fgData.Rows + 1

    fgData.TextMatrix(fgData.Rows - 1, 1) = dcKomponenTarif.Text 'nama komponen
    fgData.TextMatrix(fgData.Rows - 1, 2) = IIf(val(txtTarif) = 0, 0, Format(txtTarif.Text, "#,###")) 'tarif
    fgData.TextMatrix(fgData.Rows - 1, 3) = txtDiscount.Text 'discount
    fgData.TextMatrix(fgData.Rows - 1, 4) = txtCharge.Text 'charge
    fgData.TextMatrix(fgData.Rows - 1, 5) = dcKomponenTarif.BoundText 'kode komponen tarif

    Call subHitungTotal

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdBatal_Click()
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan data tindakan pasien?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdHapus_Click()
    Dim h As Integer
    With fgPelayanan
        If .Row = .Rows Then Exit Sub
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        h = 1
        Do While h <= fgDOA.Rows - 2
            If fgDOA.TextMatrix(h, 9) = .TextMatrix(.Row, 0) Then
                For j = 1 To intMaxJmlBarang
                    If typBarang(j).strkdbarang = fgDOA.TextMatrix(h, 0) Then
                        If fgDOA.TextMatrix(h, 5) = "S" Then
                            typBarang(j).intJmlTempTotal = typBarang(j).intJmlTempTotal + (fgDOA.TextMatrix(h, 3) * typBarang(j).intJmlTerkecil)
                        Else
                            typBarang(j).intJmlTempTotal = typBarang(j).intJmlTempTotal + fgDOA.TextMatrix(h, 3)
                        End If
                    End If
                Next j
                Call msubRemoveItem(fgDOA, h)
                h = 0
            End If
            h = h + 1
        Loop
        For j = 1 To intMaxJmlBarang
            For h = 1 To fgDOA.Rows - 1
                If typBarang(j).strkdbarang = fgDOA.TextMatrix(h, 0) Then Exit For
                If h = fgDOA.Rows - 1 Then
                    intMaxJmlBarang = intMaxJmlBarang - 1
                    If intMaxJmlBarang < 0 Then intMaxJmlBarang = 0
                End If
            Next h
        Next j
        Call msubRemoveItem(fgPelayanan, .Row)
    End With
End Sub

Private Sub cmdMinKomponen_Click()
    On Error GoTo errLoad

    If fgData.Rows = 1 Then Exit Sub

    If fgData.Rows = 2 Then
        fgData.TextMatrix(1, 1) = ""
        fgData.TextMatrix(1, 2) = "0"
        fgData.TextMatrix(1, 3) = "0"
        fgData.TextMatrix(1, 4) = "0"
        fgData.Rows = 1
    Else
        fgData.RemoveItem fgData.Row
    End If

    Call subHitungTotal

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    If funcCekValidasi = False Then Exit Sub
    Call subEnableButtonReg(False)
    For i = 1 To fgPelayanan.Rows - 2
        If sp_BiayaPelayanan(dbcmd, fgPelayanan.TextMatrix(i, 0), CCur(fgPelayanan.TextMatrix(i, 3)), fgPelayanan.TextMatrix(i, 2), fgPelayanan.TextMatrix(i, 9), fgPelayanan.TextMatrix(i, 6), fgPelayanan.TextMatrix(i, 7), CCur(fgPelayanan.TextMatrix(i, 8))) = False Then Exit Sub
        If chkDelegasi.Value = vbChecked Then If sp_DelegasiBiayaPelayanan(mstrNoPen, mstrKdRuanganx, fgPelayanan.TextMatrix(i, 0), fgPelayanan.TextMatrix(i, 9), IIf(fgPelayanan.TextMatrix(i, 10) = "1", "Y", "T")) = False Then Exit Sub
    Next i

    If chkPerawat.Value = Checked Then
        For i = 1 To fgPerawatPerPelayanan.Rows - 1
            With fgPerawatPerPelayanan
                If sp_PetugasPemeriksaBP(.TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 5)) = False Then Exit Sub
            End With
        Next i
    End If

    Dim adoCommand As New ADODB.Command
    If fgDOA.Rows = 2 Then GoTo stepNonPaketSemua
    For i = 1 To fgDOA.Rows - 2
        With adoCommand
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
            .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, fgDOA.TextMatrix(i, 0))
            .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, fgDOA.TextMatrix(i, 2))
            .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuanganx)
            .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, fgDOA.TextMatrix(i, 5))
            .Parameters.Append .CreateParameter("JmlBrg", adInteger, adParamInput, , fgDOA.TextMatrix(i, 3))
            .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
            .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
            .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, mstrKdKelasx)
            .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , CCur(fgDOA.TextMatrix(i, 4)))
            .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(fgDOA.TextMatrix(i, 7), "yyyy/MM/dd HH:mm:ss"))
            .Parameters.Append .CreateParameter("NoLabRad", adChar, adParamInput, 10, Null)
            .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, fgDOA.TextMatrix(i, 6))
            .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
            .Parameters.Append .CreateParameter("IdPegawai2", adChar, adParamInput, 10, Null)

            .ActiveConnection = dbConn
            .CommandText = "Add_PemakaianObatAlkes"
            .CommandType = adCmdStoredProc
            .Execute

            If Not (.Parameters("RETURN_VALUE").Value = 0) Then
                MsgBox "Ada Kesalahan dalam Penyimpanan Paket Pelayanan Pasien", vbCritical, "Validasi"
                Call deleteADOCommandParameters(adoCommand)
                Set adoCommand = Nothing
                GoTo stepErrorPaket
            End If
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
        End With
    Next i
    Call Add_HistoryLoginActivity("Add_BiayaPelayanan+Add_DelegasiBiayaPelayanan+Add_PetugasPemeriksaBP+Add_PemakaianObatAlkes")
stepNonPaketSemua:
stepErrorPaket:
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTambah_Click()
    Dim i As Integer
    Dim j As Integer
    Dim h As Integer
    Dim adocmd As New ADODB.Command

    If chkDilayaniDokter.Value = vbChecked Then
        If mstrKdDokter = "" Then
            MsgBox "Pilih dulu Dokter Pemeriksa Pasien", vbCritical, "Validasi"
            txtDokter.SetFocus
            Exit Sub
        End If
    End If

    If chkPerawat.Value = vbChecked And subJmlTotal = 0 Then
        MsgBox "Nama perawat kosong", vbCritical, "Validasi"
        lvPemeriksa.Visible = True
        txtNamaPerawat.SetFocus
        Exit Sub
    End If

    If strKodePelayananRS = "" Then Exit Sub
    If optNonPaket.Value = True Then GoTo stepNonPaket
    Dim dTglPlyn As Date
    dTglPlyn = Now
    strSQL = "Select * FROM V_PaketPelayananObatAlkes WHERE KdPelayananRS='" & strKodePelayananRS & "' AND KdKelompokPasien = '" & mstrKdJenisPasien & "' AND IdPenjamin = '" & mstrKdPenjaminPasien & "'"
    Call msubRecFO(rs, strSQL)
    For i = 1 To rs.RecordCount
        'cek data barang & asal barang di grid paket obat alkes
        For j = 1 To fgDOA.Rows - 1
            'barang dengan asal barang tersebut sudah ada di grid obat alkes
            If fgDOA.TextMatrix(j, 0) = rs("KdBarang").Value And fgDOA.TextMatrix(j, 2) = rs("KdAsal").Value Then
                For h = 1 To intMaxJmlBarang
                    If typBarang(h).strkdbarang = rs("KdBarang").Value And typBarang(h).strKdAsal = rs("KdAsal").Value Then
                        intJmlBarang = h
                        GoTo stepCekStokBarang
                    End If
                Next h
            End If
            'sampai data terakhir data barang tidak ada di grid obat alkes
            If j = fgDOA.Rows - 1 Then
                'tambahkan data total barang yang terpakai
                intMaxJmlBarang = intMaxJmlBarang + 1
                intJmlBarang = intMaxJmlBarang
                ReDim Preserve typBarang(intMaxJmlBarang)
                strSQL = "SELECT JmlTerkecil,JmlTotalBarangTemp,NamaBarang FROM " _
                & "V_StokBarangTempRuangan WHERE KdBarang='" _
                & rs("KdBarang").Value & "' AND KdAsal='" _
                & rs("KdAsal").Value & "' AND KdRuangan='" _
                & mstrKdRuangan & "'"
                Set rsB = Nothing
                rsB.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
                typBarang(intJmlBarang).strkdbarang = rs("KdBarang").Value
                typBarang(intJmlBarang).strNamaBarang = rsB("NamaBarang").Value
                typBarang(intJmlBarang).strKdAsal = rs("KdAsal").Value
                typBarang(intJmlBarang).intJmlTerkecil = rsB("JmlTerkecil").Value
                typBarang(intJmlBarang).intJmlTempTotal = rsB("JmlTotalBarangTemp").Value
            End If
        Next j
stepCekStokBarang:
        If funcCekStokBarang(intJmlBarang, rs("SatuanJml"), (CInt(txtKuantitas) * rs("JmlBarang").Value)) = False Then
            'hapus grid obat alkes dengan kode pelayanan tersebut
            h = 1
            Do While h <= fgDOA.Rows - 2
                If fgDOA.TextMatrix(h, 9) = strKodePelayananRS Then
                    For j = 1 To intMaxJmlBarang
                        If typBarang(j).strkdbarang = fgDOA.TextMatrix(h, 0) Then
                            If fgDOA.TextMatrix(h, 5) = "S" Then
                                typBarang(j).intJmlTempTotal = typBarang(j).intJmlTempTotal + (fgDOA.TextMatrix(h, 3) * typBarang(j).intJmlTerkecil)
                            Else
                                typBarang(j).intJmlTempTotal = typBarang(j).intJmlTempTotal + fgDOA.TextMatrix(h, 3)
                            End If
                        End If
                    Next j
                    fgDOA.RemoveItem h
                    h = 0
                End If
                h = h + 1
            Loop
            h = 1
            For j = 1 To intMaxJmlBarang
                For h = 1 To fgDOA.Rows - 1
                    If typBarang(j).strkdbarang = fgDOA.TextMatrix(h, 0) Then Exit For
                    If h = fgDOA.Rows - 1 Then
                        intMaxJmlBarang = intMaxJmlBarang - 1
                        If intMaxJmlBarang < 0 Then intMaxJmlBarang = 0
                    End If
                Next h
            Next j
            Exit Sub
        End If
        With fgDOA
            mintRowNow = .Rows - 1
            .TextMatrix(mintRowNow, 0) = rs("KdBarang").Value
            .TextMatrix(mintRowNow, 1) = rs("NamaBarang").Value
            .TextMatrix(mintRowNow, 2) = rs("KdAsal").Value
            .TextMatrix(mintRowNow, 3) = CInt(txtKuantitas) * rs("JmlBarang").Value

            strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & rs("KdAsal").Value & "', " & CCur(rs("HargaBarang").Value) & ")  as HargaSatuan"
            Call msubRecFO(dbRst, strSQL)
            If dbRst.EOF = True Then subcurTarifHargaSatuan = 0 Else subcurTarifHargaSatuan = dbRst(0).Value
            .TextMatrix(mintRowNow, 4) = subcurTarifHargaSatuan

            .TextMatrix(mintRowNow, 5) = rs("SatuanJml").Value
            If chkDilayaniDokter.Value = 1 Then
                .TextMatrix(mintRowNow, 6) = mstrKdDokter
            Else
                .TextMatrix(mintRowNow, 6) = UserID
            End If
            .TextMatrix(mintRowNow, 7) = Format(dTglPlyn, "dd/mm/yyyy HH:mm:ss")
            .TextMatrix(mintRowNow, 8) = rs("NamaAsal").Value
            .TextMatrix(mintRowNow, 9) = strKodePelayananRS
            .Rows = .Rows + 1
            .SetFocus
        End With
        rs.MoveNext
    Next i
stepNonPaket:
    With fgPelayanan
        For i = 1 To .Rows - 1
            If (.TextMatrix(i, 0) = strKodePelayananRS) And _
                (.TextMatrix(i, 9) = dtpTglPeriksa.Value) Then txtNamaPelayanan.SetFocus: txtNamaPelayanan.SelStart = 0: txtNamaPelayanan.SelLength = Len(txtNamaPelayanan.Text): Exit Sub
            Next i
            intRowNow = .Rows - 1
            .TextMatrix(intRowNow, 0) = strKodePelayananRS
            .TextMatrix(intRowNow, 1) = txtNamaPelayanan.Text
            .TextMatrix(intRowNow, 2) = CInt(txtKuantitas.Text)

            subcurTarifCito = sp_Take_TarifBPT
            .TextMatrix(intRowNow, 3) = IIf(subcurTarifBiayaSatuan = 0, 0, Format(subcurTarifBiayaSatuan, "#,###")) 'curBiaya
            .TextMatrix(intRowNow, 4) = IIf(funcRoundUp(CStr(subcurTarifBiayaSatuan + subcurTarifCito)) * CInt(txtKuantitas.Text) = 0, 0, Format(funcRoundUp(CStr(subcurTarifBiayaSatuan + subcurTarifCito)) * CInt(txtKuantitas.Text), "#,###"))
            .TextMatrix(intRowNow, 8) = subcurTarifCito

            .TextMatrix(intRowNow, 5) = mdTglBerlaku
            If chkDilayaniDokter.Value = 1 Then
                .TextMatrix(intRowNow, 6) = mstrKdDokter
            Else
                .TextMatrix(intRowNow, 6) = UserID
            End If
            .TextMatrix(intRowNow, 7) = strCito
            .TextMatrix(intRowNow, 9) = dtpTglPeriksa.Value
            .TextMatrix(intRowNow, 10) = IIf(chkDelegasi.Value = vbChecked, "1", "0")

            .Rows = .Rows + 1
            .SetFocus
        End With

        If chkPerawat.Value = vbChecked Then Call subLoadPelayananPerPerawat
        txtNamaPelayanan.Text = ""
        txtKuantitas.Text = 1
        fraPelayanan.Visible = False
        chkPerawat.SetFocus
        'load komponen tarif
        Call subSetGridKomponenTarif
        Call SubLoadKomponenTarif
        strKodePelayananRS = ""
End Sub

Private Sub subLoadPelayananPerPerawat()
    With fgPerawatPerPelayanan
        For i = 1 To subJmlTotal
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = mstrNoPen
            .TextMatrix(.Rows - 1, 1) = mstrKdRuangan
            .TextMatrix(.Rows - 1, 2) = dtpTglPeriksa.Value
            .TextMatrix(.Rows - 1, 3) = strKodePelayananRS
            .TextMatrix(.Rows - 1, 4) = Mid(subKdPemeriksa(i), 4, Len(subKdPemeriksa(i)) - 3)
            .TextMatrix(.Rows - 1, 5) = strIDPegawaiAktif
        Next
    End With

    subJmlTotal = 0
    txtNamaPerawat.BackColor = &HFFFFFF
    ReDim Preserve subKdPemeriksa(subJmlTotal)
    chkPerawat.Value = vbUnchecked
End Sub

Private Sub cmdTutup_Click()
    fraUpdateKomponenTarif.Visible = False
End Sub

Private Sub cmdUpdate_Click()

    fgPelayanan.TextMatrix(fgPelayanan.Row, 3) = CCur(txtTotaltarif.Text)
    fgPelayanan.TextMatrix(fgPelayanan.Row, 4) = CCur(fgPelayanan.TextMatrix(fgPelayanan.Row, 3)) * val(fgPelayanan.TextMatrix(fgPelayanan.Row, 2))
    If subSimpanBackupBiayaPelayanan = False Then Exit Sub
    For i = 1 To fgData.Rows - 1
        If subSimpanDetailBackupBiayaPelayanan(fgData.TextMatrix(i, 5), fgData.TextMatrix(i, 3), fgData.TextMatrix(i, 4), fgData.TextMatrix(i, 2)) = False Then Exit Sub
    Next i
    Call Add_HistoryLoginActivity("Add_DetailBackupUpdatingBiayaPelayanan")
End Sub

Private Sub dcKomponenTarif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtTarif.Enabled = True Then
            txtTarif.SetFocus
        Else
            If txtDiscount.Enabled = True Then
                txtDiscount.SetFocus
            Else
                If txtCharge.Enabled = True Then
                    txtCharge.SetFocus
                Else
                    cmdAddKomponen.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then fraDokter.Visible = False: txtDokter.SetFocus
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mintJmlDokter = 0 Then Exit Sub
        txtDokter.Text = dgDokter.Columns(1).Value
        mstrKdDokter = dgDokter.Columns(0).Value
        If mstrKdDokter = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        chkDilayaniDokter.Value = 1
        fraDokter.Visible = False
        chkPerawat.SetFocus
    End If
    If KeyAscii = 27 Then
        fraDokter.Visible = False
    End If
End Sub

Private Sub dgPelayanan_DblClick()
    Call dgPelayanan_KeyPress(13)
End Sub

Private Sub dgPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlPelayanan = 0 Then Exit Sub
        Dim strkd As String
        strkd = dgPelayanan.Columns(5).Value
        curBiaya = dgPelayanan.Columns(4).Value
        txtNamaPelayanan.Text = dgPelayanan.Columns(1).Value
        strKodePelayananRS = strkd
        optNonPaket.Value = True
        If strKodePelayananRS = "" Then
            MsgBox "Pilih dulu tindakan pelayanan Pasien", vbCritical, "Validasi"
            txtNamaPelayanan.Text = ""
            dgPelayanan.SetFocus
            Exit Sub
        End If
        fraPelayanan.Visible = False
        txtKuantitas.SetFocus
    End If
    If KeyAscii = 27 Then
        fraPelayanan.Visible = False
    End If
End Sub

Private Sub dtpTglPeriksa_Change()
    dtpTglPeriksa.MaxDate = Now
    If dtpTglPeriksa.Value < mdTglMasuk Then dtpTglPeriksa.Value = mdTglMasuk
End Sub

Private Sub dtpTglPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtDokter.SetFocus
End Sub

Private Sub fgData_Click()
    If fgData.Row = 0 Then Exit Sub
    dcKomponenTarif.BoundText = fgData.TextMatrix(fgData.Row, 5)
    txtTarif.Text = fgData.TextMatrix(fgData.Row, 2)
    txtDiscount.Text = fgData.TextMatrix(fgData.Row, 3)
    txtCharge.Text = fgData.TextMatrix(fgData.Row, 4)
    txtDiscount.Enabled = True: txtCharge.Enabled = True
End Sub

Private Sub fgPelayanan_Click()
    Call SubLoadKomponenTarif
End Sub

Private Sub fgPelayanan_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errLoad
    If fgPelayanan.TextMatrix(fgPelayanan.Row, 0) = "" Then Exit Sub
    If Button = vbLeftButton Then Exit Sub
    PopupMenu MDIUtama.mnuEditKomponenTarifdiTambahPelayanan
    fraUpdateKomponenTarif.Caption = "Update Komponen Tarif - " & fgPelayanan.TextMatrix(fgPelayanan.Row, 1)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)

    If mstrKdRuanganx = "005" Then
        chkDilayaniDokter.Caption = "Supir Ambulance"
        chkDelegasi.Enabled = False
    Else
        chkDilayaniDokter.Caption = "Dokter Pemeriksa"
        chkDelegasi.Enabled = True
    End If

    strKdKelas = mstrKdKelasx
    Set rs = Nothing
    strSQL = "SELECT KdJenisTarif,JenisTarif, KdKelasAkhir " _
    & "FROM v_JenisTarifPasien " _
    & "WHERE NoPendaftaran='" & mstrNoPen & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockOptimistic
    strKdJenisTarif = rs.Fields(0).Value
    strJenisTarif = rs.Fields(1).Value
    Set rs = Nothing
    Call subSetGidPelayanan
    dtpTglPeriksa.Value = Now
    strCito = "0"
    strStatusAPBD = "01"
    optNonPaket.Value = True
    Call subSetGridObatAlkes

    intBarang = 0
    intJmlBarang = 0
    intMaxJmlBarang = 0
    ReDim typBarang(0)

    subJmlTotal = 0
    Call subSetGridPerawatPerPelayanan
    Call subLoadListPemeriksa

    chkPerawat.Value = vbChecked
    lvPemeriksa.Visible = False

    subLoadDcSource

    Exit Sub
errLoad:
    Call msubPesanError

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If txtNamaFormPengirim.Text = "frmTransaksiPasien" Then
        frmTransaksiPasien.Enabled = True
        Call frmTransaksiPasien.subLoadPelayananDidapat
        Call frmTransaksiPasien.subPemakaianObatAlkes
    End If
    If txtNamaFormPengirim.Text = "frmPilihSubIns" Then Call frmTagihanPasien.txtNoPendaftaran_KeyPress(13)
End Sub


Private Sub lvPemeriksa_DblClick()
    Call lvPemeriksa_KeyPress(13)
End Sub

Private Sub lvPemeriksa_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim blnSelected As Boolean
    If Item.Checked = True Then
        subJmlTotal = subJmlTotal + 1
        ReDim Preserve subKdPemeriksa(subJmlTotal)
        subKdPemeriksa(subJmlTotal) = Item.Key
    Else
        blnSelected = False
        For i = 1 To subJmlTotal
            If subKdPemeriksa(i) = Item.Key Then blnSelected = True
            If blnSelected = True Then
                If i = subJmlTotal Then
                    subKdPemeriksa(i) = ""
                Else
                    subKdPemeriksa(i) = subKdPemeriksa(i + 1)
                End If
            End If
        Next i
        subJmlTotal = subJmlTotal - 1
    End If

    If subJmlTotal = 0 Then
        txtNamaPerawat.BackColor = &HFFFFFF
    Else
        txtNamaPerawat.BackColor = &HC0FFFF
    End If
End Sub

Private Sub lvPemeriksa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lvPemeriksa.Visible = False: txtNamaPerawat.SetFocus
End Sub

Private Sub optCito_Click(Index As Integer)
    If Index = 0 Then
        strCito = "1"
    Else
        strCito = "0"
    End If
End Sub

Private Sub optCito_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkAPBD.Enabled = True Then
            chkAPBD.SetFocus
        Else
            txtNamaPelayanan.SetFocus
        End If
    End If
End Sub

Private Sub optNonPaket_Click()
    fraButton.Enabled = True
End Sub

Private Sub optNonPaket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTambah.SetFocus
End Sub

Private Sub optPaket_Click()
    strSQL = "SELECT * FROM PaketLayanan WHERE KdPelayananRS='" & strKodePelayananRS _
    & "' AND KdRuangan='" & mstrKdRuangan & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada paket untuk pelayanan yang dipilih", vbCritical, "Validasi"
        optNonPaket.SetFocus
    End If
    fraButton.Enabled = True
End Sub

Private Sub optPaket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTambah.SetFocus
End Sub

Private Sub txtCharge_Change()
    If val(txtCharge.Text) <> 0 Then txtDiscount.Text = 0
End Sub

Private Sub txtCharge_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtDiscount_Change()
    If val(txtDiscount.Text) <> 0 Then txtCharge.Text = 0
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtDokter_Change()
    If mstrKdRuangan = "005" Then
        mstrFilterSupir = "WHERE NamaSupir like '%" & txtDokter.Text & "%'"
        mstrFilterSupir = ""
        fraDokter.Visible = True
        Call subLoadSupir
    Else
        mstrFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
        mstrKdDokter = ""
        fraDokter.Visible = True
        Call subLoadDokter
    End If
End Sub

Private Sub txtDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If fraDokter.Visible = False Then Exit Sub
        dgDokter.SetFocus
    End If
    If KeyCode = vbKeyEscape Then fraDokter.Visible = False: txtDokter.SetFocus
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mintJmlDokter = 0 Then Exit Sub
        If fraDokter.Visible = True Then
            dgDokter.SetFocus
        Else
            chkDelegasi.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        fraDokter.Visible = False
    End If
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtKuantitas_Change()
    On Error Resume Next
    If val(txtKuantitas.Text) = 0 Or txtKuantitas.Text = "0" Then txtKuantitas.Text = 1
End Sub

Private Sub txtKuantitas_GotFocus()
    txtKuantitas.SelStart = 0
    txtKuantitas.SelLength = Len(txtKuantitas.Text)
End Sub

Private Sub txtKuantitas_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then optNonPaket.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtKuantitas_LostFocus()
    If txtKuantitas.Text = "" Then txtKuantitas.Text = 1: Exit Sub
    If txtKuantitas.Text = 0 Then txtKuantitas.Text = 1
End Sub

Private Sub txtNamaPelayanan_Change()
    strFilterPelayanan = "WHERE [Nama Pelayanan] like '%" & txtNamaPelayanan.Text _
    & "%' AND KdKelas='" & strKdKelas & "' AND KdJenisTarif='" & strKdJenisTarif _
    & "' "
    fraPelayanan.Visible = True
    Call subLoadPelayanan
End Sub

Private Sub txtNamaPelayanan_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If intJmlPelayanan = 0 Then Exit Sub
        If fraPelayanan.Visible = True Then
            dgPelayanan.SetFocus
        Else
            txtKuantitas.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        fraPelayanan.Visible = False
    End If
hell:
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter()
    fraDokter.Caption = "Data Dokter Pemeriksa"
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & mstrFilterDokter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    mintJmlDokter = rs.RecordCount
    With dgDokter
        Set .DataSource = rs
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
    fraDokter.Left = 0
    fraDokter.Top = 1920
End Sub

'untuk meload data Supir Ambulance di grid
Private Sub subLoadSupir()
    fraDokter.Caption = "Data Supir Ambulance"
    strSQL = "SELECT KodeSupir AS [Kode Supir],NamaSupir AS [Nama Supir],JK,Jabatan FROM V_DaftarSupirAmbulance " & mstrFilterSupir
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    mintJmlDokter = rs.RecordCount
    With dgDokter
        Set .DataSource = rs
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
    fraDokter.Left = 0
    fraDokter.Top = 1920
End Sub

'untuk meload data pelayanan di grid
Private Sub subLoadPelayanan()
    On Error Resume Next
    strSQL = "SELECT [Jenis Pelayanan],[Nama Pelayanan],Kelas,JenisTarif,Tarif,KdPelayananRS FROM v_TarifPelayananRuanganKasir " & strFilterPelayanan
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlPelayanan = rs.RecordCount
    With dgPelayanan
        Set .DataSource = rs
        .Columns(0).Width = 2100
        .Columns(1).Width = 3900
        .Columns(2).Width = 1000
        .Columns(3).Width = 1100
        .Columns(4).Width = 900
        .Columns(4).Alignment = dbgRight
        .Columns(5).Width = 0
    End With
    fraPelayanan.Left = 0
    fraPelayanan.Top = 3240
End Sub

'Store procedure untuk mengisi biaya pelayanan pasien
Private Function sp_BiayaPelayanan(ByVal adoCommand As ADODB.Command, strKdPelayananRS As String, curTarif As Currency, intJmlPel As Integer, dtTanggalPelayanan As Date, strkodedokter As String, strStatusCITO As String, f_TarifCito As Currency) As Boolean
    On Error GoTo errLoad
    sp_BiayaPelayanan = True
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuanganx)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, strKdPelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, strKdKelas)
        .Parameters.Append .CreateParameter("StatusCITO", adChar, adParamInput, 1, strStatusCITO)
        .Parameters.Append .CreateParameter("Tarif", adInteger, adParamInput, , curTarif)
        .Parameters.Append .CreateParameter("JmlPelayanan", adInteger, adParamInput, , intJmlPel)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss"))

        Call msubRecFO(rs, "SELECT KdPelayananRS FROM dbo.PelayananRuangan WHERE (Status IN ('CU', 'MA', 'RG')) AND (KdPelayananRS = '" & strKdPelayananRS & "')")
        If rs.EOF = False Then
            Call msubRecFO(rs, "SELECT NoPakai FROM dbo.V_DaftarPasienRIAktif WHERE (NoPendaftaran = '" & mstrNoPen & "')")
            If rs.EOF = False Then
                .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, rs(0))
            Else
                .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, Null)
            End If
        Else
            .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, Null)
        End If

        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strkodedokter)
        .Parameters.Append .CreateParameter("StatusAPBD", adChar, adParamInput, 2, strStatusAPBD)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, strKdJenisTarif)
        .Parameters.Append .CreateParameter("TarifCito", adInteger, adParamInput, , f_TarifCito)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("IdPegawai2", adChar, adParamInput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "Add_BiayaPelayanan"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Biaya Pelayanan Pasien", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            sp_BiayaPelayanan = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With

    Exit Function
errLoad:
    sp_BiayaPelayanan = False
    Call msubPesanError
End Function

'simpan data perawat
Private Function sp_PetugasPemeriksaBP(F_dtTanggalPelayanan As Date, F_strKodePelayanan As String, F_StrIdPerawat As String, F_IdUser As String) As Boolean
    On Error GoTo errLoad

    sp_PetugasPemeriksaBP = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuanganx)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(F_dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, F_strKodePelayanan)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, F_StrIdPerawat)  'kode perawat
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, F_IdUser)

        .ActiveConnection = dbConn
        .CommandText = "Add_PetugasPemeriksaBP"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data petugas pemeriksa BP", vbExclamation, "Validasi"
            sp_PetugasPemeriksaBP = False
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    sp_PetugasPemeriksaBP = False
    Call msubPesanError
End Function

'untuk set grid pelayanan
Private Sub subSetGidPelayanan()
    With fgPelayanan
        .Clear
        .Rows = 2
        .Cols = 11
        .TextMatrix(0, 0) = "Kode Pelayanan"
        .TextMatrix(0, 1) = "Nama Pelayanan"
        .TextMatrix(0, 2) = "Jumlah"
        .TextMatrix(0, 3) = "Biaya Satuan"
        .TextMatrix(0, 4) = "Biaya Total"
        .TextMatrix(0, 5) = "Tgl Berlaku"
        .TextMatrix(0, 6) = "Kode Dokter"
        .TextMatrix(0, 7) = "Status CITO"
        .TextMatrix(0, 8) = "Biaya CITO"
        .TextMatrix(0, 9) = "Tgl Pelayanan"
        .TextMatrix(0, 10) = "StatusDelegasi"
        .ColWidth(0) = 0
        .ColWidth(1) = 4500
        .ColWidth(2) = 700
        .ColWidth(3) = 1200
        .ColWidth(4) = 1400
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 1200
        .ColWidth(9) = 0
        .ColWidth(10) = 0
    End With
End Sub

'untuk set grid obat alkes
Private Sub subSetGridObatAlkes()
    With fgDOA
        .Clear
        .Rows = 2
        .Cols = 10
        .TextMatrix(0, 0) = "Kode Barang"
        .TextMatrix(0, 1) = "Nama Barang"
        .TextMatrix(0, 2) = "Kode Asal"
        .TextMatrix(0, 3) = "Jumlah"
        .TextMatrix(0, 4) = "Harga Satuan"
        .TextMatrix(0, 5) = "Satuan"
        .TextMatrix(0, 6) = "Kode Dokter"
        .TextMatrix(0, 7) = "tgl Pelayanan"
        .TextMatrix(0, 8) = "Asal Barang"
        .TextMatrix(0, 9) = "KdPelayananRS"
        .ColWidth(0) = 0
        .ColWidth(1) = 4500
        .ColWidth(2) = 0
        .ColWidth(3) = 700
        .ColWidth(4) = 1200
        .ColWidth(5) = 900
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 1000
        .ColWidth(9) = 0
    End With
End Sub

'untuk cek validasi
Private Function funcCekValidasi() As Boolean
    If fgPelayanan.TextMatrix(1, 0) = "" Then
        MsgBox "Pilihan Pelayanan Pasien Harus Diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        txtNamaPelayanan.SetFocus
        Exit Function
    End If
    funcCekValidasi = True
End Function

'untuk enable/disable button reg
Private Sub subEnableButtonReg(blnStatus As Boolean)
    fraPDokter.Enabled = blnStatus
    fraPPelayanan.Enabled = blnStatus
    fgPelayanan.Enabled = blnStatus
    fgDOA.Enabled = blnStatus
    cmdSimpan.Enabled = blnStatus
End Sub

'untuk mengecek stok barang
Private Function funcCekStokBarang(intBarang As Integer, strSatuanJml As String, intJml As Integer) As Boolean
    If strSatuanJml = "S" Then
        If (intJml * typBarang(intBarang).intJmlTerkecil) > _
            typBarang(intBarang).intJmlTempTotal Then
            MsgBox "Stok Barang '" & typBarang(intBarang).strNamaBarang & "' Tidak Cukup !", vbCritical, "Validasi"
            funcCekStokBarang = False
            Exit Function
        Else
            typBarang(intBarang).intJmlTempTotal = typBarang(intBarang).intJmlTempTotal - (intJml * typBarang(intBarang).intJmlTerkecil)
        End If
    Else
        If intJml > typBarang(intBarang).intJmlTempTotal Then
            MsgBox "Stok Barang '" & typBarang(intBarang).strNamaBarang & "' Tidak Cukup !", vbCritical, "Validasi"
            funcCekStokBarang = False
            Exit Function
        Else
            typBarang(intBarang).intJmlTempTotal = typBarang(intBarang).intJmlTempTotal - intJml
        End If
    End If
    funcCekStokBarang = True
End Function

Private Sub txtNamaPerawat_Change()
    On Error GoTo errLoad

    Call subLoadListPemeriksa("where [Nama Pemeriksa] LIKE '%" & txtNamaPerawat.Text & "%'")
    lvPemeriksa.Visible = True

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtNamaPerawat_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If lvPemeriksa.Visible = True Then If lvPemeriksa.ListItems.Count > 0 Then lvPemeriksa.SetFocus
        Case vbKeyEscape
            lvPemeriksa.Visible = False
    End Select
End Sub

Private Sub txtNamaPerawat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If lvPemeriksa.Visible = True Then
            lvPemeriksa.SetFocus
        Else
            optCito(1).SetFocus
        End If
    End If
End Sub

Private Sub subSetGridPerawatPerPelayanan()
    With fgPerawatPerPelayanan
        .Cols = 6
        .Rows = 1

        .MergeCells = flexMergeFree

        .TextMatrix(0, 0) = "NoPendaftaran"
        .TextMatrix(0, 1) = "Kode Ruangan"
        .TextMatrix(0, 2) = "Tgl Pelayanan"
        .TextMatrix(0, 3) = "Kode Pelayanan"
        .TextMatrix(0, 4) = "IdPegawai"
        .TextMatrix(0, 5) = "IdUser"

    End With
End Sub

Private Sub subLoadListPemeriksa(Optional strKriteria As String)
    Dim strKey As String

    strSQL = "select * from v_daftarpemeriksapasien " & strKriteria & " order by [Nama Pemeriksa]"
    Call msubRecFO(rs, strSQL)

    With lvPemeriksa
        .ListItems.Clear
        For i = 0 To rs.RecordCount - 1
            strKey = "key" & rs(0).Value
            .ListItems.Add , strKey, rs(1).Value
            rs.MoveNext
        Next

        .Top = fraPDokter.Top + txtNamaPerawat.Top + txtNamaPerawat.Height
        .Left = fraPDokter.Left + txtNamaPerawat.Left
        .Height = 1815
        .ColumnHeaders.Item(1).Width = lvPemeriksa.Width - 500

        If subJmlTotal = 0 Then Exit Sub
        For i = 1 To .ListItems.Count
            For j = 1 To subJmlTotal
                If .ListItems(i).Key = subKdPemeriksa(j) Then .ListItems(i).Checked = True
            Next j
        Next i
    End With
End Sub

Private Function sp_Take_TarifBPT() As Currency
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, strKodePelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, mstrKdKelasx)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, strKdJenisTarif)
        .Parameters.Append .CreateParameter("TarifCito", adCurrency, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("TarifTotal", adCurrency, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("StatusCito", adChar, adParamInput, 1, IIf(optCito(0).Value = True, "Y", "T"))
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, IIf(chkDilayaniDokter.Value = vbChecked, mstrKdDokter, Null))
        .Parameters.Append .CreateParameter("IdDokter2", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("IdDokter3", adChar, adParamInput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "Take_TarifBPT"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam Pengambilan biaya tarif", vbExclamation, "Validasi"
            sp_Take_TarifBPT = 0
            subcurTarifBiayaSatuan = 0
        Else
            sp_Take_TarifBPT = .Parameters("TarifCito").Value
            subcurTarifBiayaSatuan = .Parameters("TarifTotal").Value
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_Take_TarifOA(f_KdAsal As String, f_HargaSatuan As Currency) As Currency
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 6, f_KdAsal)
        .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , CCur(f_HargaSatuan))
        .Parameters.Append .CreateParameter("TarifTotal", adCurrency, adParamOutput, , Null)

        .ActiveConnection = dbConn
        .CommandText = "Take_TarifOA"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam Pengambilan biaya tarif", vbExclamation, "Validasi"
            sp_Take_TarifOA = 0
        Else
            sp_Take_TarifOA = .Parameters("TarifTotal").Value
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub txtTarif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtDiscount.Enabled = True Then
            txtDiscount.SetFocus
        Else
            If txtCharge.Enabled = True Then txtCharge.SetFocus Else cmdAddKomponen.SetFocus
        End If
    End If

    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtTarif_LostFocus()
    txtTarif = IIf(val(txtTarif) = 0, 0, Format(txtTarif, "#,###"))
End Sub

