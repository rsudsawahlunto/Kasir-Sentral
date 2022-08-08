VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmTarifPelMedikBaru 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Tarif Pelayanan Medik"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTarifPelMedikBaru.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   9975
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   2160
      TabIndex        =   31
      Top             =   8880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   370
      Left            =   3720
      TabIndex        =   16
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   370
      Left            =   6840
      TabIndex        =   14
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   370
      Left            =   5280
      TabIndex        =   15
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   370
      Left            =   8445
      TabIndex        =   17
      Top             =   8880
      Width           =   1455
   End
   Begin VB.Frame FrameInput 
      Caption         =   "Input Data"
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
      TabIndex        =   18
      Top             =   960
      Width           =   9975
      Begin VB.TextBox TxtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7500
         TabIndex        =   9
         Top             =   3315
         Width           =   2250
      End
      Begin VB.CommandButton cmdAddKomponen 
         Caption         =   "+"
         Height          =   375
         Left            =   8880
         TabIndex        =   6
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cmdMinKomponen 
         Caption         =   "-"
         Height          =   375
         Left            =   9375
         TabIndex        =   7
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtTarif 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7500
         MaxLength       =   12
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid fgKomponen 
         Height          =   1455
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   2566
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.TextBox TxtPelayananRS 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6360
         TabIndex        =   1
         Top             =   360
         Width           =   3405
      End
      Begin MSDataListLib.DataCombo DcJenisTarif 
         Height          =   330
         Left            =   7500
         TabIndex        =   3
         Top             =   840
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DcKelas 
         Height          =   330
         Left            =   1680
         TabIndex        =   2
         Top             =   840
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcPelayanan 
         Height          =   330
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcKomponenTarif 
         Height          =   330
         Left            =   1680
         TabIndex        =   4
         Top             =   1320
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label2 
         Caption         =   "Harga"
         Height          =   255
         Left            =   6360
         TabIndex        =   27
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Komponen Tarif"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pelayanan"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Kelas"
         Height          =   210
         Left            =   240
         TabIndex        =   22
         Top             =   885
         Width           =   405
      End
      Begin VB.Label Label5 
         Caption         =   "Total"
         Height          =   255
         Left            =   6360
         TabIndex        =   21
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Jenis Tarif"
         Height          =   255
         Left            =   6360
         TabIndex        =   20
         Top             =   878
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Data Tarif Pelayanan Yang Sudah Diinput"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      TabIndex        =   19
      Top             =   4800
      Width           =   9975
      Begin VB.TextBox txtCariJenisPelayanan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         TabIndex        =   11
         Top             =   3555
         Width           =   2655
      End
      Begin VB.TextBox txtCarinamapelayanan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         TabIndex        =   12
         Top             =   3555
         Width           =   4575
      End
      Begin MSDataGridLib.DataGrid dgTarifPelMedik 
         Height          =   2805
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4948
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   2
         RowHeight       =   15
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
      Begin MSDataListLib.DataCombo cboJenisTarif 
         Height          =   330
         Left            =   7620
         TabIndex        =   13
         Top             =   3555
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label8 
         Caption         =   "Jenis Tarif"
         Height          =   255
         Left            =   7680
         TabIndex        =   30
         Top             =   3285
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Pelayanan"
         Height          =   210
         Left            =   120
         TabIndex        =   29
         Top             =   3240
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pelayanan"
         Height          =   210
         Left            =   2880
         TabIndex        =   25
         Top             =   3285
         Width           =   1320
      End
      Begin VB.Label LblJumData 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Jumlah Data : 0"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   8520
         TabIndex        =   24
         Top             =   120
         Width           =   1275
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   28
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
      Picture         =   "frmTarifPelMedikBaru.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8160
      Picture         =   "frmTarifPelMedikBaru.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTarifPelMedikBaru.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmTarifPelMedikBaru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vEdit As Boolean
Dim subBolTampil As Boolean
Dim vBolErr As Boolean
Dim vTmpKdPelayananRS As String
Dim subStrKdKomponen() As String
Dim subIntJmlKomponen As Integer
Dim kdKelas As String, KdJnsTarif As String

Private Sub cboJenisTarif_Change()
    Call GridSource("Tarif")
    txtCariJenisPelayanan.SetFocus: txtCariJenisPelayanan.SelStart = Len(txtCariJenisPelayanan.Text)
End Sub

Private Sub cboJenisTarif_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then dgTarifPelMedik.SetFocus

End Sub

Private Sub cmdAddKomponen_Click()
On Error GoTo errLoad
    
    If dcKomponenTarif.MatchedWithList = False Then dcKomponenTarif.SetFocus: Exit Sub
    For i = 1 To fgKomponen.Rows - 1
        If fgKomponen.TextMatrix(i, 3) = dcKomponenTarif.BoundText Then
            fgKomponen.TextMatrix(i, 2) = IIf(val(txtTarif) = 0, 0, Format(txtTarif.Text, "#,###,###")) 'tarif
            Call HitungTotal
            Exit Sub
        End If
    Next i
    
    fgKomponen.Rows = fgKomponen.Rows + 1

    fgKomponen.TextMatrix(fgKomponen.Rows - 1, 1) = dcKomponenTarif.Text 'nama komponen
    fgKomponen.TextMatrix(fgKomponen.Rows - 1, 2) = IIf(val(txtTarif) = 0, 0, Format(txtTarif.Text, "#,###")) 'tarif
    fgKomponen.TextMatrix(fgKomponen.Rows - 1, 3) = dcKomponenTarif.BoundText 'kode komponen tarif
    
    Call HitungTotal

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdBatal_Click()
On Error GoTo errLoad
    
    Call kosong
    Call subLoadDcSource
    Call GridSource("Tarif")
    Call subSetGrid
    subBolTampil = False

Exit Sub
errLoad:
End Sub

Private Sub cmdCetak_Click()
On Error GoTo hell
'On Error Resume Next
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmtarif.Show
hell:
End Sub

Private Sub cmdHapus_Click()
On Error GoTo hell
    If MsgBox("Apakah yakin data akan dihapus.. ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, dcPelayanan.BoundText)
                .Parameters.Append .CreateParameter("kdkelas", adChar, adParamInput, 2, DcKelas.BoundText)
                .Parameters.Append .CreateParameter("Total", adCurrency, adParamInput, , CCur(TxtTotal))
                .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, DcJenisTarif.BoundText)
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "D")
               
                .ActiveConnection = dbConn
                .CommandText = "AUD_TarifPelayanan"
                .CommandType = adCmdStoredProc
                .Execute
                
                If .Parameters("return_value").Value <> 0 Then
                    MsgBox "Ada kesalahan saat penyimpanan data", vbExclamation, "Validasi"
                    Exit Sub
                End If
                 Call deleteADOCommandParameters(dbcmd)
                Set dbcmd = Nothing
            End With
    
    MsgBox "Data dihapus..", vbInformation, "Informasi"
    cmdBatal_Click
    Call RefreshGrid(txtCarinamapelayanan)
    Exit Sub
hell:
    If Err.Number = "-2147217873" Then
        MsgBox "Data tidak bisa dihapus", vbInformation
    Else
        MsgBox "Ada kesalahan dalam proses penghapusan data", vbInformation
    End If
    cmdHapus.Enabled = False
End Sub

Private Sub cmdMinKomponen_Click()
On Error GoTo errLoad
    
    If fgKomponen.Rows = 1 Then Exit Sub
    If MsgBox("Apakah anda yakin akan menghapus data ini?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub

    strQuery = " DELETE FROM HargaKomponen" & _
            " WHERE (KdKomponen = '" & dcKomponenTarif.BoundText & "') AND (KdKelas = '" & DcKelas.BoundText & "') AND (KdPelayananRS = '" & dcPelayanan.BoundText & "') AND (KdJenisTarif = '" & DcJenisTarif.BoundText & "')"
    dbConn.Execute strQuery
    
'    subStrKdKomponen(subIntJmlKomponen) = fgKomponen.TextMatrix(fgKomponen.Row, 3)
    subIntJmlKomponen = subIntJmlKomponen + 1
    
    If fgKomponen.Rows = 2 Then
        fgKomponen.TextMatrix(1, 1) = ""
        fgKomponen.TextMatrix(1, 2) = "0"
        fgKomponen.Rows = 1
    Else
        fgKomponen.RemoveItem fgKomponen.Row
    End If
    
    Call HitungTotal
        If Len(dcPelayanan) <> 0 And Len(DcKelas) <> 0 And Len(TxtPelayananRS) <> 0 Then
'        Call subTampil(dgTarifPelMedik.Columns("KdPelayananRS").Value, dgTarifPelMedik.Columns("KdKelas").Value, dgTarifPelMedik.Columns("KdJenisTarif").Value)
    'KdPelayananRS, KdKelas, KdJenisTarif,KdJnsPelayanan
        strSQL = "SELECT * FROM V_AmbilTarifPelayanan WHERE KdPelayananRS = '" & dcPelayanan.BoundText & "' and KdKelas = '" & DcKelas.BoundText & "'"
        msubRecFO rs, strSQL
        If rs.RecordCount <> 0 Then
             Call Simpan("TarifPelayanan")
             Call GridSource("Tarif")
'             subBolTampil = False

'            Call subLoadData
         Else
            Exit Sub
        End If
    End If

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo errLoad
    
    If Periksa("datacombo", dcPelayanan, "Nama pelayanan yang diinput," & vbNewLine & " Tidak sesuai dengan daftar") = False Then Exit Sub
    If Periksa("datacombo", DcKelas, "Kelas pelayanan kosong") = False Then Exit Sub
    If Periksa("datacombo", DcJenisTarif, "Jenis tarif kosong") = False Then Exit Sub
    
    If vBolErr = True Then Exit Sub
    
    Call Simpan("TarifPelayanan")
    Call Simpan("HargaKomponen")
    
    Call Add_HistoryLoginActivity("AUD_TarifPelayanan+AUD_HargaKomponen")
    cmdBatal_Click
    Call RefreshGrid(txtCarinamapelayanan)

Exit Sub
errLoad:
    Call msubPesanError
    Call RefreshGrid(txtCarinamapelayanan)
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub DcJenisTarif_Change()
On Error GoTo errLoad
    If dcPelayanan.MatchedWithList = False Then Exit Sub
    If DcKelas.MatchedWithList = False Then Exit Sub
    If DcJenisTarif.MatchedWithList = False Then Exit Sub
    If subBolTampil = True Then Exit Sub
    Call subTampil(dcPelayanan.BoundText, DcKelas.BoundText, DcJenisTarif.BoundText)
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub DcJenisTarif_Click(Area As Integer)
    Call DcJenisTarif_Change
End Sub

Private Sub dcJenisTarif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcKomponenTarif.SetFocus
End Sub

Private Sub dcKelas_Change()
    If subBolTampil = True Then Exit Sub
'    dcJenisTarif.BoundText = ""
End Sub

Private Sub dcKelas_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 13 Then
        If Len(Trim(DcKelas.Text)) = 0 Then Exit Sub
        strSQL = "Select KdKelas, DeskKelas " & _
            " FROM KelasPelayanan" & _
            " WHERE DeskKelas LIKE '%" & DcKelas.Text & "%'" & _
            " ORDER BY DeskKelas"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        DcKelas.BoundText = rs(0).Value
        DcJenisTarif.SetFocus
    End If
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKomponenTarif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTarif.SetFocus
End Sub

Private Sub dcPelayanan_Change()
On Error GoTo errLoad
    If dcPelayanan.MatchedWithList = False Then TxtPelayananRS.Text = "": Exit Sub
    strSQL = "Select ListPelayananRS.KdPelayananRS, ListPelayananRS.NamaPelayanan, JenisPelayanan.Deskripsi" & _
        " FROM ListPelayananRS INNER JOIN JenisPelayanan ON ListPelayananRS.KdJnsPelayanan = JenisPelayanan.KdJnsPelayanan" & _
        " WHERE ListPelayananRS.KdPelayananRS ='" & dcPelayanan.BoundText & "'" & _
        " ORDER BY NamaPelayanan"
    Call msubRecFO(rsB, strSQL)
    If rsB.EOF = True Then Exit Sub
    dcPelayanan.BoundText = rsB(0).Value: TxtPelayananRS.Text = rsB("Deskripsi").Value
    
    DcKelas.BoundText = ""
    
    DcJenisTarif.BoundText = ""
    fgKomponen.Clear
    Call subSetGrid
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcPelayanan_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 13 Then
        If Len(Trim(dcPelayanan.Text)) = 0 Then Exit Sub
        If TxtPelayananRS.Text <> "" Then GoTo stepNext
        strSQL = "Select ListPelayananRS.KdPelayananRS, ListPelayananRS.NamaPelayanan, JenisPelayanan.Deskripsi" & _
            " FROM ListPelayananRS INNER JOIN JenisPelayanan ON ListPelayananRS.KdJnsPelayanan = JenisPelayanan.KdJnsPelayanan" & _
            " WHERE ListPelayananRS.KdPelayananRS ='" & dcPelayanan.BoundText & "'" & _
            " ORDER BY NamaPelayanan"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then TxtPelayananRS.Text = "": Exit Sub
        dcPelayanan.BoundText = rs(0).Value: TxtPelayananRS.Text = rs("Deskripsi").Value
stepNext:
    DcKelas.BoundText = "": DcJenisTarif.BoundText = "": DcKelas.SetFocus
    End If
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgTarifPelMedik_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdHapus.SetFocus
On Error GoTo errLoad
    If dgTarifPelMedik.ApproxCount = 0 Then Exit Sub
    
    subBolTampil = True
    Call subTampil(dgTarifPelMedik.Columns("KdPelayananRS").Value, dgTarifPelMedik.Columns("KdKelas").Value, dgTarifPelMedik.Columns("KdJenisTarif").Value)
    With dgTarifPelMedik
        dcPelayanan.BoundText = .Columns("KdPelayananRS").Value
        TxtPelayananRS.Text = .Columns("Pelayanan").Value
        DcKelas.BoundText = .Columns("KdKelas").Value
        DcJenisTarif.BoundText = .Columns("KdJenisTarif").Value
        TxtTotal = IIf(val(.Columns("Total").Value) = 0, 0, Format(.Columns("Total").Value, "#,###,###"))
    End With
Exit Sub
errLoad:
End Sub

Private Sub dgTarifPelMedik_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

On Error GoTo errLoad
    If dgTarifPelMedik.ApproxCount = 0 Then Exit Sub
    
    subBolTampil = True
    Call subTampil(dgTarifPelMedik.Columns("KdPelayananRS").Value, dgTarifPelMedik.Columns("KdKelas").Value, dgTarifPelMedik.Columns("KdJenisTarif").Value)
    With dgTarifPelMedik
        dcPelayanan.BoundText = .Columns("KdPelayananRS").Value
        TxtPelayananRS.Text = .Columns("Pelayanan").Value
        DcKelas.BoundText = .Columns("KdKelas").Value
        DcJenisTarif.BoundText = .Columns("KdJenisTarif").Value
        TxtTotal = IIf(val(.Columns("Total").Value) = 0, 0, Format(.Columns("Total").Value, "#,###,###"))
    End With
Exit Sub
errLoad:
End Sub

Private Sub fgKomponen_RowColChange()
On Error Resume Next
    If fgKomponen.Row = 0 Then Exit Sub
    dcKomponenTarif.Text = fgKomponen.TextMatrix(fgKomponen.Row, 1)
    txtTarif.Text = fgKomponen.TextMatrix(fgKomponen.Row, 2)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    openConnection
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call cmdBatal_Click
    Call subSetGrid
    
End Sub

Private Sub txtCariJenisPelayanan_Change()
    Call GridSource("Tarif")
    txtCariJenisPelayanan.SetFocus: txtCariJenisPelayanan.SelStart = Len(txtCariJenisPelayanan.Text)
End Sub

Private Sub txtCariJenisPelayanan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then dgTarifPelMedik.SetFocus
Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtCarinamapelayanan_Change()
    Call GridSource("Tarif")
    txtCarinamapelayanan.SetFocus: txtCarinamapelayanan.SelStart = Len(txtCarinamapelayanan.Text)
End Sub

Private Sub subLoadDcSource()
On Error GoTo errLoad

    Call msubDcSource(dcPelayanan, rs, "Select KdPelayananRS, NamaPelayanan From ListPelayananRS where StatusEnabled=1 Order By NamaPelayanan")
    Call msubDcSource(DcKelas, rs, "Select KdKelas, DeskKelas From KelasPelayanan where StatusEnabled=1 ORDER BY DeskKelas")
    Call msubDcSource(DcJenisTarif, rs, "Select KdJenisTarif, JenisTarif From JenisTarif where StatusEnabled=1")
    Call msubDcSource(cboJenisTarif, rs, "Select KdJenisTarif, JenisTarif From JenisTarif where StatusEnabled=1")
    Call msubDcSource(dcKomponenTarif, rs, "SELECT KdKomponen, NamaKomponen FROM KomponenTarif where StatusEnabled=1")

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub GridSource(vCase As String)
On Error GoTo errLoad
    Select Case vCase
        Case "Tarif" 'u/ grid dgTarifPelMedik
            strSQL = "SELECT Deskripsi AS JenisPelayanan, Pelayanan, Kelas, JenisTarif, Total, KdPelayananRS, KdKelas, KdJenisTarif" & _
                " FROM V_AmbilTarifPelayanan" & _
                " WHERE (KdJenisTarif LIKE '%" & cboJenisTarif.BoundText & "%') AND (Pelayanan LIKE '%" & txtCarinamapelayanan & "%')" & _
                " AND (Deskripsi LIKE '%" & txtCariJenisPelayanan.Text & "%')"
            Call msubRecFO(rs, strSQL)
            
            LblJumData = "Jumlah Data: " & rs.RecordCount
            With dgTarifPelMedik
                Set .DataSource = rs
                .Columns("JenisPelayanan").Width = 2200
                .Columns("Pelayanan").Width = 2800
                .Columns("Kelas").Width = 1700
                .Columns("JenisTarif").Width = 1200
                .Columns("Total").Width = 1200
                .Columns("Total").Alignment = dbgRight
                .Columns("Total").NumberFormat = "#,###"
                .Columns("KdPelayananRS").Width = 0
                .Columns("KdKelas").Width = 0
                .Columns("KdJenisTarif").Width = 0
                .Visible = True
            End With
    End Select
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub kosong()
On Error Resume Next
    dcPelayanan.BoundText = ""
    TxtPelayananRS = ""
    DcKelas.BoundText = ""
    DcJenisTarif.BoundText = ""
    TxtTotal = "0"
    fgKomponen.Clear
    fgKomponen.Rows = 1
    dcKomponenTarif.BoundText = ""
    txtTarif.Text = ""
    dcPelayanan.SetFocus
End Sub

Private Sub subTampil(s_KdPelayanan As String, s_KdKelas As String, s_KdJenisTarif As String)
On Error GoTo errLoad
    
    If dgTarifPelMedik.ApproxCount = 0 Then Exit Sub
    strSQL = "SELECT dbo.KomponenTarif.NamaKomponen, dbo.HargaKomponen.Harga, dbo.HargaKomponen.KdKomponen " & _
            " FROM dbo.HargaKomponen INNER JOIN dbo.KomponenTarif ON dbo.HargaKomponen.KdKomponen = dbo.KomponenTarif.KdKomponen " & _
            " WHERE (dbo.HargaKomponen.KdPelayananRS = '" & s_KdPelayanan & "') AND " & _
            " (dbo.HargaKomponen.KdKelas = '" & s_KdKelas & "') AND (dbo.HargaKomponen.KdJenisTarif = '" & s_KdJenisTarif & "')"
    Call msubRecFO(rs, strSQL)
    Call subSetGrid
    
    If rs.RecordCount < 1 Then Exit Sub
    subBolTampil = True
    
    ReDim Preserve subStrKdKomponen(rs.RecordCount)
    subIntJmlKomponen = 1
    
    TxtTotal.Text = 0
    fgKomponen.Rows = rs.RecordCount + 1
    For i = 1 To rs.RecordCount
        With fgKomponen
            .TextMatrix(i, 1) = CStr(rs("NamaKomponen"))
            .TextMatrix(i, 2) = IIf(rs("Harga").Value = 0, 0, Format(rs("Harga").Value, "#,###"))
            TxtTotal.Text = val(TxtTotal.Text) + val(rs("Harga"))
            .TextMatrix(i, 3) = CStr(rs("KdKomponen"))
        End With
        rs.MoveNext
    Next
    
    dcKomponenTarif.Text = ""
    txtTarif.Text = ""
    subBolTampil = False

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub Simpan(vCase As String)
On Error GoTo hell
    Select Case vCase
        Case "TarifPelayanan"
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, dcPelayanan.BoundText)
                .Parameters.Append .CreateParameter("kdkelas", adChar, adParamInput, 2, DcKelas.BoundText)
                .Parameters.Append .CreateParameter("Total", adCurrency, adParamInput, , CCur(TxtTotal))
                .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, DcJenisTarif.BoundText)
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")
               
                .ActiveConnection = dbConn
                .CommandText = "AUD_TarifPelayanan"
                .CommandType = adCmdStoredProc
                .Execute
                
                If .Parameters("return_value").Value <> 0 Then
                    MsgBox "Ada kesalahan saat penyimpanan data", vbExclamation, "Validasi"
                    Exit Sub
                End If
            End With

        Case "HargaKomponen"
            For i = 1 To fgKomponen.Rows - 1
                Set dbcmd = New ADODB.Command
                With dbcmd
                    .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                    .Parameters.Append .CreateParameter("KdKomponen", adChar, adParamInput, 2, fgKomponen.TextMatrix(i, 3))
                    .Parameters.Append .CreateParameter("kdkelas", adChar, adParamInput, 2, DcKelas.BoundText)
                    .Parameters.Append .CreateParameter("KdPelayananRs", adChar, adParamInput, 6, dcPelayanan.BoundText)
                    .Parameters.Append .CreateParameter("Harga", adCurrency, adParamInput, , fgKomponen.TextMatrix(i, 2))
                    .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, DcJenisTarif.BoundText)
                    .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")
                    
                    .ActiveConnection = dbConn
                    .CommandText = "AUD_HargaKomponen"
                    .CommandType = adCmdStoredProc
                    .Execute
                
                    If .Parameters("return_value").Value <> 0 Then
                        MsgBox "Ada kesalahan saat penyimpanan data", vbExclamation, "Validasi"
                        Exit Sub
                    End If
                End With
            Next
    End Select
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Function sp_DeleteKomponen() As Boolean
On Error GoTo errLoad
Dim i As Integer
    sp_DeleteKomponen = True

    If subIntJmlKomponen = 0 Then Exit Function
    For i = 0 To subIntJmlKomponen
        dbConn.Execute _
            " DELETE FROM HargaKomponen" & _
            " WHERE (KdKomponen = '" & subStrKdKomponen(i) & "') AND (KdKelas = '" & DcKelas.BoundText & "') AND (KdPelayananRS = '" & dcPelayanan.BoundText & "') AND (KdJenisTarif = '" & DcJenisTarif.BoundText & "')"
    Next

Exit Function
errLoad:
    Call msubPesanError
    sp_DeleteKomponen = False
End Function

Sub HitungTotal()
On Error GoTo hell
    TxtTotal = 0
    If fgKomponen.Rows = 1 Then TxtTotal = 0: Exit Sub
    For i = 1 To fgKomponen.Rows - 1
        TxtTotal = val(TxtTotal) + IIf(Len(fgKomponen.TextMatrix(i, 2)) = 0, 0, CCur(fgKomponen.TextMatrix(i, 2)))
    Next
    
    TxtTotal = IIf(val(TxtTotal.Text) = 0, 0, Format(TxtTotal.Text, "#,###,###"))

Exit Sub
hell:
    MsgBox "Ada kesalahan data saat menjumlahlah harga " & vbCr _
        & "Harga total maximalnya 99.999.999,00", vbInformation
    Call HitungTotal
End Sub

Private Sub subSetGrid()
On Error Resume Next
    With fgKomponen
        .Clear
        .Cols = 4
        .Rows = 1
                
        .RowHeight(0) = 400
        .ColWidth(0) = 0
        .ColWidth(1) = 6750
        .ColWidth(2) = 2000
        .ColWidth(3) = 0
                
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignRightCenter
                
        .TextMatrix(0, 1) = "Komponen Tarif"
        .TextMatrix(0, 2) = "Harga"
        .TextMatrix(0, 3) = "Kode Komponen Tarif"
    End With
End Sub

Private Sub txtCarinamapelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dgTarifPelMedik.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtTarif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdAddKomponen.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

