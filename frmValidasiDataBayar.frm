VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmValidasiDataBayar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 -Validasi Data"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmValidasiDataBayar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   13500
   Begin VB.TextBox txtNoCM 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1400
      TabIndex        =   9
      Top             =   1200
      Width           =   1035
   End
   Begin VB.TextBox txtNamaPasien 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   1200
      Width           =   3000
   End
   Begin VB.TextBox txtNoPendaftaran 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   1400
   End
   Begin MSComctlLib.ProgressBar pbData 
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   7200
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Max             =   200
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdPerbaiki 
      Caption         =   "&Perbaiki Data"
      Height          =   495
      Left            =   8400
      TabIndex        =   0
      Top             =   7200
      Width           =   1695
   End
   Begin VB.TextBox txtIsi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   4320
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdValidasiData 
      Caption         =   "&Validasi Data"
      Height          =   495
      Left            =   10080
      TabIndex        =   1
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   11760
      TabIndex        =   2
      Top             =   7200
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   5415
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   9551
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      Appearance      =   0
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
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
      Picture         =   "frmValidasiDataBayar.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   11640
      Picture         =   "frmValidasiDataBayar.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmValidasiDataBayar.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12135
   End
End
Attribute VB_Name = "frmValidasiDataBayar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim substrNomorRetur As String
Dim subbolSimpan As Boolean
Dim sqlQuery As String
Dim rsQuery As New ADODB.recordset
Dim iData As Integer

Private Sub subLoadData(Optional s_Kriteria As String)
    On Error Resume Next
    Dim strStatusData As String

    Set rs = Nothing
    Call subSetGrid
    Set rs = Nothing
    strSQL = " SELECT     NoPendaftaran, NoCM, NamaPasien, RuanganPelayanan, RuanganKasir, NoStruk, TglStruk, NoBKM" & _
    " FROM  V_PembayaranTagihanPasien4Validasi Where (NoPendaftaran like '" & txtNoPendaftaran.Text & "%' and NamaPasien like '" & txtNamaPasien.Text & "%' and NoCM like '" & txtNoCM.Text & "%') and NoBKM is null "
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then Exit Sub

    For i = 1 To rs.RecordCount
        pbData.Value = i
        pbData.Max = rs.RecordCount
        DoEvents
        With fgData
            .TextMatrix(i, 0) = rs("NoPendaftaran")
            .TextMatrix(i, 1) = rs("NoCM")
            .TextMatrix(i, 2) = rs("NamaPasien")
            .TextMatrix(i, 3) = rs("RuanganPelayanan")

            .TextMatrix(i, 4) = rs("RuanganKasir")
            .TextMatrix(i, 5) = rs("NoStruk")
            .TextMatrix(i, 6) = rs("TglStruk")
            If IsNull(rs("NoBKM")) Then
                .TextMatrix(i, 7) = "-"
            Else
                .TextMatrix(i, 7) = rs("NoBKM")

            End If
            .Rows = .Rows + 1
            rs.MoveNext
        End With
        pbData.Value = Int(pbData.Value) + 1
    Next i

    MsgBox "Load Data sukses!!", vbInformation, vbOK, "Informasi"
    i = 0
    pbData.Value = 0.0001

End Sub

Private Sub subLoadText()
    txtIsi.Left = fgData.Left
    Select Case fgData.Col
        Case 7
            txtIsi.MaxLength = 10
        Case Else
            Exit Sub
    End Select

    For i = 0 To fgData.Col - 1
        txtIsi.Left = txtIsi.Left + fgData.ColWidth(i)
    Next i
    txtIsi.Visible = True
    txtIsi.Top = fgData.Top - 7

    For i = 0 To fgData.Row - 1
        txtIsi.Top = txtIsi.Top + fgData.RowHeight(i)
    Next i

    If fgData.TopRow > 1 Then
        txtIsi.Top = txtIsi.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If

    txtIsi.Width = fgData.ColWidth(fgData.Col)
    txtIsi.Height = fgData.RowHeight(fgData.Row)

    txtIsi.Visible = True
    txtIsi.SelStart = Len(txtIsi.Text)
    txtIsi.SetFocus
End Sub

Private Sub subSetGrid()
    With fgData
        .Rows = 2
        .Cols = 8

        .RowHeight(0) = 400
        .TextMatrix(0, 0) = "No Pendaftaran"
        .TextMatrix(0, 1) = "NoCM"
        .TextMatrix(0, 2) = "Nama Pasien"
        .TextMatrix(0, 3) = "Ruang Pelayanan"
        .TextMatrix(0, 4) = "Ruang Kasir"
        .TextMatrix(0, 5) = "No Struk"
        .TextMatrix(0, 6) = "Tgl Struk"
        .TextMatrix(0, 7) = "NoBKM"

        .ColWidth(0) = 1400
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColWidth(1) = 1000
        .ColWidth(2) = 2800
        .ColWidth(3) = 2200
        .ColWidth(4) = 1800
        .ColWidth(5) = 1200
        .ColWidth(6) = 1800
        .ColWidth(7) = 1200
        .ColAlignment(7) = flexAlignCenterCenter
    End With
End Sub

Private Sub cmdPerbaiki_Click()
    On Error GoTo hell_

    With fgData
        If .TextMatrix(.Row, 7) <> "-" Then
            If MsgBox("Yakin akan memperbaiki data pasien dengan meng-update NoBKM ke " & vbCrLf _
                & " Tabel Pembayaran Tagihan pasien", vbInformation + vbYesNo, "validasi") = vbNo Then Exit Sub
                MsgBox "data berhasil di insert ke tabel pembayaran tagihan pasien!!", vbInformation, "Informasi"
                cmdValidasiData.SetFocus
            Else
                On Error GoTo duplicate_
                If MsgBox("Yakin akan memperbaiki data pasien ke  " & vbCrLf _
                    & " Tabel Pasien Sudah Bayar", vbInformation + vbYesNo, "validasi") = vbNo Then Exit Sub
                    strSQL = "insert into PasienSudahBayar values('" & .TextMatrix(.Row, 0) & "','" & .TextMatrix(.Row, 1) & "','" & .TextMatrix(.Row, 5) & "') "
                    dbConn.Execute strSQL
                    MsgBox "data berhasil di insert ke tabel pasien sudah bayar!!", vbInformation, "Informasi"
                    cmdValidasiData.SetFocus
                    Exit Sub
duplicate_:
                    MsgBox "Data sudah ada di tabel pasien sudah bayar", vbInformation, "informasi"
                End If
            End With
            Exit Sub
hell_:
            msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdValidasiData_Click()
    On Error GoTo hell_
    Call subLoadData
    Exit Sub
hell_:
    msubPesanError
End Sub

Private Sub fgData_DblClick()
    Call fgData_KeyDown(13, 0)
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)

    Select Case KeyCode
        Case 13
            If fgData.TextMatrix(fgData.Row, 2) = "" Then Exit Sub
            Call subLoadText
            txtIsi.Text = Trim(fgData.TextMatrix(fgData.Row, fgData.Col))
            txtIsi.SelStart = 0
            txtIsi.SelLength = Len(txtIsi.Text)
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subSetGrid
    pbData.Value = 0.0001
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If KeyAscii = 13 Then
        fgData.TextMatrix(fgData.Row, fgData.Col) = txtIsi.Text
        txtIsi.Visible = False
        If fgData.RowPos(fgData.Row) >= fgData.Height - 360 Then
            fgData.SetFocus
            SendKeys "{DOWN}"
            Exit Sub
        End If
        fgData.SetFocus
    ElseIf KeyAscii = 27 Then
        txtIsi.Visible = False
        fgData.SetFocus
    End If
End Sub

Private Function Add_TempHargaKomponenForPenunjang(f_KdRuangan As String, f_tglPelayanan As Date, f_KdPelayananRS As String, f_KdKelas As String, f_KdJenisTarif As String, f_TarifCito As Integer, f_JmlPelayanan As Integer, f_StatusCito As String, f_Kdlaboratory As String, f_KdRuanganAsal As String) As Boolean
    On Error GoTo errLoad

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_tglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, f_KdKelas)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, f_KdJenisTarif)
        .Parameters.Append .CreateParameter("TarifCito", adInteger, adParamInput, , f_TarifCito)
        .Parameters.Append .CreateParameter("JmlPelayanan", adInteger, adParamInput, , f_JmlPelayanan)
        .Parameters.Append .CreateParameter("StatusCito", adChar, adParamInput, 1, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("KdLaboratory", adChar, adParamInput, 3, f_Kdlaboratory)
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, f_KdRuanganAsal)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_TempHargaKomponenForPenunjangM"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data ", vbCritical, "Validasi"
        End If
    End With
    Set dbcmd = Nothing
    Call deleteADOCommandParameters(dbcmd)
    Exit Function
errLoad:
    Call msubPesanError
End Function

Private Function Add_TempHargaKomponen(f_KdRuangan As String, f_tglPelayanan As Date, f_KdPelayananRS As String, f_KdKelas As String, f_KdJenisTarif As String, f_TarifCito As Integer, f_JmlPelayanan As Integer, f_StatusCito As String, f_kdDokter As String, f_KdRuanganAsal As String) As Boolean
    On Error GoTo errLoad

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_tglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, f_KdKelas)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, f_KdJenisTarif)
        .Parameters.Append .CreateParameter("TarifCito", adInteger, adParamInput, , f_TarifCito)
        .Parameters.Append .CreateParameter("JmlPelayanan", adInteger, adParamInput, , f_JmlPelayanan)
        .Parameters.Append .CreateParameter("StatusCito", adChar, adParamInput, 1, f_StatusCito)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, f_kdDokter)
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, IIf(f_KdRuanganAsal = "", Null, f_KdRuanganAsal))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_TempHargaKomponenNew"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data ", vbCritical, "Validasi"
        End If
    End With
    Set dbcmd = Nothing
    Call deleteADOCommandParameters(dbcmd)

    Exit Function
errLoad:
    Call msubPesanError
End Function

Private Function Add_TempHargaKomponenForIBS_DBNew(f_KdRuangan As String, f_tglPelayanan As Date, f_KdPelayananRS As String, f_KdKelas As String, f_KdJenisTarif As String, f_JmlPelayanan As Integer, f_KdRuanganAsal As String) As Boolean
    On Error GoTo errLoad

    sp_Retur = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_tglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, f_KdKelas)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, f_KdJenisTarif)
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, f_KdRuanganAsal)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_TempHargaKomponenForIBS_DBNew"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data ", vbCritical, "Validasi"
            sp_Retur = False
        End If
    End With
    Set dbcmd = Nothing
    Call deleteADOCommandParameters(dbcmd)

    Exit Function
errLoad:
    sp_Retur = False
    Call msubPesanError
End Function

Private Function Add_TempHargaKomponenForIBSNew(f_KdRuangan As String, f_tglPelayanan As Date, f_KdPelayananRS As String, f_KdKelas As String, f_KdJenisTarif As String, f_JmlPelayanan As Integer, f_KdRuanganAsal As String) As Boolean
    On Error GoTo errLoad

    sp_Retur = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_tglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, f_KdKelas)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, f_KdJenisTarif)
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, f_KdRuanganAsal)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_TempHargaKomponenForIBSNew"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data ", vbCritical, "Validasi"
            sp_Retur = False
        End If
    End With
    Set dbcmd = Nothing
    Call deleteADOCommandParameters(dbcmd)

    Exit Function
errLoad:
    sp_Retur = False
    Call msubPesanError
End Function

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

