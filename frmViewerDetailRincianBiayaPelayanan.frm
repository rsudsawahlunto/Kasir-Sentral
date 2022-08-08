VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmViewerDetailRincianBiayaPelayanan 
   Caption         =   "Laporan Detail Rincian Biaya Pelayanan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmViewerDetailRincianBiayaPelayanan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
End
Attribute VB_Name = "frmViewerDetailRincianBiayaPelayanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report1 As New crDetailRincianBiayaPelayananPasienv1
Dim report2 As New crDetailRincianBiayaPelayananPasienv2
Dim Report3 As New crDetailRincianBiayaPelayananPasienv3
Dim Report4 As New crDetailRincianBiayaPelayananPasienv4
Dim Report5 As New crDetailRincianBiayaPelayananPasienv5

Private Sub subDataReport(s_NamaReport As CRAXDRT.Report, s_Versi As Integer)
    With s_NamaReport
        dbcmd.CommandText = strSQL
        dbcmd.CommandType = adCmdText
        .Database.AddADOCommand dbConn, dbcmd

        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        .txtPeriode.SetText "PERIODE " & Format(mdTglAwal, "dd MMMM yyyy") & " s/d " & Format(mdTglAkhir, "dd MMMM yyyy")

        .usJenisPasien.SetUnboundFieldSource ("{Ado.JenisPasien}")
        .usNoRegister.SetUnboundFieldSource ("{Ado.NoPendaftaran}")
        .usNamaPasien.SetUnboundFieldSource ("{Ado.NamaPasien}")
        .usInstalasiPelayanan.SetUnboundFieldSource ("{Ado.InstalasiPelayanan}")
        .usRuanganPelayanan.SetUnboundFieldSource ("{Ado.RuanganPelayanan}")
        .usNamaPaket.SetUnboundFieldSource ("{Ado.NamaPaket}")
        .usjenispelayanan.SetUnboundFieldSource ("{Ado.JenisItem}")
        .usNamaPelayanan.SetUnboundFieldSource ("{Ado.NamaItem}")
        .unqty.SetUnboundFieldSource ("{Ado.QtyItem}")
        .usKelasPelayanan.SetUnboundFieldSource ("{Ado.KelasPelayanan}")
        .usKelasDitanggung.SetUnboundFieldSource ("{Ado.KelasDitanggung}")
        .ucTotalBiaya.SetUnboundFieldSource ("{Ado.TotalBiaya}")
        .ucTotalTarifPenjamin.SetUnboundFieldSource ("{Ado.TotalTarifPenjamin}")
        .ucTotalDitanggungPenjamin.SetUnboundFieldSource ("{Ado.TotalDitanggungPenjamin}")
        .ucTotalDitanggungRS.SetUnboundFieldSource ("{Ado.TotalDitanggungRS}")
        .ucTotalDibebaskan.SetUnboundFieldSource ("{Ado.TotalDibebaskan}")
        .ucTotalHarusDibayar.SetUnboundFieldSource ("{Ado.TotalHarusDibayar}")
        If s_Versi = 4 Then
            .usDokterPemeriksa.SetUnboundFieldSource ("{Ado.Dokter}")
        End If
        If s_Versi = 5 Then
            .usNmPerusahaan.SetUnboundFieldSource ("{ado.PenjaminPasien}")
            .usNoKPK.SetUnboundFieldSource ("{ado.IdAsuransi}")
            .usNmPeserta.SetUnboundFieldSource ("{ado.NamaPeserta}")
            .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
            .usDokterPemeriksa.SetUnboundFieldSource ("{Ado.Dokter}")
        End If
    End With

    With CRViewer1
        .EnableGroupTree = True
        .ReportSource = s_NamaReport
        .ViewReport
        .Zoom (110)
    End With
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    If frmCetakDetailRincianBiayaPelayanan.optVersi1.Value = True Then
        Call subDataReport(Report1, 1)
    ElseIf frmCetakDetailRincianBiayaPelayanan.optVersi2.Value = True Then
        Call subDataReport(report2, 2)
    ElseIf frmCetakDetailRincianBiayaPelayanan.optVersi3.Value = True Then
        Call subDataReport(Report3, 3)
    ElseIf frmCetakDetailRincianBiayaPelayanan.optVersi4.Value = True Then
        Call subDataReport(Report4, 4)
    ElseIf frmCetakDetailRincianBiayaPelayanan.optVersi5.Value = True Then
        Call subDataReport(Report5, 5)
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmViewerDetailRincianBiayaPelayanan = Nothing
End Sub
