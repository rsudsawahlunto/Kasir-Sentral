VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmViewerSJP 
   Caption         =   "Medifirst2000 - Cetak No SJP"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmViewerSJP.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   4125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3765
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmViewerSJP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crCetakSJP

Private Sub Form_Load()
    Call openConnection

    Set frmViewerSJP = Nothing
    Set dbcmd = New ADODB.Command

    strSQL = "select * " & _
    " from V_CetakSuratJaminanPelayanan where " & _
    " NoSJP ='" & mstrNoSJP & "'"
    Call msubRecFO(rs, strSQL)

    With dbcmd
        .ActiveConnection = dbConn
        .CommandText = strSQL
        .CommandType = adCmdText
    End With

    With Report
        .Database.AddADOCommand dbConn, dbcmd
        .txtTanggalSJP.SetText IIf(IsNull(rs("TglSJP")), "", rs("TglSJP"))
        .txtNomorRujukan.SetText IIf(IsNull(rs("NoRujukan")), "", rs("NoRujukan"))
        .txtTanggalRujukan.SetText IIf(IsNull(rs("TglDirujuk")), "", rs("TglDirujuk"))
        .txtNomorKartuAskes.SetText IIf(IsNull(rs("NoKartuPeserta")), "", rs("NoKartuPeserta"))
        .txtAsalRujukan.SetText IIf(IsNull(rs("AsalRujukan")), "", rs("AsalRujukan"))
        .txtDiagnosa.SetText IIf(IsNull(rs("DiagnosaRujukan")), "", rs("DiagnosaRujukan"))

        .txtKelasPerawatan.SetText IIf(IsNull(rs("KelasPerawatan")), "", rs("KelasPerawatan"))
        .txtRuangPerawatan.SetText IIf(IsNull(rs("RuanganPerawatan")), "", rs("RuanganPerawatan"))
        .txtTanggalMasuk.SetText ""
        .txtTanggalKeluar.SetText ""
        .txtJumlahHariRawat.SetText ""

        .txtJaminanPelayananLuarPaket1.SetText ""
        .txtJaminanPelayananLuarPaket2.SetText ""
        .txtJaminanPelayananLuarPaket3.SetText ""
        .txtJaminanPelayananLuarPaket4.SetText ""

        .txtCatatanKhusus.SetText ""

        .txtNamaPasien.SetText IIf(IsNull(rs("NamaPasien")), "", rs("NamaPasien"))
        .txtJenisKelamin.SetText IIf(IsNull(rs("JK")), "", rs("JK"))
        .txtTanggalLahir.SetText IIf(IsNull(rs("TglLahir")), "", rs("TglLahir"))
        .txtStatus.SetText ""
        .txtBadanUsaha.SetText ""
        .txtNoMR.SetText IIf(IsNull(rs("NoCM")), "", rs("NoCM"))
        .txtRuangICU.SetText ""
        .txtTanggalMasukSJP.SetText ""
        .txtTanggalKeluarSJP.SetText ""
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .EnableGroupTree = False
        .Zoom 1
    End With

    Screen.MousePointer = vbDefault
    Set dbcmd = Nothing

End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmViewerSJP = Nothing
End Sub
