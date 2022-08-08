VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmViewerKuitansiBend26versi2 
   Caption         =   "Kuitansi Model : Bend 26"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmViewerKuitansiBend26versi2.frx":0000
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
Attribute VB_Name = "frmViewerKuitansiBend26versi2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crKuitansiBend26

Private Sub subDataReport(s_NamaReport As CRAXDRT.Report, Optional s_Versi As Integer)
    On Error GoTo errLoad

    With Report
        .txtNamaRS.SetText strNNamaRS
        .txtNoCM.SetText varNoCMCek
        .txtTotalBiaya.SetText Format(varJmlBayarCek, "#,###.-")
        .txtTotalBiayaTerbilang.SetText "(" & NumToText(CDbl(varJmlBayarCek)) & ")"
        .Total1.SetText Format(varJmlBayarCek, "#,###.-")
        .txtNamaPasien.SetText varAnPasienCek
        .txtAlamat.SetText varAlamatCek
        .txtKeterangan.SetText ""
        .txtUntuk.SetText "Biaya perawatan an. " & UCase(varNamaPasienCek) & "/ No. RM : " & varNoCMCek & " di ruang " & varRuangCek & " dari tanggal " & CStr(Format(varTglMasukCek, "dd mmmm yyyy hh:MM:ss")) & " s/d " & CStr(Format(varTglKeluarCek, "dd mmmm yyyy hh:MM:ss")) & ", Syarat Penjamin : " & varKelompokPasienCek & " /" & varPenjaminCek
        .txtBKM.SetText varNoBKMCek
        .txtTglTerima.SetText varTglMasukCek
        .txtTglBayar.SetText varTglKeluarCek
        .txtNoBKM.SetText varNoBKMCek
        .txtPenyetor.SetText varAnPasienCek
    End With

    With CRViewer1
        .EnableGroupTree = False
        .ReportSource = s_NamaReport
        .ViewReport
        .Zoom (110)
    End With
    Exit Sub
errLoad:
    Call msubPesanError("subDataReport")
    MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Call subDataReport(Report, 1)
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
    Set frmViewerKuitansiBend26versi2 = Nothing
End Sub
