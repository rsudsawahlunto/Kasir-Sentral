VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucLoopingQuery 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6825
   ScaleHeight     =   2520
   ScaleWidth      =   6825
   Begin VB.Timer tmrQueryLooping 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   4680
      Top             =   1680
   End
   Begin VB.Frame frLoopingQuery 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   495
         Left            =   5040
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin MSComctlLib.ProgressBar pgbProsesLooping 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblPesan 
         AutoSize        =   -1  'True
         Caption         =   "Silahkan tunggu..."
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label lblPersen 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   390
      End
   End
End
Attribute VB_Name = "ucLoopingQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************'
'* penting!! yang di comment jangan di hapus *'
'* ttd: Rahmat Hidayat                       *'
'*********************************************'
Option Explicit

Private strTglAwal As String
Private strTglAkhir As String

'Private lngJumHari As Long
'Private lngCountHari As Long
'Private intHalfDay As Integer
'Private lngPgbVal As Long
'
'Private strTglAwalFilter As String, strTglAkhirFilter As String
'Private strTglAwalFilter12 As String, strTglAkhirFilter12 As String
'Private strFullQuery As String
Private strQuery As String, strQueryFilter As String

Private blnBatal As Boolean ', blnExecuteQuery As Boolean

Property Get TanggalAwal() As String
    TanggalAwal = strTglAwal
End Property

Property Let TanggalAwal(ByVal val As String)
    strTglAwal = val
End Property

Property Get TanggalAkhir() As String
    TanggalAkhir = strTglAkhir
End Property

Property Let TanggalAkhir(ByVal val As String)
    strTglAkhir = val
End Property

Property Get Pesan() As String
    Pesan = lblPesan.Caption
End Property

Property Let Pesan(ByVal val As String)
    lblPesan.Caption = val
End Property

'** double loop **'
'Public Sub subStartLooping(ByVal FullQueryString As String, ByVal NamaFieldTgl As String)
'    On Error GoTo errhandler
'    Dim lngJumHari As Long
'    Dim i As Long, j As Integer
'    Dim strTglAwalFilter As String, strTglAkhirFilter As String
'    Dim strTglAwalFilter12 As String, strTglAkhirFilter12 As String
'    Dim strQuery As String, strQueryFilter As String
'    Dim lngPgbVal As Long
'
'    UserControl.Extender.Visible = True
'    If InStr(1, FullQueryString, "[ft]", vbTextCompare) = 0 Then
'        MsgBox "Tidak ada kombinasi karakter '[ft]' pada query string yang di berikan!" & vbCrLf & _
'            "Proses looping dibatalkan.", vbCritical, "Error Query"
'        GoTo jump1
'    End If
'    lngJumHari = DateDiff("d", CDate(strTglAwal), CDate(strTglAkhir))
'    If lngJumHari = 0 Then lngJumHari = 1
'    With pgbProsesLooping
'        .Max = lngJumHari * 2
'        .Min = 0
'        .Value = 0
'    End With
'    lblPersen.Caption = "0%"
'    For i = 1 To lngJumHari
'        DoEvents
'        If i = 1 Then
'            strTglAwalFilter = Format(strTglAwal, "yyyy/MM/dd HH:mm:00")
'            strTglAkhirFilter = Format(strTglAwal, "yyyy/MM/dd 23:59:59")
'        Else
'            strTglAwalFilter = Format(DateAdd("d", 1, strTglAkhirFilter), "yyyy/MM/dd 00:00:00")
'            strTglAkhirFilter = Format(strTglAwalFilter, "yyyy/MM/dd 23:59:59")
'        End If
'        If i = lngJumHari Then
'            strTglAkhirFilter = Format(strTglAkhir, "yyyy/MM/dd HH:mm:ss")
'        End If
'        For j = 1 To 2
'            DoEvents
'            If blnBatal Then Exit Sub: blnBatal = False
'            If j = 1 Then
'                strTglAwalFilter12 = strTglAwalFilter
'                strTglAkhirFilter12 = Format(strTglAwalFilter, "yyyy/MM/dd 12:00:00")
'            Else
'                strTglAwalFilter12 = strTglAkhirFilter12
'                strTglAkhirFilter12 = strTglAkhirFilter
'            End If
'            strQueryFilter = NamaFieldTgl & " between '" & strTglAwalFilter12 & "' and '" & strTglAkhirFilter12 & "'"
'            strQuery = Replace(FullQueryString, "[ft]", strQueryFilter)
'            DoEvents
'            dbConn.CommandTimeout = 100
'            dbConn.Execute strQuery
'            DoEvents
'            lngPgbVal = lngPgbVal + 1
'            With pgbProsesLooping
'                .Value = lngPgbVal
'                lblPersen.Caption = CStr(CInt((.Value / .Max) * 100)) & "%"
'            End With
'        Next
'    Next
'jump1:
'    UserControl.Extender.Visible = False
'    Exit Sub
'errhandler:
'    MsgBox Err.Description, vbCritical, "Error Query Looping"
'    UserControl.Extender.Visible = False
'End Sub

'** single loop **'
Public Sub subStartLooping(ByVal FullQueryString As String, ByVal NamaFieldTgl As String)
    On Error GoTo errhandler
    Dim lngJumHari As Long
    Dim i As Long, j As Integer
    Dim strTglAwalFilter As String, strTglAkhirFilter As String
    Dim strTglAwalFilter12 As String, strTglAkhirFilter12 As String
    Dim strQuery As String, strQueryFilter As String

    UserControl.Extender.Visible = True
    If InStr(1, FullQueryString, "[ft]", vbTextCompare) = 0 Then
        MsgBox "Tidak ada kombinasi karakter '[ft]' pada query string yang di berikan!" & vbCrLf & _
            "Proses looping dibatalkan.", vbCritical, "Error Query"
        GoTo jump1
    End If
    lngJumHari = (DateDiff("d", CDate(strTglAwal), CDate(strTglAkhir)) + 1) * 2
    If lngJumHari = 0 Then lngJumHari = 1
    With pgbProsesLooping
        .Max = lngJumHari
        .Min = 0
        .Value = 0
    End With
    lblPersen.Caption = "0%"
    For i = 1 To lngJumHari
        DoEvents
        j = i Mod 2
        If i = 1 Then
            strTglAwalFilter = Format(strTglAwal, "yyyy/MM/dd HH:mm:ss")
            strTglAkhirFilter = Format(strTglAwal, "yyyy/MM/dd 23:59:59")
        ElseIf j = 1 Then
            strTglAwalFilter = Format(DateAdd("d", 1, strTglAkhirFilter), "yyyy/MM/dd 00:00:00")
            strTglAkhirFilter = Format(strTglAwalFilter, "yyyy/MM/dd 23:59:59")
        End If
        If i = lngJumHari Then
            strTglAkhirFilter = Format(strTglAkhir, "yyyy/MM/dd HH:mm:59")
        End If
        If blnBatal Then Exit Sub: blnBatal = False
        If j = 1 Then
            strTglAwalFilter12 = strTglAwalFilter
            strTglAkhirFilter12 = Format(strTglAwalFilter, "yyyy/MM/dd 12:00:00")
        ElseIf j = 0 Then
            strTglAwalFilter12 = strTglAkhirFilter12
            strTglAkhirFilter12 = strTglAkhirFilter
        End If
'        MsgBox strTglAwalFilter12 & ", " & strTglAkhirFilter12
'    Next
'    Exit Sub
        strQueryFilter = NamaFieldTgl & " between '" & strTglAwalFilter12 & "' and '" & strTglAkhirFilter12 & "'"
        strQuery = Replace(FullQueryString, "[ft]", strQueryFilter)
        DoEvents
        dbConn.CommandTimeout = 120
        dbConn.Execute strQuery
        DoEvents
        With pgbProsesLooping
            .Value = i
            lblPersen.Caption = CStr(CInt((.Value / .Max) * 100)) & "%"
        End With
    Next
jump1:
    UserControl.Extender.Visible = False
    Exit Sub
errhandler:
    MsgBox Err.Description, vbCritical, "Error Query Looping"
    UserControl.Extender.Visible = False
End Sub

'Public Sub subStartLooping(ByVal FullQueryString As String, ByVal NamaFieldTgl As String)
'    On Error GoTo errHandler
'    UserControl.Extender.Visible = True
'    If InStr(1, FullQueryString, "[ft]", vbTextCompare) = 0 Then
'        MsgBox "Tidak ada kombinasi karakter '[ft]' pada query string yang di berikan!" & vbCrLf & _
'            "Proses looping dibatalkan.", vbCritical, "Error Query"
'        GoTo InvisibleUC
'    End If
'    lngJumHari = DateDiff("d", CDate(strTglAwal), CDate(strTglAkhir))
'    If lngJumHari = 0 Then lngJumHari = 1
'    lngCountHari = 1
'    intHalfDay = 1
'    lngPgbVal = 0
'    With pgbProsesLooping
'        .Max = lngJumHari * 2
'        .Min = 0
'        .Value = 0
'    End With
'    lblPersen.Caption = "0%"
'    tmrQueryLooping.Enabled = True
'    Exit Sub
'InvisibleUC:
'    UserControl.Extender.Visible = False
'    Exit Sub
'errHandler:
'    MsgBox Err.Description, vbCritical, "Error Query Looping"
'End Sub

Private Sub cmdBatal_Click()
    blnBatal = True
End Sub

'Private Sub tmrQueryLooping_Timer()
'    If lngCountHari > lngJumHari Then
'        tmrQueryLooping.Enabled = False
'        Exit Sub
'    End If
'    If lngCountHari = 1 Then
'        strTglAwalFilter = Format(strTglAwal, "yyyy/MM/dd HH:mm:00")
'        strTglAkhirFilter = Format(strTglAwal, "yyyy/MM/dd 23:59:59")
'        lngCountHari = lngCountHari + 1
'    ElseIf lngCountHari > 1 And intHalfDay > 2 Then
'        strTglAwalFilter = Format(DateAdd("d", 1, strTglAkhirFilter), "yyyy/MM/dd 00:00:00")
'        strTglAkhirFilter = Format(strTglAwal, "yyyy/MM/dd 23:59:59")
'        intHalfDay = 1
'        lngCountHari = lngCountHari + 1
'    ElseIf lngCountHari = lngJumHari Then
'        strTglAkhirFilter = Format(strTglAkhir, "yyyy/MM/dd HH:mm:00")
'    End If
'    If intHalfDay = 1 Then
'        strTglAwalFilter12 = strTglAwalFilter
'        strTglAkhirFilter12 = Format(strTglAwalFilter, "yyyy/MM/dd 12:00:00")
'    ElseIf intHalfDay = 2 Then
'        strTglAwalFilter12 = strTglAkhirFilter12
'        strTglAkhirFilter12 = strTglAkhirFilter
'    End If
'    strQueryFilter = NamaFieldTgl & " between '" & strTglAwalFilter12 & "' and '" & strTglAkhirFilter12 & "'"
'    strQuery = Replace(FullQueryString, "[ft]", strQueryFilter)
'    dbConn.CommandTimeout = 100
'    dbConn.Execute strQuery
'    lngPgbVal = lngPgbVal + 1
'    With pgbProsesLooping
'        .Value = lngPgbVal
'        lblPersen.Caption = CStr(CInt((.Value / .Max) * 100)) & "%"
'    End With
'    intHalfDay = intHalfDay + 1
'End Sub

Private Sub UserControl_Initialize()
    cmdBatal.Visible = False
End Sub

Private Sub UserControl_InitProperties()
    UserControl.Extender.Visible = False
End Sub
