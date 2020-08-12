VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSimpanan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Simpanan Uang"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSimpanan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8205
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6480
      TabIndex        =   31
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   30
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdCari 
      Caption         =   "&Cari"
      Height          =   375
      Left            =   4800
      TabIndex        =   29
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox txtJenis 
      Alignment       =   2  'Center
      BackColor       =   &H00F4FEFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   27
      Top             =   1920
      Width           =   6855
   End
   Begin VB.TextBox txtNomor 
      Alignment       =   2  'Center
      BackColor       =   &H00F4FEFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   26
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   4920
      TabIndex        =   17
      Top             =   2400
      Width           =   3255
      Begin VB.TextBox txtdebet 
         BackColor       =   &H00F4FEFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtKredit 
         BackColor       =   &H00F4FEFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtSaldo 
         BackColor       =   &H00F4FEFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtJumlah 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debet :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   90
         TabIndex        =   25
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kredit :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "- Data Anggota -"
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   4695
      Begin VB.TextBox txtPekerjaan 
         BackColor       =   &H00F4FEFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox txtAlamat 
         BackColor       =   &H00F4FEFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtNama 
         BackColor       =   &H00F4FEFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtNoAnggota 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pekerjaan :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   510
         TabIndex        =   15
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   600
         TabIndex        =   14
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&No. Anggota :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   13
         Top             =   360
         Width           =   1020
      End
   End
   Begin KoperasiSys.Line Line2 
      Height          =   30
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "&Tambah"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   4440
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtTanggal 
      Height          =   330
      Left            =   1200
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Format          =   172883969
      CurrentDate     =   39813
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   120
      TabIndex        =   28
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Trans :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "&No Trans :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   810
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   120
      Picture         =   "frmSimpanan.frx":06EA
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transaksi Simpanan Uang Anggota  Koperasi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   6315
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmSimpanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xTotal As Currency
Dim xSaldo As Currency
Dim Kata As String

Private Sub cmdCari_Click()
Kata = InputBox("Masukkan No Transaksi Simpanan Yang akan dicari..", "seacrh...")
If Kata = "" Then Exit Sub
    oKoperasi.Open "Select * from QSimpan Where NoTrans='" & Kata & "'", cnKoperasi, adOpenDynamic, adLockPessimistic
    If Not oKoperasi.EOF Then
        With oKoperasi
            txtNomor = !NoTrans: dtTanggal.Value = !tglTrans
            txtNoAnggota = !NoAnggota: txtNama = !Nama
            txtAlamat = !Alamat: txtPekerjaan = !Pekerjaan
            txtdebet = Format(!Debet, "#,##0"): txtKredit = Format(!Kredit, "#,##0")
            txtSaldo = Format(!Saldo, "#,##0")
            .Close
        End With
        Set oKoperasi = Nothing
        cmdHapus.Enabled = True
        cmdEdit.Enabled = True
        cmdTambah.Caption = "&Batal"
        txtJumlah.Text = Format(Val(Int(txtdebet.Text)) - Val(Int(txtKredit.Text)), "#,##0")
        Exit Sub
    End If
    oKoperasi.Close
    Set oKoperasi = Nothing
    MsgBox " No Pinjaman [" & Kata & "] tidak terdaftar...", vbCritical
    
End Sub

Private Sub cmdEdit_Click()
Fokus txtdebet
blnEdit = True
cmdHapus.Enabled = False
cmdSimpan.Enabled = True
cmdCari.Enabled = False
cmdEdit.Enabled = False
End Sub

Private Sub cmdHapus_Click()

On Error GoTo Salah
    Kata = MsgBox("Anda Yakin Untuk Menghapus Data Simpanan [ " & txtNama & "] ini...", vbCritical + vbYesNo)
    If Kata = vbYes Then
        cnKoperasi.Execute "Delete * From tblsimpanan Where NoTrans='" & txtNomor.Text & "'"
        MsgBox " Data Simpanan Telah Di Hapus...", vbInformation
        Semula
        Exit Sub
    End If
    Exit Sub
Salah:
    MsgBox " Data Simpanan Tidak Dapat Di Hapus..", vbCritical
End Sub

Private Sub cmdSimpan_Click()
Me.MousePointer = 11
If Not blnEdit Then
    Simpan
Else
    Edit
End If
Me.MousePointer = 1
End Sub

Private Sub cmdTambah_Click()
If cmdTambah.Caption = "&Tambah" Then
    cmdTambah.Caption = "&Batal"
    GetNumber
    dtTanggal.Enabled = True
    dtTanggal.SetFocus
Else
    Semula
End If
End Sub

Private Sub cmdTutup_Click()
Unload Me
End Sub

Private Sub dtTanggal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Fokus txtJenis
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Ketengah Me
KeyForm Me

End Sub

Sub GetNumber()

    Dim Counter As String * 11
    Dim Hitung As Integer
    Dim Tgl As String
    
    
    Query "Select * from tblSimpanan order By [NoTrans]"
    Tgl = Format(Now, "dd/mm/yyyy")
    With oKoperasi
        If .RecordCount = 0 Then
            Counter = "SM-" + Right(Tgl, 2) + Mid(Tgl, 4, 2) + "0001"
        Else
           .MoveLast
            If Left(![NoTrans], 7) <> "SM-" + Right(Tgl, 2) + Mid(Tgl, 4, 2) Then
                Counter = "BK-" + Right(Tgl, 2) + Mid(Tgl, 4, 2) + "001"
            Else
                Hitung = Val(Right(!NoTrans, 4)) + 1
               Counter = "SM-" + Right(Tgl, 2) + Mid(Tgl, 4, 2) + Right("0000" & Hitung, 4)
            End If
        End If
        txtNomor.Text = Counter
        dtTanggal.SetFocus
    End With
End Sub

Sub Semula()
    ClearControl Me
    cmdTambah.Caption = "&Tambah"
    cmdSimpan.Enabled = False
    
End Sub

Private Sub txtdebet_Change()
On Error Resume Next
If txtdebet.Text <> 0 Then
    txtdebet.Text = Format(txtdebet, "#,#"): SendKeys "{end}"
End If
End Sub

Private Sub txtdebet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtdebet = "" Then
        txtdebet = 0
        Fokus txtKredit
    Else
        txtKredit = 0
        txtJumlah = Format(Val(Int(txtdebet.Text)) + Val(Int(txtSaldo.Text)), "#,##0")
        cmdSimpan.Enabled = True
        cmdSimpan.SetFocus
    End If
End If
End Sub


Private Sub txtJenis_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtJenis <> "" Then
        Fokus txtNoAnggota
        Exit Sub
    End If
End If
End Sub

Private Sub txtKredit_Change()
On Error Resume Next
If txtKredit <> 0 Then
    txtKredit.Text = Format(txtKredit, "#,#"): SendKeys "{end}"
End If
End Sub

Private Sub txtKredit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtKredit <> "" Then
        txtJumlah.Text = Format(Val(Int(txtSaldo.Text)) - Val(Int(txtKredit.Text)), "#,##0")
        cmdSimpan.Enabled = True
        cmdSimpan.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub txtNoAnggota_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtNoAnggota <> "" Then
        Query "Select * from tblAnggota Where NoAnggota='" & txtNoAnggota.Text & "'"
        If oKoperasi.EOF Then
            oKoperasi.Close
            Set oKoperasi = Nothing
                Fokus txtNoAnggota
                MsgBox " No Anggota = " & txtNoAnggota & " Tidak Terdaftar...", vbInformation
                Exit Sub
        End If
        Daftar
        oKoperasi.Close
        
        Query "Select * from QSimpan Where NoAnggota='" & txtNoAnggota & "'"
        If oKoperasi.EOF Then
            oKoperasi.Close
            Set oKoperasi = Nothing
            Fokus txtdebet
            txtSaldo.Text = 0
            Exit Sub
        End If
        txtSaldo = Format(oKoperasi!Total, "#,##0")
        oKoperasi.Close
        Set oKoperasi = Nothing
        Fokus txtdebet
    End If
End If
            
End Sub

Sub Daftar()
    With oKoperasi
        txtNama = !Nama
        txtAlamat = !Alamat
        txtPekerjaan = !Pekerjaan
    End With
End Sub

Function Simpan()
    On Error GoTo Salah
    If cnKoperasi Is Nothing Then
        If cnKoperasi.State > 0 Then cnKoperasi.Close
    End If
    
    cnKoperasi.Execute "Insert Into tblSimpanan Values ('" & txtNomor.Text & "'," & _
                                                        "'" & dtTanggal.Value & "'," & _
                                                        "'" & txtJenis.Text & "', " & _
                                                        "'" & txtJumlah & "')"
                                                        
                                                        
    cnKoperasi.Execute "Insert Into tblDetail values ('" & txtNomor.Text & "'," & _
                                                     "'" & txtNoAnggota.Text & "'," & _
                                                     "'" & txtdebet.Text & "'," & _
                                                     "'" & txtKredit.Text & "'," & _
                                                     "'" & txtJumlah.Text & "')"
                                                     
    cnKoperasi.Execute "Update tblDetail Set Total='" & txtJumlah & "' Where NoAnggota='" & txtNoAnggota.Text & "'"
    
    Semula
    Exit Function
Salah:
    MsgBox Err.Description
                               
End Function

Function Edit()
    On Error GoTo ErrSalah
        If cnKoperasi Is Nothing Then
            If cnKoperasi.State > 0 Then cnKoperasi.Close
        End If
            
            cnKoperasi.Execute "Update tblSimpana Set Saldo='" & txtJumlah.Text & "' Where NoTrans='" & txtNomor.Text & "'"
            
            cnKoperasi.Execute "Update tblDetail Set debet='" & txtdebet.Text & "',Kredit='" & _
            txtKredit.Text & "',Total='" & txtJumlah.Text & "' Where NoTrans='" & txtNomor.Text & "'"
            Semula
            Exit Function
ErrSalah:
            MsgBox Err.Description
    
End Function
