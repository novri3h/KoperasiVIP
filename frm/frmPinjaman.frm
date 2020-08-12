VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPinjaman 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Pinjaman Uang"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8880
   Begin VB.CommandButton cmdCari 
      Caption         =   "&Cari"
      Height          =   375
      Left            =   5400
      TabIndex        =   46
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   45
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   44
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox txtNomor 
      Alignment       =   1  'Right Justify
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
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Height          =   3615
      Left            =   4920
      TabIndex        =   20
      Top             =   1920
      Width           =   3855
      Begin VB.TextBox dtSelesai 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         TabIndex        =   47
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtAngsuran 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtJumlah 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtBunga 
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
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtAdmin 
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
         Left            =   960
         TabIndex        =   23
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtTotalAdmin 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtTotalBunga 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtMulai 
         Height          =   330
         Left            =   1710
         TabIndex        =   32
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   172556289
         CurrentDate     =   39813
      End
      Begin MSComCtl2.DTPicker dtBayar 
         Height          =   330
         Left            =   1680
         TabIndex        =   42
         Top             =   2880
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         OLEDropMode     =   1
         Format          =   172556291
         CurrentDate     =   39813
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Bayar"
         Height          =   195
         Index           =   17
         Left            =   120
         TabIndex        =   41
         Top             =   2880
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angsuran :"
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   40
         Top             =   2520
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pinjaman :"
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   38
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Bunga :"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   36
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Selesai :"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   34
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Mulai :"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bunga [%]:"
         Height          =   240
         Index           =   11
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Admin [%]:"
         Height          =   240
         Index           =   13
         Left            =   135
         TabIndex        =   25
         Top             =   600
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Keterangan Pinjaman"
      Height          =   1575
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   4695
      Begin VB.TextBox txtLama 
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
         Left            =   1440
         TabIndex        =   29
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtPinjaman 
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
         Left            =   1440
         TabIndex        =   27
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtKet 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
         Height          =   195
         Index           =   9
         Left            =   2160
         TabIndex        =   31
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lama Angsuran :"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   30
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pinjaman Pokok :"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   28
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan :"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "- Data Anggota -"
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   4695
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
         TabIndex        =   12
         Top             =   360
         Width           =   1695
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
         TabIndex        =   11
         Top             =   720
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
         TabIndex        =   10
         Top             =   1080
         Width           =   3375
      End
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
         TabIndex        =   9
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&No. Anggota :"
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   16
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama :"
         Height          =   240
         Index           =   3
         Left            =   600
         TabIndex        =   15
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat :"
         Height          =   240
         Index           =   4
         Left            =   510
         TabIndex        =   14
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pekerjaan :"
         Height          =   240
         Index           =   5
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   870
      End
   End
   Begin MSComCtl2.DTPicker dtTanggal 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   1440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      Format          =   172556289
      CurrentDate     =   39833
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "&Tambah"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   7920
      TabIndex        =   0
      Top             =   5640
      Width           =   855
   End
   Begin KoperasiSys.Line Line2 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   53
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tgl Pinjam :"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&No. Pinjam :"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pinjaman Uang Koperasi"
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
      TabIndex        =   4
      Top             =   240
      Width           =   3435
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmPinjaman.frx":0000
      Top             =   120
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmPinjaman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const xStatus = "Lunas"
Const xKet = "Pinjam"
Dim Kata As String

Function Semula()
    ClearControl Me
    KeyForm Me
    cmdSimpan.Enabled = False
    cmdTambah.Caption = "&Tambah"
    cmdTambah.SetFocus
    cmdCari.Enabled = True
    cmdEdit.Enabled = False
    cmdHapus.Enabled = False
End Function

Private Sub cmdCari_Click()
Kata = InputBox("Masukkan No Pinjaman Yang akan dicari..", "Seacrh...")
If Kata = "" Then Exit Sub
    oKoperasi.Open "Select * from QPinjam Where NoPinjam='" & Kata & "'", cnKoperasi, adOpenDynamic, adLockPessimistic
    If Not oKoperasi.EOF Then
        With oKoperasi
            txtNomor = !NoPinjam: dtTanggal.Value = !tglPinjam
            txtNoAnggota = !NoAnggota: txtNama = !Nama
            txtAlamat = !Alamat: txtPekerjaan = !Pekerjaan
            txtKet = !Keterangan: txtPinjaman = Format(![Pinjaman Pokok], "#,##0")
            txtLama = !Lama: txtBunga = !Bunga
            txtAdmin = !Admin: dtMulai.Value = !tglMulai
            dtSelesai = !TglSelesai: txtJumlah = Format(!JlhBunga, "#,##0")
            txtTotal = Format(![Total Pinjaman], "#,##0"): txtAngsuran = Format(!Angsuran, "#,##0")
            dtBayar.Value = !TglBayar
            .Close
        End With
        Set oKoperasi = Nothing
        cmdHapus.Enabled = True
        cmdEdit.Enabled = True
        cmdTambah.Caption = "&Batal"
        Exit Sub
    End If
    oKoperasi.Close
    Set oKoperasi = Nothing
    MsgBox " No Pinjaman [" & Kata & "] tidak terdaftar...", vbCritical
    
            
End Sub

Private Sub cmdEdit_Click()
blnEdit = True
Fokus txtKet
cmdSimpan.Enabled = True
cmdEdit.Enabled = False
cmdHapus.Enabled = False
cmdCari.Enabled = False
End Sub

Private Sub cmdHapus_Click()
On Error GoTo Salah
    Kata = MsgBox("Anda Yakin Untuk Menghapus Data Pinjaman = " & txtNoAnggota & " ini...", vbCritical + vbYesNo)
    If Kata = vbYes Then
        cnKoperasi.Execute "Delete From tblPinjaman Where NoPinjam='" & txtNomor.Text & "'"
        MsgBox " Data Pinjaman Telah Di Hapus...", vbInformation
        Semula
        Exit Sub
    End If
    Exit Sub
Salah:
    MsgBox " Data Pinjaman Tidak Dapat Di Hapus.." & Chr(10) & _
          "Silahkan Periksa Angsuran Angoota...", vbCritical
        
End Sub

Private Sub cmdSimpan_Click()

Me.MousePointer = 11
    If Not blnEdit Then
        Save
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
    dtTanggal.Value = Format(Now, "dd/mm/yyyy")
    dtTanggal.SetFocus
Else
    Semula
End If

End Sub

Private Sub Command3_Click()

End Sub

Private Sub dtBayar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdSimpan.Enabled = True
    cmdSimpan.SetFocus
End If
End Sub

Private Sub dtMulai_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    dtSelesai.Text = DateAdd("m", txtLama.Text, dtMulai.Value)
    dtBayar.Enabled = True
    dtBayar.Value = Format(Now, "dd/mm")
    
    dtBayar.SetFocus
End If
End Sub

Private Sub dtTanggal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Fokus txtNoAnggota
End Sub

Private Sub Form_Load()
Ketengah Me
End Sub

Private Sub txtAdmin_Change()
txtTotalAdmin.Text = Format(Val(Int(txtPinjaman.Text)) * Val(txtAdmin.Text) / 100, "#,##0"): SendKeys "{end}"
txtJumlah.Text = Format(Val(Int(txtTotalBunga.Text)) + Val(Int(txtTotalAdmin.Text)), "#,##0")
txtTotal.Text = Format(Val(Int(txtPinjaman.Text)) + Val(Int(txtJumlah.Text)), "#,##0")
txtAngsuran.Text = Format(Val(Int(txtTotal.Text)) / Val(Int(txtLama.Text)), "#,##0")
dtMulai.Enabled = True
dtMulai.Value = Format(Now, "dd/mm/yyyy")
dtMulai.SetFocus
End Sub

Private Sub txtAdmin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtAdmin <> "" Then
        dtMulai.SetFocus
    End If
End If
End Sub

Private Sub txtBunga_Change()
txtTotalBunga.Text = Format(Val(Int(txtPinjaman.Text)) * Val(txtBunga.Text) / 100, "#,##0"): SendKeys "{end}"
End Sub

Private Sub txtBunga_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtBunga <> "" Then
        Fokus txtAdmin
    End If
End If
End Sub

Private Sub txtKet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Fokus txtPinjaman
    
End Sub

Private Sub txtLama_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Fokus txtBunga
    
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
        Set oKoperasi = Nothing
        Query "Select * from tblPinjaman Where NoAnggota='" & txtNoAnggota.Text & "' And Status<>'" & xStatus & "'"
        If oKoperasi.EOF Then
            oKoperasi.Close
            Set oKoperasi = Nothing
            Fokus txtKet
            Exit Sub
        End If
            MsgBox "No Anggota = " & txtNoAnggota & " " & Chr(10) & _
                  "Nama       = " & txtNama.Text & " " & Chr(10) & _
                  "Belum Melunasi Uang Pinjaman...", vbCritical
            Fokus txtNoAnggota
            txtNoAnggota = "": txtNama = ""
            txtAlamat = "": txtPekerjaan = ""
        oKoperasi.Close
        Set oKoperasi = Nothing
    End If
End If

            
End Sub

Private Sub cmdTutup_Click()
Unload Me
End Sub

Sub GetNumber()
On Error GoTo Salah
    Dim Counter As String * 11
    Dim Hitung As Integer
    Dim Tgl As String
    Query "Select * from tblPinjaman order By [NoPinjam]"
    Tgl = Format(Now, "dd/mm/yyyy")
    With oKoperasi
        If .RecordCount = 0 Then
            Counter = "PJ-" + Right(Tgl, 2) + Mid(Tgl, 4, 2) + "0001"
        Else
           .MoveLast
            If Left(![NoPinjam], 7) <> "PJ-" + Right(Tgl, 2) + Mid(Tgl, 4, 2) Then
                Counter = "PJ-" + Right(Tgl, 2) + Mid(Tgl, 4, 2) + "001"
            Else
                Hitung = Val(Right(!NoPinjam, 4)) + 1
               Counter = "PJ-" + Right(Tgl, 2) + Mid(Tgl, 4, 2) + Right("0000" & Hitung, 4)
            End If
        End If
        txtNomor.Enabled = True
        txtNomor.Text = Counter
    End With
    Exit Sub
Salah:
    MsgBox Err.Description
End Sub

Sub Daftar()
    With oKoperasi
        txtNama = !Nama
        txtAlamat = !Alamat
        txtPekerjaan = !Pekerjaan
    End With
End Sub

Private Sub txtPinjaman_Change()
txtPinjaman.Text = Format(txtPinjaman, "#,#"): SendKeys "{end}"
End Sub

Private Sub txtPinjaman_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Fokus txtLama
End Sub

Sub Save()
    On Error GoTo Salah
    If cnKoperasi Is Nothing Then
        If cnKoperasi.State > 0 Then cnKoperasi.Close
    End If
    
    cnKoperasi.Execute "Insert Into tblPinjaman Values ('" & txtNomor.Text & "'," & _
                                                        "'" & dtTanggal.Value & "'," & _
                                                        "'" & txtNoAnggota.Text & "'," & _
                                                        "'" & txtKet.Text & "'," & _
                                                        "'" & txtPinjaman.Text & "'," & _
                                                        "'" & txtLama.Text & "'," & _
                                                        "'" & txtBunga.Text & "'," & _
                                                        "'" & txtAdmin.Text & "'," & _
                                                        "'" & dtMulai.Value & "'," & _
                                                        "'" & dtSelesai.Text & "'," & _
                                                        "'" & txtJumlah.Text & "'," & _
                                                        "'" & txtTotal.Text & "'," & _
                                                        "'" & txtAngsuran.Text & "'," & _
                                                        "'" & dtBayar.Value & "'," & _
                                                        "'" & xKet & "')"
                                                        
    Semula
    Exit Sub
Salah:
    MsgBox Err.Description
End Sub

Sub Edit()
    On Error GoTo Salah
    
    cnKoperasi.Execute "Update tblPinjaman Set Keterangan='" & txtKet.Text & "',[Pinjaman Pokok]='" & _
    txtPinjaman.Text & "',Lama='" & txtLama.Text & "',Bunga='" & txtBunga.Text & "',Admin='" & _
    txtAdmin.Text & "',TglMulai='" & dtMulai.Value & "',TglSelesai='" & dtSelesai.Text & "',JlhBunga='" & _
    txtJumlah.Text & "',[Total Pinjaman]='" & txtTotal.Text & "',Angsuran='" & txtAngsuran.Text & "',TglBayar='" & _
    dtBayar.Value & "',Status='" & xKet & "' Where NoPinjam='" & txtNomor.Text & "'"
    
    Semula
    Exit Sub
Salah:
    MsgBox Err.Description
End Sub
