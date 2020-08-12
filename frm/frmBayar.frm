VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBayar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Pembayaran "
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
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
   ScaleHeight     =   7095
   ScaleWidth      =   8670
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   59
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   58
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton cmdCari 
      Caption         =   "&Cari"
      Height          =   375
      Left            =   5040
      TabIndex        =   57
      Top             =   6600
      Width           =   855
   End
   Begin VB.ComboBox cmbPinjam 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1200
      TabIndex        =   49
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtNomor 
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
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   4920
      TabIndex        =   44
      Top             =   2160
      Width           =   3495
      Begin VB.TextBox txtPriode 
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
         Left            =   1245
         TabIndex        =   45
         Top             =   600
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtBayar 
         Height          =   330
         Left            =   1200
         TabIndex        =   55
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   171704321
         CurrentDate     =   39813
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Bayar :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   21
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angsuran Ke :"
         Height          =   195
         Index           =   13
         Left            =   90
         TabIndex        =   46
         Top             =   600
         Width           =   1020
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1335
      Left            =   4920
      TabIndex        =   37
      Top             =   3240
      Width           =   3495
      Begin VB.TextBox txtJumlah 
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
         Left            =   1245
         TabIndex        =   40
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtBayar 
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
         Left            =   1245
         TabIndex        =   39
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtSisa 
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
         Left            =   1245
         TabIndex        =   38
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jlh Pinjaman :"
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sudah Bayar :"
         Height          =   195
         Index           =   16
         Left            =   90
         TabIndex        =   42
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sisa :"
         Height          =   195
         Index           =   17
         Left            =   720
         TabIndex        =   41
         Top             =   960
         Width           =   390
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1935
      Left            =   4920
      TabIndex        =   30
      Top             =   4560
      Width           =   3495
      Begin VB.TextBox txtterlambat 
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
         Left            =   1245
         TabIndex        =   50
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtAngsuran 
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
         Left            =   1245
         TabIndex        =   33
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtDenda 
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
         Left            =   1245
         TabIndex        =   32
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtTotal 
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
         Left            =   1245
         TabIndex        =   31
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hari"
         Height          =   195
         Index           =   22
         Left            =   2040
         TabIndex        =   56
         Top             =   675
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terlambat:"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   51
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angsuran :"
         Height          =   195
         Index           =   18
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Denda :"
         Height          =   195
         Index           =   19
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   34
         Top             =   1440
         Width           =   465
      End
      Begin VB.Line Line3 
         X1              =   1080
         X2              =   3360
         Y1              =   1320
         Y2              =   1320
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Keterangan Pinjaman"
      Height          =   2295
      Left            =   120
      TabIndex        =   21
      Top             =   4200
      Width           =   4695
      Begin VB.TextBox dtTahun 
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
         Left            =   2400
         TabIndex        =   54
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox dtBulan 
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
         Left            =   1920
         TabIndex        =   53
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox dtTempo 
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
         Left            =   1440
         TabIndex        =   52
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtKet 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txtPinjam 
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
         Left            =   1440
         TabIndex        =   23
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtLama 
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
         Left            =   1440
         TabIndex        =   22
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Tempo:"
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
         Index           =   12
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan :"
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   28
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pinjaman Pokok:"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   27
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lama Angsuran :"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   26
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
         Height          =   195
         Index           =   9
         Left            =   2160
         TabIndex        =   25
         Top             =   1120
         Width           =   390
      End
   End
   Begin MSComCtl2.DTPicker dtPinjam 
      Height          =   330
      Left            =   1200
      TabIndex        =   19
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   171769857
      CurrentDate     =   39813
   End
   Begin KoperasiSys.Line Line2 
      Height          =   30
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   53
   End
   Begin VB.Frame Frame1 
      Caption         =   "- Data Anggota -"
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   4695
      Begin VB.TextBox txtNoAnggota 
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
         TabIndex        =   16
         Top             =   360
         Width           =   1020
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
         TabIndex        =   15
         Top             =   720
         Width           =   510
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
         TabIndex        =   14
         Top             =   1080
         Width           =   600
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
         TabIndex        =   13
         Top             =   1440
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "&Tambah"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   6600
      Width           =   855
   End
   Begin KoperasiSys.Line Line1 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   53
   End
   Begin MSComCtl2.DTPicker dtBukti 
      Height          =   330
      Left            =   6720
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Format          =   171769857
      CurrentDate     =   39813
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Pinjaman :"
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
      TabIndex        =   20
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "No. Pinjam :"
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
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Bukti :"
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
      Left            =   5880
      TabIndex        =   7
      Top             =   1560
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "&No Bukti :"
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
      Left            =   5880
      TabIndex        =   5
      Top             =   1200
      Width           =   780
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "frmBayar.frx":0000
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transaksi Pembayaran Pinjaman Uang Koperasi"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   6690
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "frmBayar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xTotal As Currency
Dim xSisa As Currency
Dim xJumlah As Currency
Const Ket = "Lunas"
Dim Kata As String
Dim xTempo As Integer

Sub GetNumber()

On Error GoTo Salah
    Dim Counter As String * 11
    Dim Hitung As Integer
    Dim Tgl As String
    Query "Select * from tblAngsuran order By [NoBukti]"
    Tgl = Format(Now, "dd/mm/yyyy")
    With oKoperasi
        If .RecordCount = 0 Then
            Counter = "AN-" + Right(Tgl, 2) + Mid(Tgl, 4, 2) + "0001"
        Else
           .MoveLast
            If Left(![NoBukti], 7) <> "AN-" + Right(Tgl, 2) + Mid(Tgl, 4, 2) Then
                Counter = "AN-" + Right(Tgl, 2) + Mid(Tgl, 4, 2) + "001"
            Else
                Hitung = Val(Right(!NoBukti, 4)) + 1
               Counter = "AN-" + Right(Tgl, 2) + Mid(Tgl, 4, 2) + Right("0000" & Hitung, 4)
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
        txtNoAnggota = !NoAnggota
        txtAlamat = !Alamat
        dtPinjam = !tglPinjam
        txtPekerjaan = !Pekerjaan
        txtKet = !Keterangan
        txtLama = !Lama
        txtPinjam = Format(![Pinjaman Pokok], "#,##0")
        dtTempo.Text = Format(![TglBayar], "dd")
        txtJumlah = Format(![Total Pinjaman], "#,##0")
        txtAngsuran = Format(!Angsuran, "#,##0")
        txtPriode = !Jumlah + 1
        txtSisa = Format(!JlhSisa, "#,##0")
        xTotal = Format(!Total, "#,##0")
        txtBayar = Format(!Total, "#,##0")
    End With
        dtBulan.Text = Format(Now, "mm")
        dtTahun.Text = Format(Now, "yyyy")
End Sub

Private Sub cmbPinjam_Click()
If cmbPinjam.Text <> "" Then
    Query "Select * from Qbayar Where NoPinjam ='" & cmbPinjam.Text & "'"
    If oKoperasi.EOF Then
        oKoperasi.Close
        
        Set oKoperasi = Nothing
            Query "Select * from QPinjam Where NoPinjam='" & cmbPinjam.Text & "'"
            If oKoperasi.EOF Then
                oKoperasi.Close
                Set oKoperasi = Nothing
                MsgBox " No Peminjaman { " & cmbPinjam.Text & " } Tidak terdaftar...", vbCritical
                cmbPinjam.SetFocus
                Exit Sub
            End If
            
            With oKoperasi
                txtNama = !Nama
                txtAlamat = !Alamat
                txtPekerjaan = !Pekerjaan
                txtKet = !Keterangan
                dtPinjam = !tglPinjam
                txtLama = !Lama
                txtPinjam = Format(![Pinjaman Pokok], "#,##0")
                txtJumlah = Format(![Total Pinjaman], "#,##0")
                txtAngsuran = Format(!Angsuran, "#,##0")
                txtSisa = Format(![Total Pinjaman], "#,##0")
                txtNoAnggota = !NoAnggota: dtTempo.Text = Format(!TglBayar, "dd")
                txtDenda = 0
                .Close
                dtBulan.Text = Format(Now, "mm")
                dtTahun.Text = Format(Now, "yyyy")
            End With
            txtPriode.Text = 1
            txtDenda = 0
            txtBayar = Format(txtAngsuran, "#,##0")
            xSisa = Format(Val(Int(txtSisa.Text)) - Val(Int(txtAngsuran.Text)), "#,##0")
            xTotal = txtAngsuran.Text
            dtBayar.Enabled = True
            dtBayar.Value = Format(Now, "dd/mm/yyyy")
            dtBayar.SetFocus
            Exit Sub
        End If
        Daftar
        xTotal = Format(Val(Int(txtBayar.Text)) + Val(Int(txtAngsuran.Text)), "#,##0")
        xSisa = Format(Val(Int(txtSisa.Text)) - Val(Int(txtAngsuran.Text)), "#,##0")
        dtBayar.SetFocus
        oKoperasi.Close
        Set oKoperasi = Nothing
        If Val(txtLama.Text) < Val(txtPriode.Text) Then
            MsgBox " Pembayaran Angsuran Bernomor Anggota = " & txtNoAnggota & " Sudah Lunas...", vbInformation
            xUpdate
            Semula
            Exit Sub
        End If
End If
                        
End Sub

Private Sub cmbPinjam_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmbPinjam.Text <> "" Then
    Query "Select * from tblPinjaman Where NoPinjam='" & cmbPinjam.Text & "'"
    If oKoperasi.EOF Then
        oKoperasi.Close
        Set oKoperasi = Nothing
        MsgBox " No Pinjaman = " & cmbPinjam.Text & " Tidak Terdaftar...", vbCritical
        cmbPinjam.SetFocus
        Exit Sub
    End If
    Daftar
    dtBayar.SetFocus
    oKoperasi.Close
    Set oKoperasi = Nothing
End If
End If
End Sub

Private Sub cmdCari_Click()
Kata = InputBox("Masukkan No Bukti Yang akan dicari..", "Seacrh...")
If Kata = "" Then Exit Sub
    oKoperasi.Open "Select * from QBayar Where NoBukti='" & Kata & "'", cnKoperasi, adOpenDynamic, adLockPessimistic
    If Not oKoperasi.EOF Then
        With oKoperasi
            cmbPinjam.Text = !NoPinjam: dtBukti = !tglBukti
            txtNomor = !NoBukti: dtPinjam.Value = !tglPinjam
            txtNoAnggota = !NoAnggota: txtNama = !Nama
            txtAlamat = !Alamat: txtPekerjaan = !Pekerjaan
            txtKet = !Keterangan: txtPinjam = Format(![Pinjaman Pokok], "#,##0")
            dtTempo.Text = Format(!TglBayar, "dd")
            dtBulan.Text = Format(Now, "mm"): dtTahun.Text = Format(Now, "yyyy")
            txtBayar.Text = Format(!Bayar, "#,##0")
            txtJumlah = Format(![Total Pinjaman], "#,##0")
            txtAngsuran = Format(!Angsuran, "#,##0")
            txtPriode = !Priode
            txtSisa = Format(!Sisa, "#,##0")
            xTotal = Format(!Total, "#,##0")
            txtterlambat = !Terlambat: txtDenda = Format(!Denda, "#,##0")
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
dtBayar.Enabled = True
dtBayar.SetFocus
cmdSimpan.Enabled = True
cmdEdit.Enabled = False
cmdHapus.Enabled = False
cmdCari.Enabled = False
End Sub

Private Sub cmdHapus_Click()
On Error GoTo Salah
    Kata = MsgBox("Anda Yakin Untuk Menghapus Data Angsuran = " & txtNoAnggota & " ini...", vbCritical + vbYesNo)
    If Kata = vbYes Then
        cnKoperasi.Execute "Delete From tblAngsuran Where NoBukti='" & txtNomor.Text & "'"
        MsgBox " Data Angsuran Telah Di Hapus...", vbInformation
        Semula
        Exit Sub
    End If
    Exit Sub
Salah:
    MsgBox " Data Angsuran Tidak Dapat Di Hapus.." & Chr(10) & _
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
    dtBukti.Enabled = True
    dtBukti.Value = Format(Now, "dd/mm/yyyy")
    dtBukti.SetFocus
    DaftarPinjam
Else
    Semula
End If
End Sub

Private Sub cmdTutup_Click()
Unload Me
End Sub

Sub DaftarPinjam()
    Me.MousePointer = 11
    Query "Select * from tblPinjaman Where Status <> '" & Ket & "'"
    cmbPinjam.Clear
    If Not oKoperasi.EOF Then
        oKoperasi.MoveFirst
        Do While Not oKoperasi.EOF
            cmbPinjam.AddItem oKoperasi!NoPinjam
            oKoperasi.MoveNext
        Loop
    End If
    oKoperasi.Close
    Set oKoperasi = Nothing
    Me.MousePointer = 1
End Sub

Function Semula()
    ClearControl Me
    cmdSimpan.Enabled = False
    cmdTambah.Caption = "&Tambah"
    cmdTambah.SetFocus
    blnEdit = False
End Function

Private Sub dtBayar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   xTempo = Day(dtBayar.Value)
    txtterlambat.Text = xTempo - dtTempo
    Fokus txtDenda
    'xSisa = Format(Val(Int(txtSisa.Text)) - Val(Int(txtBayar.Text)), "#,##0")
End If
End Sub

Private Sub dtBukti_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmbPinjam.Enabled = True
    cmbPinjam.SetFocus
End If
End Sub

Private Sub Form_Load()
Ketengah Me
End Sub

Private Sub txtDenda_Change()
On Error Resume Next
If txtDenda.Text = "" Then
    txtDenda.Text = 0
Else
txtDenda.Text = Format(txtDenda, "#,#"): SendKeys "{end}"
txtTotal.Text = Format(Val(Int(txtAngsuran.Text)) + Val(Int(txtDenda.Text)), "#,##0"): SendKeys "{end}"

End If
End Sub

Private Sub txtDenda_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If txtDenda <> "" Then
            cmdSimpan.Enabled = True
            cmdSimpan.SetFocus
        End If
        Exit Sub
    'End If
End If
End Sub

Function xLunas() As Boolean
    xLunas = False
    If Val(txtLama) <= Val(txtPriode) Then
        xLunas = True
    Else
        xLunas = False
    End If
End Function

Function Save()
    On Error GoTo Salah
Me.MousePointer = 11
    cnKoperasi.Execute "Insert Into tblAngsuran Values ('" & txtNomor.Text & "'," & _
    "'" & dtBukti.Value & "','" & cmbPinjam.Text & "','" & txtBayar.Text & "'," & _
    "'" & xSisa & "','" & txtPriode.Text & "','" & xTotal & "'," & _
    "'" & xSisa & "','" & txtterlambat.Text & " ','" & txtDenda.Text & "'," & _
    "'" & txtPriode.Text & "')"
    
    cnKoperasi.Execute "Update tblAngsuran Set Total='" & xTotal & "',JlhSisa='" & xSisa & "', Jumlah='" & _
    txtPriode.Text & "' Where NoPinjam='" & cmbPinjam.Text & "'"
    
    xUpdate
    Semula
    Me.MousePointer = 1
    Exit Function
Salah:
    MsgBox Err.Description & Err.Number
    Me.MousePointer = 1
End Function

Function Edit()
    On Error GoTo Salah
    cnKoperasi.Execute "Update tblAngsuran Set Terlambat='" & _
    txtterlambat.Text & "',Denda='" & txtDenda.Text & "' Where NoBukti='" & txtNomor.Text & "'"
    
    cnKoperasi.Execute "Update "
    Semula
    Exit Function
Salah:
    MsgBox Err.Description
End Function
    

Sub xUpdate()
     If xLunas Then
        cnKoperasi.Execute "Update tblPinjaman Set Status ='" & Ket & "' Where NoPinjam='" & cmbPinjam.Text & "'"
    End If
End Sub
