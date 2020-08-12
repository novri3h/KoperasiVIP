VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiUtama 
   BackColor       =   &H8000000C&
   Caption         =   "Mini Koperasi [SISP] Versi 1.0.0"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdiUtama.frx":0000
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2940
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   11
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Text            =   "Mini Koperasi System © 2018 - 2019 Nadhif Studio"
            TextSave        =   "Mini Koperasi System © 2018 - 2019 Nadhif Studio"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "mdiUtama.frx":3D97D6
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "mdiUtama.frx":3D9D70
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "Date"
            TextSave        =   "Date"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "23/12/2018"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "mdiUtama.frx":3DA10A
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   2460
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   4339
      ButtonWidth     =   1931
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Anggota"
            Object.ToolTipText     =   "Data Anggota"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pinjaman"
            Object.ToolTipText     =   "Data Pinjaman"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Simpanan"
            Object.ToolTipText     =   "Data Simpanan "
            ImageIndex      =   19
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pembayaran"
            Object.ToolTipText     =   "Pembayaran Pinjaman"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Laporan"
            Object.ToolTipText     =   "Laporan-laporan"
            ImageIndex      =   8
            Style           =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Tutup"
            Object.ToolTipText     =   "Keluar"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin Crystal.CrystalReport CrtLaporan 
         Left            =   1560
         Top             =   2280
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   600
         Top             =   2160
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3DA4A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3DBE36
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3DCB12
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3DE4A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3DFE36
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3E17C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3E315A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3E3E34
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3E4B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3E57E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3E64C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3E71A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3E7A7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3E8758
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3E9434
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3EA110
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3EA9F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3EB6D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3EBFAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3ECC88
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3EE61C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiUtama.frx":3EFFB0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "&Setup"
      Begin VB.Menu mnuAnggota 
         Caption         =   "&Anggota Koperasi"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuProses 
      Caption         =   "&Proses"
      Begin VB.Menu mnuPinjaman 
         Caption         =   "1. P&injaman"
      End
      Begin VB.Menu mnuSimpanan 
         Caption         =   "2. &Simpanan"
      End
      Begin VB.Menu mnuBayar 
         Caption         =   "3. Pem&bayaran"
      End
   End
   Begin VB.Menu mnuLaporan 
      Caption         =   "&Laporan"
      Begin VB.Menu mnuDaftarAnggota 
         Caption         =   "a. Daftar Anggota"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLapPinjaman 
         Caption         =   "b. Laporan Pinjaman"
      End
      Begin VB.Menu mnuLapSimpanan 
         Caption         =   "c. Laporan Simpanan"
      End
      Begin VB.Menu mnuLapPembayaran 
         Caption         =   "d. Laporan Pembayaran"
      End
   End
End
Attribute VB_Name = "mdiUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Me.Caption = "Mini Koperasi System"
End Sub

Private Sub mnuAnggota_Click()
frmAnggota.Show
End Sub

Private Sub mnuBayar_Click()
frmBayar.Show
End Sub

Private Sub mnuDaftarAnggota_Click()
Me.MousePointer = 11

With CrtLaporan
        .Reset
        .ReportFileName = App.Path & "\Lap Anggota.rpt"
        .DataFiles(0) = App.Path & "\dbMaster.mdb"
        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowParentHandle = mdiUtama.hWnd
        .WindowState = crptMaximized
        .Action = 1
        
    End With
Me.MousePointer = 1
End Sub

Private Sub mnuLapPembayaran_Click()
Me.MousePointer = 11

With CrtLaporan
        .Reset
        .ReportFileName = App.Path & "\Lap Bayar.rpt"
        .DataFiles(0) = App.Path & "\dbMaster.mdb"
        .Formulas(0) = "Ket='" & " Bulan: " & Space(2) & Format(Now, "mmmm/yyyy") & "'"
        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowParentHandle = mdiUtama.hWnd
        .WindowState = crptMaximized
        .Action = 1
        
    End With
Me.MousePointer = 1
End Sub

Private Sub mnuLapPinjaman_Click()
Me.MousePointer = 11

With CrtLaporan
        .Reset
        .ReportFileName = App.Path & "\Lap Pinjam.rpt"
        .DataFiles(0) = App.Path & "\dbMaster.mdb"
        .Formulas(0) = "Ket='" & " Bulan: " & Space(2) & Format(Now, "mmmm/yyyy") & "'"
        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowParentHandle = mdiUtama.hWnd
        .WindowState = crptMaximized
        .Action = 1
        
    End With
Me.MousePointer = 1
End Sub

Private Sub mnuLapSimpanan_Click()
Me.MousePointer = 11

With CrtLaporan
        .Reset
        .ReportFileName = App.Path & "\Lap Simpanan.rpt"
        .DataFiles(0) = App.Path & "\dbMaster.mdb"
        .Formulas(0) = "Ket='" & " Bulan: " & Space(2) & Format(Now, "mmmm/yyyy") & "'"
        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowParentHandle = mdiUtama.hWnd
        .WindowState = crptMaximized
        .Action = 1
        
    End With
Me.MousePointer = 1
End Sub

Private Sub mnuPinjaman_Click()
frmPinjaman.Show
End Sub

Private Sub mnuSimpanan_Click()
frmSimpanan.Show
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(5).Text = Format(Now, "HH:MM:SS")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Index
    Case 2
        frmAnggota.Show
    Case 4
        frmPinjaman.Show
    Case 5
        frmSimpanan.Show
    Case 6
        frmBayar.Show
    Case 10
        Closedb
End Select
End Sub

Private Sub Toolbar1_ButtonDropDown(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Index
    Case 8
    PopupMenu mnuLaporan, 4
End Select
End Sub


