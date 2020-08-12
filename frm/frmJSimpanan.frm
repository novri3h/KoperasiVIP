VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmJSimpanan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Jenis Simpanan"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
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
   ScaleHeight     =   5130
   ScaleWidth      =   5220
   Begin KoperasiSys.Line Line1 
      Height          =   30
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   53
   End
   Begin VB.TextBox txtId 
      Height          =   285
      Left            =   20
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtJenis 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1320
      Width           =   3855
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "&Tambah"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4680
      Width           =   855
   End
   Begin MSComctlLib.ListView Lv1 
      Height          =   2400
      Left            =   15
      TabIndex        =   7
      Top             =   2055
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   4233
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Id Simpanan"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Pinjaman"
         Object.Width           =   6879
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   120
      Picture         =   "frmJSimpanan.frx":0000
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Daftar"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1770
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Master Jenis Simpanan Uang"
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
      TabIndex        =   10
      Top             =   240
      Width           =   4065
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   5175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   1680
      Width           =   5175
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000040C0&
      FillColor       =   &H00FFECCE&
      FillStyle       =   0  'Solid
      Height          =   2445
      Left            =   0
      Top             =   2040
      Width           =   5175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "&Id Pinjaman:"
      Height          =   195
      Index           =   0
      Left            =   20
      TabIndex        =   9
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Jenis :"
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   8
      Top             =   1080
      Width           =   3825
   End
End
Attribute VB_Name = "frmJSimpanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
