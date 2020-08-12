VERSION 5.00
Begin VB.Form frmAnggota 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Anggota"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAnggota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7320
   Begin VB.TextBox Text 
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
      Index           =   5
      Left            =   1440
      TabIndex        =   19
      Top             =   2880
      Width           =   3615
   End
   Begin KoperasiSys.Line Line1 
      Height          =   30
      Left            =   0
      TabIndex        =   17
      Top             =   3240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   53
   End
   Begin KoperasiSys.Line Line2 
      Height          =   30
      Left            =   0
      TabIndex        =   16
      Top             =   960
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   14
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "&Tambah"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text 
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
      Index           =   4
      Left            =   1440
      TabIndex        =   10
      Top             =   2520
      Width           =   3615
   End
   Begin VB.TextBox Text 
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
      Index           =   3
      Left            =   1440
      TabIndex        =   8
      Top             =   2160
      Width           =   5895
   End
   Begin VB.TextBox Text 
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
      Index           =   2
      Left            =   1440
      TabIndex        =   6
      Top             =   1800
      Width           =   5895
   End
   Begin VB.TextBox Text 
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
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   1440
      Width           =   5895
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pekerjaan :"
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "No &Telp :"
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "&Kota :"
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "A&lamat : "
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "N&ama :"
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "&No. Anggota :"
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
      TabIndex        =   1
      Top             =   1080
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   120
      Picture         =   "frmAnggota.frx":038A
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Master Anggota Koperasi"
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
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmAnggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdEdit_Click()
blnEdit = True
Fokus Text(1)
cmdEdit.Enabled = False
cmdHapus.Enabled = False
End Sub

Private Sub cmdHapus_Click()
On Error GoTo ErrDelete
    cnKoperasi.Execute "Delete * from tblAnggota Where NoAnggota='" & Text(0).Text & "'"
    Semula
    Exit Sub
ErrDelete:
    MsgBox Err.Description & Err.Number
End Sub

Private Sub cmdSimpan_Click()
For i = 0 To 2
    If Text(i).Text = "" Then
        MsgBox " Data Anggota Belum Lengkap", vbCritical
        Fokus Text(i)
        Exit Sub
    End If
Next
Me.MousePointer = 11
    If Not blnEdit Then
        SimpanData
    Else
        EditData
    End If
Semula
Me.MousePointer = 1
    
End Sub

Private Sub cmdTambah_Click()
If cmdTambah.Caption = "&Tambah" Then
    cmdTambah.Caption = "&Batal"
    Fokus Text(0)
Else
    Semula
End If
End Sub

Private Sub cmdTutup_Click()
Unload Me
End Sub

Private Sub Form_Load()
Ketengah Me

End Sub

Sub Semula()
    cmdTambah.Caption = "&Tambah"
    cmdTambah.SetFocus
    cmdSimpan.Enabled = False
    cmdEdit.Enabled = False
    cmdHapus.Enabled = False
    blnEdit = False
    ClearControl Me
    
End Sub

Sub Daftar()
    With oKoperasi
        Text(1) = !Nama
        Text(2) = !Alamat
        Text(3) = !Kota
        Text(4) = !NoTelp
        Text(5) = !Pekerjaan
    End With
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case 0
        If KeyAscii = 13 Then
            cmdSimpan.Enabled = True
        Query "Select * from tblAnggota Where NoAnggota='" & Text(0).Text & "'"
        If oKoperasi.EOF Then
            oKoperasi.Close
            Set oKoperasi = Nothing
            Fokus Text(1)
            cmdSimpan.Enabled = True
            Exit Sub
        End If
        Daftar
        oKoperasi.Close
        Set oKoperasi = Nothing
        cmdEdit.Enabled = True
        cmdHapus.Enabled = True
    End If
    Case 1 To 4
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    Case 5
        If KeyAscii = 13 Then
            cmdSimpan.Enabled = True
            cmdSimpan.SetFocus
        End If
    End Select

            
End Sub

Sub SimpanData()
    On Error GoTo ErrSimpan
        If Not cnKoperasi Is Nothing Then
            If cnKoperasi.State > 0 Then
                cnKoperasi.Execute "Insert Into tbLAnggota Values ('" & Text(0).Text & "'," & _
                                                                  "'" & Text(1).Text & "'," & _
                                                                  "'" & Text(2).Text & "'," & _
                                                                  "'" & Text(3).Text & "'," & _
                                                                  "'" & Text(4).Text & "'," & _
                                                                  "'" & Text(5).Text & "')"
            End If
        End If
    Exit Sub
ErrSimpan:
    MsgBox " System Tidak Dapat melakukan penyimpanan datat...", vbCritical
    
            
End Sub

Sub EditData()
    On Error GoTo ErrEdit
        If Not cnKoperasi Is Nothing Then
            If cnKoperasi.State > 0 Then
                cnKoperasi.Execute "Update tblAnggota Set Nama='" & Text(1).Text & "',Alamat='" & _
                Text(2).Text & "',Kota='" & Text(3).Text & "',NoTelp='" & Text(4).Text & "',Pekerjaan='" & _
                Text(5).Text & "' Where NoAnggota='" & Text(0).Text & "'"
            End If
        End If
        Exit Sub
ErrEdit:
    MsgBox " System Tidak dapat melakukan Edit data Anggota...", vbCritical
    
End Sub
