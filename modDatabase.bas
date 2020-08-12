Attribute VB_Name = "modDatabase"
Public Function ConnectDb() As Boolean
On Error GoTo Salah
    ConnectDb = False
    Set cnKoperasi = New ADODB.Connection
        cnKoperasi.CursorLocation = adUseClient
        cnKoperasi.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" + dbName + ";Persist Security Info=False"
    ConnectDb = True
    mdiUtama.Show
    Exit Function
Salah:
    MsgBox Err.Description & Err.Number
End Function

Sub Main()
    If ConnectDb Then
        With mdiUtama
            .StatusBar1.Panels(3).Text = GetFileSize(StripPath(App.Path) & "" + dbName + " ")
            .StatusBar1.Panels(5).Text = Format(Now, "HH:MM:SS")
            .StatusBar1.Panels(9).Text = "Admin"
        End With
    End If
            
End Sub
Public Function Query(Perintah As String)
On Error GoTo Err
    If oKoperasi.State > 0 Then oKoperasi.Close
        oKoperasi.Open Perintah, cnKoperasi, _
        adOpenDynamic, adLockPessimistic
        Exit Function
Err:
        MsgBox Err.Description
End Function

Function Closedb()

Dim Form  As Form
   For Each Form In Forms
       Unload Form
       Set Form = Nothing      'Bersihkan memori yang digunakan sebelumnya
    Next Form
    
    Set xCon = Nothing  'close connecting
    
End Function

Public Sub Fokus(ByVal KotakTesk As TextBox)
    With KotakTesk
        .SelStart = 0
        .Enabled = True
        .Locked = False
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Public Sub Ketengah(ByVal Frm As Form)
    Frm.Left = (mdiUtama.Width - Frm.Width) / 2
    Frm.Top = (mdiUtama.Height - Frm.Height) / 2 - 500
End Sub

Public Function GetFileSize(File As Variant) As String
    On Error Resume Next
    Dim Bytes As Long
    Const Kb As Long = 1024
    Const Mb As Long = 1024 * Kb
    Const Gb As Long = 1024 * Mb
    Bytes = FileLen(File)
    If Bytes < Kb Then
        GetFileSize = Format(Bytes) & " bytes"
    ElseIf Bytes < Mb Then
        GetFileSize = Format(Bytes / Kb, "0.00") & " Kb"
    ElseIf Bytes < Gb Then
        GetFileSize = Format(Bytes / Mb, "0.00") & " Mb"
    Else
        GetFileSize = Format(Bytes / Gb, "0.00") & " Gb"
    End If
End Function

Function StripPath(nPath As String) As String
If Right(nPath, 1) = "\" Then
   StripPath = nPath
Else
   StripPath = nPath & "\"
End If
End Function

Sub ClearControl(Frm As Form, Optional IncludeHide As Boolean = True)
On Error Resume Next
Dim j As Control
For Each j In Frm.Controls
    If IncludeHide Then
kembali:
       If TypeOf j Is TextBox Then
          j.Text = ""
       ElseIf TypeOf j Is ComboBox Then
          j.Text = ""
       ElseIf TypeOf j Is OptionButton Then
          j.Value = False
       End If
    Else
      If j.Visible Then GoSub kembali
    End If
Next
End Sub

Sub KeyForm(Frm As Form, Optional IncludeHide As Boolean = True)
On Error Resume Next
Dim j As Control
For Each j In Frm.Controls
    If IncludeHide Then
kembali:
       If TypeOf j Is TextBox Then
          j.Enabled = False
       ElseIf TypeOf j Is ComboBox Then
          j.Enabled = False
       ElseIf TypeOf j Is OptionButton Then
          j.Enabled = False
       End If
    Else
      If j.Visible Then GoSub kembali
    End If
Next
End Sub

Public Function IsDelete(nTable As String, nField As String, nKey As String)
On Error GoTo Salah
    IsDelete = False
    If Not cnKoperasi Is Nothing Then
        If cnKoperasi.State > 0 Then
            cnKoperasi.Execute "Delete * from nTable Where nField ='" & nKey & "'"
            
            IsDelete = True
        End If
    End If
    Exit Function
Salah:
    MsgBox Err.Description & Err.Number
End Function

