Attribute VB_Name = "modDeklarasi"
Option Explicit
Global cnKoperasi As ADODB.Connection
Public oKoperasi As New ADODB.Recordset
Public xKata As String
Public blnEdit As Boolean
Public xCicilan As Currency
Public Const dbName = "dbMaster.mdb"


