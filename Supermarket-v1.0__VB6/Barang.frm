VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supermarket X"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Entry Data Barang"
      Height          =   3135
      Left            =   960
      TabIndex        =   13
      Top             =   1080
      Width           =   5175
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Harga"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   20
         Top             =   2160
         Width           =   435
      End
      Begin VB.Label satuan_barang 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Satuan"
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   1680
         Width           =   510
      End
      Begin VB.Label kode_barang 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Kode Barang"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   480
         Width           =   930
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   135
         Left            =   1560
         TabIndex        =   17
         Top             =   240
         Width           =   15
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   135
         Left            =   480
         TabIndex        =   16
         Top             =   2280
         Width           =   15
      End
      Begin VB.Label Nama_barang 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nama Barang"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Stock Awal"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   14
         Top             =   2640
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pilihan :"
      Height          =   1335
      Left            =   360
      TabIndex        =   8
      Top             =   4320
      Width           =   6375
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5040
         MaskColor       =   &H00C0C0FF&
         Picture         =   "Barang.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2640
         MaskColor       =   &H00C0C0FF&
         Picture         =   "Barang.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdedit 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1440
         MaskColor       =   &H00C0C0FF&
         Picture         =   "Barang.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         MaskColor       =   &H00C0C0FF&
         Picture         =   "Barang.frx":114E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Abort"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3840
         MaskColor       =   &H00C0C0FF&
         Picture         =   "Barang.frx":1A18
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   6615
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Entry Data Barang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1800
         TabIndex        =   7
         Top             =   120
         Width           =   3135
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      X1              =   240
      X2              =   6720
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CnSuzuya As ADODB.Connection
Dim CommBar As ADODB.Command
Dim rsBar As ADODB.Recordset
Dim StrSql As String
Dim psn As Byte

Private Sub Form_Load()
Set CnSuzuya = New ADODB.Connection
Set CommBar = New ADODB.Command
Set rsBar = New ADODB.Recordset

With CnSuzuya
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\superm.mdb" & ";Persist Security Info=False"
    .Open
End With

With CommBar
    .ActiveConnection = CnSuzuya
    .CommandText = "select * from Barang"
    .CommandType = adCmdText
End With
With rsBar
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open CommBar
End With
Text1.MaxLength = 8
Text2.MaxLength = 30
Text3.MaxLength = 15

End Sub


Private Sub cmdedit_click()
If Not Len(Trim(Text1.Text)) = 0 Then
     rsBar.Filter = "KdBar = '" & Trim(Text1.Text) & "'"
     If rsBar.RecordCount > 0 Then
        psn = MsgBox("Data Barang Akan Disimpan ..?", vbInformation + vbYesNo, "PESAN")
        If psn = vbYes Then
            rsBar(0) = Text1.Text
            rsBar(1) = Text2.Text
            rsBar(2) = Text3.Text
            rsBar(3) = Text4.Text
            rsBar(4) = Text5.Text
            rsBar(5) = Text5.Text
            rsBar.Update
        End If
    Else
        On Error Resume Next
        psn = MsgBox("Data Barang Belum Dientrykan ", vbExclamation)
    End If
    Call blankform 'kosongkan form
Else
  MsgBox "Masukkan Kode Barang Yang Akan Diedit Ok..", , "Pesan"
  Text1.SetFocus
End If
End Sub

Private Sub Cmddelete_Click()
If Not Len(Trim(Text1.Text)) = 0 Then
    rsBar.Filter = "KdBar = '" & Trim(Text1.Text) & "'"
     If rsBar.RecordCount > 0 Then
        psn = MsgBox("Data Barang Akan Dihapus ..?", vbInformation + vbYesNo, "PESAN")
        If psn = vbYes Then
            rsBar.Delete
            rsBar.MoveNext
        End If
    Else
        On Error Resume Next
        psn = MsgBox("Data Barang Belum Dientrykan ", vbInformation)
    End If
Else
    psn = MsgBox("Masukkan Kode Barang Yang Akan Dihapus Ok.. ", vbInformation, "Pesan")
    Text1.SetFocus
End If
Call blankform
End Sub
Private Sub blankform()
Text2 = ""
Text3.Text = ""
Text4 = ""
End Sub
Private Sub cmdexit_Click()
psn = MsgBox(" Yakin Akan Keluar ..?", vbCritical + vbYesNo, "barang")
If psn = vbYes Then
    Unload Me
End If
End Sub

Private Sub cmdsave_click()
If Not Len(Trim(Text1.Text)) = 0 Then
    'rsBar.Filter = "KdBar = '" & Trim(Text1.Text) & "'"
    ' If rsBar.RecordCount > 0 Then
    '    psn = MsgBox("Data Barang Sudah Pernah di Entrykan..", vbInformation + vbOKOnly, "PESAN")
    'Else
        psn = MsgBox("Data Barang Akan Disimpan ..?", vbInformation + vbYesNo, "PESAN")
        If psn = vbYes Then
            rsBar.AddNew
            rsBar(0) = Text1.Text
            rsBar(1) = Text2.Text
            rsBar(2) = Text3.Text
            rsBar(3) = Text4.Text
            rsBar(4) = Text5.Text
            rsBar(5) = Text5.Text
            rsBar.Update 'simpan
        End If
        If Err.Number <> 0 Then
            If rsBar.EditMode = adEditAdd Then
                MsgBox "Gagal Tambah Record", vbCritical & Err.Description
            Else
                MsgBox "Gagal Ganti Record", vbCritical & Err.Description
            End If
            rsBar.CancelUpdate
        End If
        Call blankform 'kosongkan form
    'End If
Else
  MsgBox "Masukkan Data Terlebih dahulu...ya? ", , "Pesan"
  Text1.SetFocus
End If
End Sub

Private Sub Command1_Click()
blankform
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
     rsBar.Filter = "KdBar = '" & Trim(Text1.Text) & "'"
     If rsBar.RecordCount > 0 Then
        Text2.Text = rsBar.Fields(1)
        Text3.Text = rsBar.Fields(2)
        Text4.Text = rsBar.Fields(3)
        Text5.Text = rsBar.Fields(4)
        psn = MsgBox("Data Barang Sudah Dientrykan ", vbExclamation)
   End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
If KeyAscii = 27 Then Unload Me
End Sub
