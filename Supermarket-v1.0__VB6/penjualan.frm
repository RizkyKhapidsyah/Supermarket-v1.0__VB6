VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Supermarket X"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Height          =   1695
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   6495
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   810
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   16744576
         Format          =   62259201
         CurrentDate     =   38465
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No Faktur"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tanggal"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   31
         Top             =   870
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Koreksi Data Barang"
      Height          =   1815
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   4575
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nama Barang"
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "kode Barang"
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Harga Barang"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   18
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   5640
      Width           =   6495
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Jumlah Jual"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pilihan :"
      Height          =   3135
      Left            =   7440
      TabIndex        =   1
      Top             =   3240
      Width           =   3615
      Begin VB.CommandButton Command5 
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
         Height          =   1215
         Left            =   1800
         MaskColor       =   &H00C0C0FF&
         Picture         =   "penjualan.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
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
         Left            =   240
         MaskColor       =   &H00C0C0FF&
         Picture         =   "penjualan.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
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
         Left            =   240
         MaskColor       =   &H00C0C0FF&
         Picture         =   "penjualan.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
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
         Picture         =   "penjualan.frx":114E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
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
         Height          =   1095
         Left            =   1800
         MaskColor       =   &H00C0C0FF&
         Picture         =   "penjualan.frx":1A18
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1665
      ScaleWidth      =   6435
      TabIndex        =   33
      Top             =   240
      Width           =   6465
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   240
      ScaleHeight     =   945
      ScaleWidth      =   6435
      TabIndex        =   13
      Top             =   5880
      Width           =   6465
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3120
      Left            =   7560
      ScaleHeight     =   3090
      ScaleWidth      =   3675
      TabIndex        =   10
      Top             =   3495
      Width           =   3705
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   3495
      Left            =   0
      TabIndex        =   21
      Top             =   1920
      Width           =   6495
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   810
         ItemData        =   "penjualan.frx":1D22
         Left            =   1920
         List            =   "penjualan.frx":1D24
         TabIndex        =   22
         Top             =   2280
         Width           =   4335
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1800
         TabIndex        =   28
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Harga Barang"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   27
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pilihan Barang :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Kode Barang"
         Height          =   195
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nama Barang"
         Height          =   195
         Left            =   360
         TabIndex        =   24
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1800
         TabIndex        =   23
         Top             =   780
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   240
      ScaleHeight     =   3465
      ScaleWidth      =   6480
      TabIndex        =   29
      Top             =   2160
      Width           =   6510
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   7200
      Picture         =   "penjualan.frx":1D26
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CnSuzuya As ADODB.Connection
Dim CommBar As ADODB.Command
Dim rsBar As ADODB.Recordset
Dim CommJual As ADODB.Command
Dim rsJual As ADODB.Recordset
Dim psn As Byte

Private Sub Form_Load()
Set CnSuzuya = New ADODB.Connection
Set CommBar = New ADODB.Command
Set rsBar = New ADODB.Recordset
Set CommJual = New ADODB.Command
Set rsJual = New ADODB.Recordset


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

With CommJual
    .ActiveConnection = CnSuzuya
    .CommandText = "select * from Penjualan"
    .CommandType = adCmdText
End With
With rsJual
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open CommJual
End With
Frame2.Visible = False
Text1.MaxLength = 8
End Sub

Private Sub blankform()
Combo1.Text = ""
Combo2.Text = ""
Combo4.Text = ""
Text4.Text = ""
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim Persediaan As String
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    rsJual.Filter = "nofak = '" & Trim(Text1.Text) & "'"
    If rsJual.RecordCount > 0 Then
        On Error Resume Next
        DTPicker1.Value = rsJual(0)
        Command1.Enabled = False
        psn = MsgBox("Data Sudah Ada", vbExclamation)
        psn = MsgBox("Data Akan Dikoreksi ..?", vbInformation + vbYesNo, "Persediaanan")
        If psn = vbYes Then
            koraktif
            Combo1.SetFocus
        Else
            Command1.Enabled = True
        End If
        Exit Sub
    End If
End If
End Sub

Private Sub Combo1_Click()
rsBar.Filter = "KdBar = '" & Trim(Combo1.Text) & "'"
If rsBar.RecordCount > 0 Then
    Label7.Caption = rsBar(1)
    Label8.Caption = rsBar(3)
End If

rsJual.Filter = "nofak = '" & Trim(Text1.Text) & "' and KdBar = '" & Trim(Combo1.Text) & "'"
If rsJual.RecordCount > 0 Then
    psn = MsgBox("Barang Sudah Ada", vbExclamation, "PESAN")
    Text4.Text = rsJual(3)
Else
    On Error Resume Next
    psn = MsgBox("Barang Belum Ada", vbExclamation, "PESAN")
End If
End Sub

Private Sub Combo3_Click()
rsBar.Filter = "KdBar = '" & Trim(Combo1.Text) & "'"
If rsBar.RecordCount > 0 Then
    Text15.Text = rsBar(1)
    Label9.Caption = rsBar(3)
End If

rsJual.Filter = "nofak = '" & Trim(Text1.Text) & "' and KdBar = '" & Trim(Combo3.Text) & "'"
If rsJual.RecordCount > 0 Then
    psn = MsgBox("Barang Sudah Ada", vbExclamation, "PESAN")
Else
    psn = MsgBox("Barang Belum Ada", vbExclamation, "PESAN")
End If
End Sub

Private Sub Command1_Click()
If Text4.Text = "" Then
    On Error Resume Next
    psn = MsgBox("Data Jumlah Barang Keluar Belum Dientrykan", vbExclamation)
Else
    rsJual.Filter = "nofak = '" & Trim(Text1.Text) & "' and KdBar = '" & Trim(Combo1.Text) & "'"
    If rsJual.RecordCount > 0 Then
        psn = MsgBox("Data Barang Sudah Ada Dalam Pesanan", vbExclamation, "PESAN")
    Else
        On Error Resume Next
        psn = MsgBox("Data Penjualan Barang Akan Disimpan ..?", vbInformation + vbYesNo, "barangAN")
        If psn = vbYes Then
            rsJual.AddNew
            rsJual(0) = DTPicker1.Value
            rsJual(1) = Text1.Text
            rsJual(2) = Combo1.Text
            rsJual(3) = Text4.Text
            rsJual.Update
            List1.AddItem Label7
            a = 0
            rsBar.Filter = "KdBar = '" & Trim(Combo1.Text) & "'"
            If rsBar.RecordCount > 0 Then
                c = rsBar(5)
                a = c - Text4
                'rsBar.Edit
                rsBar.Fields(5) = a
                rsBar.Update
            End If
        End If
        psn = MsgBox(" Apakah Masih Ada Pesanan Lagi ..?", vbExclamation + vbYesNo, "PESAN")
        If psn = vbYes Then
            Combo1.SetFocus
        Else
            On Error Resume Next
            List1.Clear
        End If
    End If
End If
End Sub

Private Sub Command2_Click()
If Text4.Text = "" Then
    On Error Resume Next
    psn = MsgBox("Data Jumlah Barang Keluar Belum Dientrykan", vbExclamation)
Else
    rsJual.Filter = "nofak = '" & Trim(Text1.Text) & "' and KdBar = '" & Trim(Combo1.Text) & "'"
    If rsJual.RecordCount > 0 Then
        psn = MsgBox("Data Barang Akan Disimpan ..?", vbInformation + vbYesNo, "PESAN")
        If psn = vbYes Then
            a = 0
            rsBar.Filter = "KdBar = '" & Trim(Combo1.Text) & "'"
            If rsBar.RecordCount > 0 Then
                c = rsBar(5)
                a = c + Text4
                'rsBar.Edit
                rsBar.Fields(5) = a
                rsBar.Update
            End If
            'rsJual.Edit
            rsJual(0) = DTPicker1.Value
            rsJual(1) = Text1.Text
            If Combo3.Text = "" Then
                On Error Resume Next
                rsJual(2) = Combo1.Text
            Else
                rsJual(2) = Combo3.Text
            End If
            rsJual(3) = Text4.Text
            a = 0
            rsBar.Filter = "KdBar = '" & Trim(Combo3.Text) & "'"
            If rsBar.RecordCount > 0 Then
                c = rsBar(5)
                a = c - Text4
                'rsBar.Edit
                rsBar.Fields(5) = a
                rsBar.Update
            End If
        End If
    End If
End If
'blank
blankor
End Sub

Private Sub Command3_Click()
rsJual.Filter = "nofak = '" & Trim(Text1.Text) & "' and KdBar = '" & Trim(Combo1.Text) & "'"
If rsJual.RecordCount > 0 Then
    psn = MsgBox("Data Barang Keluar Akan Dihapus ..?", vbExclamation + vbYesNo, "PESAN")
    If psn = vbYes Then
        rsBar.Filter = "KdBar = '" & Trim(Combo1.Text) & "'"
        If rsBar.RecordCount > 0 Then
            c = rsBar(5)
            a = c + Text4
            'rsBar.Edit
            rsBar.Fields(5) = a
            rsBar.Update
        End If
        rsJual.Delete
        rsJual.MoveNext
    End If
Else
    psn = MsgBox("Data Barang Belum Dientrykan", vbExclamation)
    Exit Sub
End If
End Sub

Private Sub Barang()
Dim a As Integer
On Error Resume Next
rsBar.MoveFirst
Do While Not rsBar.EOF = True
    Combo1.List(a) = rsBar.Fields(0)
    Combo3.List(a) = rsBar.Fields(0)
    a = a + 1
    rsBar.MoveNext
Loop
Text1.SetFocus
End Sub

Private Sub koraktif()
Frame2.Visible = True
Combo3.Enabled = True
Text5.Enabled = True
'Combo3.SetFocus
End Sub
Private Sub Tkoraktif()
Frame2.Visible = False
Combo3.Enabled = False

'Text5.Enabled = False
End Sub
Sub blankor()
Combo3.Text = ""
Text15.Text = ""
Tkoraktif
End Sub

Private Sub Command4_Click()
blankform
End Sub

Private Sub Command5_Click()
psn = MsgBox(" Yakin Akan Keluar ..?", vbCritical + vbYesNo, "barang")
If psn = vbYes Then
    Unload Me
End If
End Sub

Private Sub Form_Activate()
Barang
End Sub
