VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Supermarket X"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Koreksi Data Barang"
      Height          =   1815
      Left            =   5760
      TabIndex        =   1
      Top             =   2760
      Width           =   4575
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Harga Barang"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "kode Barang"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nama Barang"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   810
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   11295
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Transaksi Pembelian Barang"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   31
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   3135
      Left            =   5520
      TabIndex        =   22
      Top             =   1080
      Width           =   5895
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   810
         ItemData        =   "Pembelian.frx":0000
         Left            =   360
         List            =   "Pembelian.frx":0002
         TabIndex        =   23
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1800
         TabIndex        =   29
         Top             =   750
         Width           =   2535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nama Barang"
         Height          =   195
         Left            =   360
         TabIndex        =   28
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Kode Barang"
         Height          =   195
         Left            =   360
         TabIndex        =   27
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Barang yg Pilihan Pilihan  :"
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
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Harga Barang"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   25
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1800
         TabIndex        =   24
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   5295
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Jumlah Beli"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Height          =   1935
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   5295
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
         TabIndex        =   5
         Top             =   810
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   16744576
         Format          =   62259201
         CurrentDate     =   38465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tanggal"
         Height          =   195
         Left            =   225
         TabIndex        =   19
         Top             =   870
         Width           =   585
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No Faktur"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pilihan :"
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   6255
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
         Height          =   975
         Left            =   3840
         MaskColor       =   &H00C0C0FF&
         Picture         =   "Pembelian.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1095
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
         Height          =   975
         Left            =   240
         MaskColor       =   &H00C0C0FF&
         Picture         =   "Pembelian.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1095
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
         Height          =   975
         Left            =   1440
         MaskColor       =   &H00C0C0FF&
         Picture         =   "Pembelian.frx":0BD8
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1095
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
         Height          =   975
         Left            =   2640
         MaskColor       =   &H00C0C0FF&
         Picture         =   "Pembelian.frx":101A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
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
         Height          =   975
         Left            =   5040
         MaskColor       =   &H00C0C0FF&
         Picture         =   "Pembelian.frx":18E4
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CnSuzuya As ADODB.Connection
Dim CommBar As ADODB.Command
Dim rsBar As ADODB.Recordset
Dim CommBeli As ADODB.Command
Dim rsBeli As ADODB.Recordset

Dim psn As Byte

Private Sub Form_Load()
Set CnSuzuya = New ADODB.Connection
Set CommBar = New ADODB.Command
Set rsBar = New ADODB.Recordset
Set CommBeli = New ADODB.Command
Set rsBeli = New ADODB.Recordset


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

With CommBeli
    .ActiveConnection = CnSuzuya
    .CommandText = "select * from Pembelian"
    .CommandType = adCmdText
End With
With rsBeli
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open CommBeli
End With
Frame2.Visible = False
Text1.MaxLength = 8
End Sub

Private Sub Command1_Click()
If Text4.Text = "" Then
    On Error Resume Next
    psn = MsgBox("Jumlah Barang Yang di Beli Belum Dientrykan", vbExclamation)
Else
     rsBeli.Filter = "nofak = '" & Trim(Text1.Text) & "' and KdBar = '" & Trim(Combo1.Text) & "'"
     If rsBeli.RecordCount > 0 Then
        On Error Resume Next
        psn = MsgBox("Data Barang Sudah Ada Dalam Pesanan", vbExclamation, "PESAN")
    Else
        psn = MsgBox("Data Barang diBeli Akan Disimpan ..?", vbInformation + vbYesNo, "PersediaanAN")
        If psn = vbYes Then
            rsBeli.AddNew 'sediakan tempat di database untuk memasukkan data
            rsBeli(0) = DTPicker1.Value
            rsBeli(3) = Val(Text4.Text)
            rsBeli(2) = Combo1.Text
            rsBeli(1) = Text1.Text 'isi kode Persediaan dari text1
            rsBeli.Update 'simpan
            List1.AddItem Label7
            a = 0
            rsBar.Filter = "KdBar = '" & Trim(Combo1.Text) & "'"
            If rsBar.RecordCount > 0 Then
                On Error Resume Next
                c = rsBar(5)
                a = Text4.Text + c
                'rsBar.edit
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
If Str(Text4.Text) = "" Then
    On Error Resume Next
    psn = MsgBox("Jumlah Barang Masuk Belum Dientrykan", vbExclamation)
Else
    rsBeli.Filter = "nofak = '" & Trim(Text1.Text) & "' and KdBar = '" & Trim(Combo1.Text) & "'"
    If rsBeli.RecordCount > 0 Then
        psn = MsgBox("Data Barang Akan Disimpan ..?", vbInformation + vbYesNo, "PESAN")
        If psn = vbYes Then
            a = 0
          '  rsBar.Filter = "select * from barang"
            rsBar.Filter = "KdBar = '" & Trim(Combo1.Text) & "'"
            If rsBar.RecordCount > 0 Then
                On Error Resume Next
                'c = rsBar(5)
                a = rsBar(5) - Text4
                'rsBar.edit
                rsBar.Fields(5) = a
                rsBar.Update
            End If
            
            'rsBeli.edit
            rsBeli.Fields(0) = DTPicker1.Value
            If Combo3.Text = "" Then
                On Error Resume Next
                rsBeli(2) = Combo1.Text
            Else
                rsBeli(2) = Combo3.Text
            End If
            rsBeli(3) = Text4.Text
            rsBeli.Update
            a = 0
            rsBar.Filter = "KdBar = '" & Trim(Combo3.Text) & "'"
            If rsBar.RecordCount > 0 Then
                On Error Resume Next
                'c = rsBar(5)
                a = Text4 + rsBar(5)
                'rsBar.edit
                rsBar.Fields(5) = a
                rsBar.Update
            End If
            
        End If
    Else
        On Error Resume Next
        psn = MsgBox("Data Barang Belum Dientrykan", vbExclamation)
    End If
End If
Tkoraktif
End Sub

Private Sub Command3_Click()
rsBeli.Filter = "nofak = '" & Trim(Text1.Text) & "' and KdBar = '" & Trim(Combo1.Text) & "'"
If rsBeli.RecordCount > 0 Then
    psn = MsgBox("Data Barang Akan Dihapus ..?", vbExclamation + vbYesNo, "PersediaanAN")
    If psn = vbYes Then
        rsBar.Filter = "KdBar = '" & Trim(Combo1.Text) & "'"
        If rsBar.RecordCount > 0 Then
            c = rsBar(5)
            a = c - Text4
            'rsBar.edit
            rsBar.Fields(5) = a
            rsBar.Update
        End If
        rsBeli.Delete
        rsBeli.MoveNext
    End If
Else
    On Error Resume Next
    psn = MsgBox("Data Barang Belum Dientrykan", vbExclamation)
End If
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

Private Sub blankform()
Combo1.Text = ""
Combo2.Text = ""
Text4.Text = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim Persediaan As String
KeyAscii = Asc(UCase(Chr(KeyAscii)))
rsBeli.Filter = "nofak = '" & Trim(Text1.Text) & "'"
If rsBeli.RecordCount > 0 Then
      On Error Resume Next
      DTPicker1.Value = rsBeli(0)
    
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
End Sub

Private Sub Combo1_Click()
rsBar.Filter = "KdBar = '" & Trim(Combo1.Text) & "'"
If rsBar.RecordCount > 0 Then
    Label7.Caption = rsBar(1)
    Label8.Caption = rsBar(3)
End If
rsBeli.Filter = "nofak = '" & Trim(Text1.Text) & "' and KdBar = '" & Trim(Combo1.Text) & "'"
If rsBeli.RecordCount > 0 Then
    psn = MsgBox("Barang Sudah Ada", vbExclamation, "PESAN")
    Text4.Text = rsBeli.Fields(3)
Else
    On Error Resume Next
    psn = MsgBox("Barang Belum Ada", vbExclamation, "PESAN")
End If
End Sub

Private Sub Combo3_Click()
rsBar.Filter = "KdBar = '" & Trim(Combo3.Text) & "'"
If rsBar.RecordCount > 0 Then
    Text15.Text = rsBar(1)
    Label9.Caption = rsBar(3)
End If

rsBeli.Filter = "nofak = '" & Trim(Text1.Text) & "' and KdBar = '" & Trim(Combo3.Text) & "'"
If rsBeli.RecordCount > 0 Then
    psn = MsgBox("Barang Sudah Ada", vbExclamation, "PESAN")
Else
    On Error Resume Next
    psn = MsgBox("Barang Belum Ada", vbExclamation, "PESAN")
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
