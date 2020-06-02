VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Supermarket X"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   Picture         =   "Menu.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6720
      Top             =   5880
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFC0C0&
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   4620
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "SIM  -  Supermarket X"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   0
         Width           =   10095
      End
   End
   Begin VB.Menu mf 
      Caption         =   "&File"
      Begin VB.Menu FEDB 
         Caption         =   "Entry Data Barang"
      End
      Begin VB.Menu fgrs 
         Caption         =   "-"
      End
      Begin VB.Menu FEBB 
         Caption         =   "Entry Pembelian Barang"
      End
      Begin VB.Menu fgrs1 
         Caption         =   "-"
      End
      Begin VB.Menu FEJB 
         Caption         =   "Entry Penjualan Barang"
      End
   End
   Begin VB.Menu lap 
      Caption         =   "&Laporan"
      Begin VB.Menu LBB 
         Caption         =   "Lap. Pembelian Bulanan"
      End
      Begin VB.Menu lgrs 
         Caption         =   "-"
      End
      Begin VB.Menu LJH 
         Caption         =   "Lap. Penjualan Harian"
      End
      Begin VB.Menu lgrs1 
         Caption         =   "-"
      End
      Begin VB.Menu LSAB 
         Caption         =   "Lap. Stock Akhir Barang"
      End
   End
   Begin VB.Menu WDW 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu WTH 
         Caption         =   "Tile Horizontal"
      End
      Begin VB.Menu WTV 
         Caption         =   "Tile Vertikal"
      End
      Begin VB.Menu WCC 
         Caption         =   "Cascade"
      End
   End
   Begin VB.Menu Klr 
      Caption         =   "&Keluar"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ket As String
Private Sub FEBB_Click()
Form2.Show
Form2.SetFocus
End Sub

Private Sub FEDB_Click()
Form1.Show
Form1.SetFocus
End Sub

Private Sub FEJB_Click()
Form3.Show
Form3.SetFocus
End Sub

Private Sub Klr_Click()
End
End Sub

Private Sub LBB_Click()
Form4.Show
Form4.SetFocus
End Sub

Private Sub LJH_Click()
Form5.Show
Form5.SetFocus
End Sub

Private Sub LSAB_Click()
Form6.Show
Form6.SetFocus
End Sub

Private Sub WCC_Click()
Me.Arrange vbCascade
End Sub

Private Sub WTH_Click()
Me.Arrange vbTileHorizontal
End Sub

Private Sub WTV_Click()
Me.Arrange vbTileVertical
End Sub

Private Sub MDIForm_Activate()
N = 1
K = 1
ket = "  *  " & Label1.Caption
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
ket = Right(ket, Len(ket) - 1) & Left(ket, 1)
Label1.Caption = ket
End Sub
