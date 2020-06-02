VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supermarket X"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tampil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2265
      TabIndex        =   1
      Top             =   1755
      Width           =   1935
   End
   Begin VB.PictureBox Crys1 
      Height          =   480
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   120
      Width           =   1200
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   120
      X2              =   4560
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   120
      X2              =   4560
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Input Bulan      :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nama Pimpinan :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   510
      TabIndex        =   5
      Top             =   1755
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Laporan Pembelian Barang Bulanan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cbln = Month(Text1.Text)
cthn = Year(Text1.Text)
Crys1.ReplaceSelectionFormula "month({pembelian.tglbeli})= " & cbln & "And year({pembelian.tglbeli})= " & cthn
Crys1.Formulas(0) = "kepada = '" & Text2.Text & "'"
Crys1.WindowTitle = " LAPORAN PEMBELIAN BARANG BULANAN"
Crys1.WindowState = crptMaximized
Crys1.Action = 1
End Sub

Private Sub Command2_Click()
psn = MsgBox(" Yakin Akan Keluar ..?", vbCritical + vbYesNo, "PESAN")
If psn = vbYes Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Crys1.ReportFileName = (App.Path & "\lapbelhar.rpt")
Crys1.DataFiles(0) = (App.Path & "\Suzuya.mdb")
Crys1.WindowState = crptMaximized
Text1.Text = Format(Date, "mm/yyyy")
End Sub
