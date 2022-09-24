VERSION 5.00
Begin VB.Form FDestroyer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FeeLCoMz Destroyer v1.x.x"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main Form.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   458
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCari 
      Caption         =   "Scan"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      ToolTipText     =   "Untuk mencari virus dalam Harddisk atau Removable Disk"
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      ToolTipText     =   "Keluar dari program"
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdHajar 
      Caption         =   "Hajar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      ToolTipText     =   "Melumpuhkan Virus dari memory dan memperbaiki Registry"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Info 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   120
      Width           =   5535
   End
   Begin VB.ListBox ListProses 
      Height          =   2595
      ItemData        =   "Main Form.frx":08CA
      Left            =   240
      List            =   "Main Form.frx":08CC
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   5415
   End
End
Attribute VB_Name = "FDestroyer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ********************
' FeeLCoMz Destroyer
' By RoNz
' Des 2006 - Agu 2009
' ********************

Dim i As Integer

Private Sub Form_Load()

    ' *** Inisialisasi Variabel ***
    
    Me.Caption = Nama_Aplikasi & " By RoNz"
    
    RefreshDaftarWindow Me, ListProses
    
End Sub

Private Sub cmdHajar_Click()
    
    Hajar_Virus
    
End Sub

Private Sub cmdCari_Click()

    FCariVirus.Show

End Sub

Private Sub cmdKeluar_Click()
    
    Unload FCariVirus
    Unload Me
    
End Sub

Private Sub cmdAbout_Click()

MsgBox Nama_Aplikasi & vbCrLf & _
      "" & vbCrLf & _
      Copyright & " (novicalz@gmail.com)" & vbCrLf & _
      "FeeLCoMz Community"

End Sub

Private Sub Judul_Click()

    RefreshDaftarWindow FDestroyer, ListProses
    
End Sub

Private Sub Judul_DblClick()

    Info.Visible = False
    ListProses.Visible = True
    
End Sub

Private Sub ListProses_DblClick()

    Info.Visible = True
    ListProses.Visible = False
    
End Sub
