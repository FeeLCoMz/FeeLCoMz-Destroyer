VERSION 5.00
Begin VB.Form FCariVirus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan Virus FeeLCoMz"
   ClientHeight    =   4335
   ClientLeft      =   2070
   ClientTop       =   1575
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Cari Virus Form.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   440
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox DaftarPath 
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FF80&
      Height          =   2400
      ItemData        =   "Cari Virus Form.frx":000C
      Left            =   240
      List            =   "Cari Virus Form.frx":000E
      TabIndex        =   7
      ToolTipText     =   "Daftar folder yang terinfeksi Virus Brontok"
      Top             =   240
      Width           =   4305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5400
      TabIndex        =   6
      ToolTipText     =   "Keluar dari Scan Virus"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   3375
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Pilih Drive yang akan di-scan"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "Scan"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2160
         TabIndex        =   0
         ToolTipText     =   "Klik untuk memulai pencarian Virus"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6375
      Begin VB.ListBox DaftarFile 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   2400
         ItemData        =   "Cari Virus Form.frx":0010
         Left            =   4440
         List            =   "Cari Virus Form.frx":0012
         TabIndex        =   2
         ToolTipText     =   "Daftar file Virus yang telah tersebar"
         Top             =   240
         Width           =   1785
      End
   End
   Begin VB.Label Info 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Klik Scan untuk memulai pencarian Virus!"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   6375
   End
End
Attribute VB_Name = "FCariVirus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *** Event Program ***

Private Sub cmdCari_Click()

    Dim t As Date
    Dim DriveYgDicari As String, FileYgDicari As String
    Dim JmlFile As Integer, JmlDir As Integer
    
    t = Now
       
    Me.MousePointer = vbHourglass
    
    DaftarFile.Clear
    
    DriveYgDicari = Left(Drive1.Drive, 2)
    FileYgDicari = "*.exe"
    
    Cari_Virus DriveYgDicari, FileYgDicari, JmlFile, JmlDir
    
    Me.MousePointer = vbDefault
    
    If JmlFile > 0 Then
        Info.Caption = "Scanning Virus selesai! " & JmlFile & " Virus telah dihapus! Waktu Proses : " & DateDiff("s", t, Now) & " detik"
        MsgBox "Scanning Virus selesai! " & JmlFile & " Virus telah dihapus!"
    Else
        Info.Caption = "Scanning Virus selesai! Tidak ada Virus di " & DriveYgDicari & ". Waktu Proses : " & DateDiff("s", t, Now) & " detik"
        MsgBox "Scanning Virus selesai! Tidak ada Virus di " & DriveYgDicari
    End If
    
End Sub

Private Sub cmdOK_Click()

    FDestroyer.SetFocus
    Unload Me
    
End Sub

Private Sub DaftarFile_Click()

    DaftarPath.ListIndex = DaftarFile.ListIndex
    
End Sub

Private Sub DaftarPath_Click()

    DaftarFile.ListIndex = DaftarPath.ListIndex
    
End Sub

Private Sub Form_Load()

    Me.Icon = FDestroyer.Icon
    Drive1.ListIndex = 1

End Sub
