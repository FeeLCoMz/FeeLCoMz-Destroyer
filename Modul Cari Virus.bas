Attribute VB_Name = "ModulCariVirus"
Option Explicit

' *** Modul Pencarian Virus ***
' *** Deklarasi Fungsi Pencarian File ***

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

' *** Konstanta ***

Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

' *** Deklarasi Type Data ***

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

' *** Fungsi Modul ***

Private Function StripNulls(OriginalStr As String) As String
    
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    
    StripNulls = OriginalStr

End Function

' *** Deklarasi Fungsi Publik ***

Public Function Cari_Virus(Path As String, FileYgDicari As String, HitungFile As Integer, HitungDir As Integer)

    Dim UkuranFile As Long
    Dim NamaFile As String 'Nama File yang sedang diproses
    Dim NamaDir As String 'Nama Sub-Direktori
    Dim BufNamaDir() As String 'Buffer untuk nama Direktori
    Dim JumlahDir As Integer 'Jumlah Direktori di Direktori sekarang
    Dim i As Integer 'Pencacah
    Dim hSearch As Long 'Handle dari Pencarian
    Dim WFD As WIN32_FIND_DATA 'Type yang didefinisikan pengguna
    Dim Cont As Integer
    Dim Jawab As VbMsgBoxResult
    
    If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    ' *** Cari Sub-Direktori ***
    JumlahDir = 0
    Cont = True
    
    ReDim BufNamaDir(JumlahDir)
    
    hSearch = FindFirstFile(Path & "*", WFD)
    
    If hSearch <> INVALID_HANDLE_VALUE Then
        
        Do While Cont
        
            NamaDir = StripNulls(WFD.cFileName)
            
            If (NamaDir <> ".") And (NamaDir <> "..") Then
                'Cek jika atributnya adalah Direktori
                If GetFileAttributes(Path & NamaDir) And FILE_ATTRIBUTE_DIRECTORY Then
                    BufNamaDir(JumlahDir) = NamaDir
                    HitungDir = HitungDir + 1
                    JumlahDir = JumlahDir + 1
                    ReDim Preserve BufNamaDir(JumlahDir)
                End If
            End If
            
            Cont = FindNextFile(hSearch, WFD) 'Cari Sub-Direktori berikutnya
        
        Loop
        
        Cont = FindClose(hSearch)
        
    End If
    
    ' Cari file di direktori sekarang dan hitung ukuran file
    hSearch = FindFirstFile(Path & FileYgDicari, WFD)
    Cont = True
    
    If hSearch <> INVALID_HANDLE_VALUE Then
        
        While Cont
        
            NamaFile = StripNulls(WFD.cFileName)
            
            If (NamaFile <> ".") And (NamaFile <> "..") Then
                
                Cari_Virus = Cari_Virus + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
                
                ' *** Tampilkan nama file di Listbox ***
                FCariVirus.Info.Caption = Path & NamaFile
                FCariVirus.Info.Refresh
                
                ' *** Apakah file ini adalah virus atau tidak ***
                UkuranFile = WFD.nFileSizeLow
                
                If (UkuranFile = UkuranVirus) Then
                    HitungFile = HitungFile + 1
                    
                    ' *** Tambahkan di ListBox ***
                    FCariVirus.Frame1.Caption = HitungFile & " Virus FeeLCoMz ditemukan di " & FCariVirus.Drive1.Drive
                    FCariVirus.DaftarPath.AddItem Path
                    FCariVirus.DaftarFile.AddItem NamaFile
                    FCariVirus.DaftarPath.Refresh
                    FCariVirus.DaftarFile.Refresh
                    
                    ' *** Hapus Virusnya ***
                    Jawab = MsgBox(Path & NamaFile & " dicurigai sebagai Virus FeeLCoMz." & vbCrLf & _
                    "Jika nama filenya diawali dengan 'Dokumen' dan nama folder diatasnya, maka anda boleh menghapusnya." & vbCrLf & _
                    "Anda yakin ingin menghapusnya ?", vbYesNo, "Virus ditemukan")
                    If Jawab = vbYes Then
                        Bunuh Path & NamaFile
                    End If
                End If
            
            End If
            
            ' *** Cari file berikutnya ***
            Cont = FindNextFile(hSearch, WFD)
            
        Wend
        
        Cont = FindClose(hSearch)
        
    End If
    
    ' Jika ada sub-direktori
    If JumlahDir > 0 Then
        ' Masuk ke Direktori secara Rekursif
        For i = 0 To JumlahDir - 1
            Cari_Virus = Cari_Virus + Cari_Virus(Path & BufNamaDir(i) & "\", FileYgDicari, HitungFile, HitungDir)
        Next i
    End If
    
End Function
