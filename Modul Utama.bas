Attribute VB_Name = "ModulUtama"
Option Explicit

' *** Modul Utama ***

Public NamaUser As String
Public WinDir As String
Public UserProfile As String
Public FolderVirus As String
Public SeluruhInfo As String

' *** Fungsi Umum ***

Public Function Nama_Aplikasi() As String

    Nama_Aplikasi = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision

End Function

Function FileExist(strPath As String) As Integer

    Dim lngRetVal As Long
    
    On Error Resume Next
        
    'Set atribut file agar dapat dideteksi dengan FileExist()
    SetAttr strPath, GetAttr(strPath) And Not vbHidden
    SetAttr strPath, GetAttr(strPath) And Not vbReadOnly
    SetAttr strPath, GetAttr(strPath) And Not vbSystem
    
    lngRetVal = Len(Dir$(strPath))
    
    If Err Or lngRetVal = 0 Then
        FileExist = False
    Else
        FileExist = True
    End If
    
End Function

Public Sub WriteToFile(ByVal strFilename As String, ByVal strFileContents As String)

    Dim lngFileHandle As Long
    
    On Error Resume Next
    
    lngFileHandle = FreeFile
    
    Open strFilename For Binary As #lngFileHandle
    Put #lngFileHandle, , strFileContents
    Close #lngFileHandle
    
End Sub

Public Sub Bunuh(ByVal Virusnya As String)

Dim Bersih As Boolean

    Bersih = False

    Informasi "  » Menghapus " & Virusnya
    
    On Error Resume Next
    
    If FileExist(Virusnya) Then
     
        If InStr(Virusnya, "*") <> 0 Then
            Kill Virusnya
            Bersih = True
        End If
       
        If Err.Number <> 0 Then
            If Err.Number = 5 And Err.Number <> 52 Then
                Kill Virusnya
                Bersih = True
            Else
                Informasi "     » " & Err.Number & " : " & Err.Description
                Informasi "     » File tidak dapat dihapus! Coba cek manual file tsb!"
            End If
        Else
            Kill Virusnya
            Bersih = True
        End If
    Else
        Informasi "     » File tidak ada!"
    End If
    
    If Bersih Then
        Informasi "     » Virus telah dihapus!"
    End If
        
End Sub

Public Sub HapusReg(SubKey As String, Entry As String)

    Informasi "      » Menghapus Entry Registry " & SubKey & "\" & Entry
    DeleteValue SubKey, Entry
    
End Sub

Public Sub UbahRegDWORD(SubKey As String, Entry As String, Nilai As Long)

    'Informasi "      » Mengubah Entry Registry " & SubKey & "\" & Entry
    SetDWORDValue SubKey, Entry, Nilai
    
End Sub

Public Sub UbahRegString(SubKey As String, Entry As String, Nilai As String)

    'Informasi "      » Mengubah Entry Registry " & Subkey & "\" & Entry
    SetStringValue SubKey, Entry, Nilai
    
End Sub

Public Sub UbahRegStringEx(SubKey As String, Entry As String, Nilai As String)

    'Informasi "      » Mengubah Entry Registry " & Subkey & "\" & Entry
    SetStringValueEx SubKey, Entry, Nilai
    
End Sub

Public Sub Informasi(ByVal Infonya As String)
    
    SeluruhInfo = SeluruhInfo & Infonya & vbCrLf
    FDestroyer.Info.Text = SeluruhInfo
    FDestroyer.Info.SelStart = Len(SeluruhInfo)
    FDestroyer.Info.Refresh
    WriteToFile "FeeLCoMz.Log", SeluruhInfo
    
End Sub

Public Sub Hajar_Virus()

    Dim Status As Boolean
    Dim JmlVirMem As Integer

    NamaUser = Environ("UserName")
    WinDir = Environ("Windir")
    UserProfile = Environ("UserProfile")
    FolderVirus = UserProfile & AppData

    Status = False
    Informasi ""
    SeluruhInfo = ""
    JmlVirMem = 0
    
    ' *** Mulai ***
    ' *** Penghapusan Virus dari Memory ***
    Informasi "*** " & Now & " ***"
    Informasi ""
    Informasi "Pencarian dan Penghapusan Virus di Memory..."
    
    Do While FindWindow(vbNullString, NamaProsesVirus) <> 0
        JmlVirMem = JmlVirMem + 1
        WindowHandle FindWindow(vbNullString, NamaProsesVirus), 0
        Status = True
    Loop
    
    If Status = True Then
        Informasi "  » " & JmlVirMem & " Virus FeeLCoMz ditemukan di memory dan telah dibersihkan (Get Out, Bro!)"
        MsgBox JmlVirMem & " Virus FeeLCoMz ditemukan di memory dan telah dibersihkan (Get Out, Bro!)", vbInformation
    'Else
    '    Informasi "  » Virus FeeLCoMz tidak ada di memory"
    '    MsgBox "Virus FeeLCoMz tidak ada di memory", vbInformation
    End If
    
    ' *** Perbaikan Registry yang dimodifikasi oleh Virus ***
    
    Informasi ""
    Informasi "Perbaikan Registry..."
    Informasi "  » Mengembalikan menu Folder Options pada Explorer..."
    UbahRegDWORD ExPol, "NoFolderOptions", 0
    Informasi "  » Menyembuhkan opsi Hidden Files & Folders pada Folder Options..."
    UbahRegDWORD ExFolderOp & "\HideFileExt", "CheckedValue", 1
    UbahRegDWORD ExFolderOp & "\HideFileExt", "UncheckedValue", 0
    UbahRegDWORD ExFolderOp & "\HideFileExt", "DefaultValue", 1
    UbahRegDWORD ExFolderOp & "\Hidden\SHOWALL", "CheckedValue", 1
    UbahRegDWORD ExFolderOp & "\SuperHidden", "CheckedValue", 0
    UbahRegDWORD ExFolderOp & "\SuperHidden", "UncheckedValue", 1
    UbahRegDWORD ExFolderOp & "\SuperHidden", "DefaultValue", 0
    Informasi "  » Mengaktifkan kembali Registry Editor..."
    UbahRegDWORD SysPol, "DisableRegistryTools", 0
    Informasi "  » Menyembuhkan Shell Open Command untuk Notepad..."
    UbahRegStringEx Notepad, "", "%SystemRoot%\system32\NOTEPAD.EXE %1"
    Informasi "  » Menyembuhkan Shell Windows..."
    UbahRegString WinLogon, "Shell", "Explorer.exe"
    Informasi "  » Menghapus Shell Alternatif Virus pada Safe Mode... "
    HapusReg CtrSet1, "AlternateShell"
    HapusReg CtrSet2, "AlternateShell"
    Informasi "  » Menghapus Entry Startup Virus ..."
    HapusReg StartupRun, "Microsoft Local Security Authority Subsystem Service"
    HapusReg StartupRun, "Winzip Quick Pick"
    HapusReg StartupRun, "Generic Host Process for Win32 Services"
    HapusReg StartupRun, "Hardware Monitor"
    HapusReg StartupRun, "Adobe Gamma Loader"
    
    ' *** Pembunuhan Virus ***
    Informasi ""
    Informasi "Pembersihan Induk Virus..."
    Bunuh WinDir & "\system32\drivers\Etc\Host.com"
    Bunuh WinDir & "\system32\drivers\Etc\Host.exe"
    Bunuh WinDir & "\system32\drivers\Etc\svchost.com"
    Bunuh WinDir & "\system32\Winzip.exe"
    Bunuh WinDir & "\system\lsass.exe"
    Bunuh WinDir & "\system\svchost.com"
    Bunuh WinDir & "\system\svchost.exe"
    Bunuh UserProfile & "\Start Menu\Programs\Startup\Hardware Monitor.exe"
    Bunuh UserProfile & "\Start Menu\Programs\Startup\Adobe Gamma Loader.Com"
    Bunuh UserProfile & "\Notepad.exe"
    
    ' *** Penghapusan Virus Alternatif ***
    Bunuh UserProfile & "\Start Menu\Programs\Startup\ReadMe Startup.Exe"
    Bunuh UserProfile & "\Start Menu\Programs\Startup\ReadMe For Startup.Exe"
    Bunuh UserProfile & "\Start Menu\Programs\Startup\Apa Itu Startup.Exe"
    
    Informasi ""
    Informasi "*** Selesai ***"
    Informasi ""
    Informasi "*** " & Now & " ***"
    Informasi ""
    
    MsgBox "Pembersihan Induk Virus FeeLCoMz telah selesai." & vbCrLf & _
            "Silahkan hapus folder yang telah terinfeksi secara manual.", vbInformation
    
    ' *** Selesai ***
    
End Sub
