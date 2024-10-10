# Mengunci Foder dengan Visual Basic

Anda dapat menggunakan skrip VBS (Visual Basic Script) untuk mengunci folder di sistem Windows. Berikut adalah contoh skrip VBS yang dapat Anda gunakan untuk mengunci folder:

```vb
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.BrowseForFolder(0, "Pilih folder yang ingin dikunci:", 0)
If objFolder Is Nothing Then
    WScript.Quit
End If
Set objFolderItem = objFolder.Self
strFolderName = objFolderItem.Path
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(strFolderName) Then
    Set objFolder = objFSO.GetFolder(strFolderName)
    objFolder.Attributes = 2
    WScript.Echo "Folder " & strFolderName & " telah dikunci."
Else
    WScript.Echo strFolderName & " tidak ditemukan."
End If
```

Skrip ini akan membuka jendela dialog pemilih folder, meminta Anda untuk memilih folder yang ingin dikunci, dan kemudian mengubah atribut folder menjadi "sembunyikan" sehingga folder tersebut tidak terlihat dalam eksplorator file. Untuk membuka kembali folder, Anda dapat menggunakan skrip yang sama, tetapi mengubah atribut folder menjadi normal.

```vb
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.BrowseForFolder(0, "Pilih folder yang ingin dibuka kunci:", 0)
If objFolder Is Nothing Then
    WScript.Quit
End If
Set objFolderItem = objFolder.Self
strFolderName = objFolderItem.Path
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(strFolderName) Then
    Set objFolder = objFSO.GetFolder(strFolderName)
    objFolder.Attributes = 0
    WScript.Echo "Folder " & strFolderName & " telah dibuka kunci."
Else
    WScript.Echo strFolderName & " tidak ditemukan."
End If
```

Anda dapat menyimpan kedua skrip ini dalam file teks dengan ekstensi .vbs dan menjalankannya dengan mengklik dua kali pada file tersebut. Pastikan bahwa Anda menjalankan skrip dengan hak akses administrator untuk memastikan bahwa Anda memiliki izin untuk mengubah atribut folder.

```vb
Set objShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strFolder = objShell.BrowseForFolder(0, "Pilih folder yang ingin dikunci:", 0).Self.Path
If objFSO.FolderExists(strFolder) Then
    Set objFolder = objFSO.GetFolder(strFolder)
    Set objFolderItem = objShell.NameSpace(objFolder.Path)
    If Not objFolderItem is Nothing Then 'tambahkan perintah kondisional
        strNewName = InputBox("Masukkan sandi untuk mengunci folder " & objFolder.Name & ":", "Kunci Folder")
        If strNewName <> "" Then
            objFolderItem.CopyHere objFolder.Path, &H80
            Set objFolderItem = objShell.NameSpace(objFolder.Path & ":Encryptable")
            objFolderItem.MoveHere objFolder.Path, &H100
            objFolderItem.MoveHere objFolder.Path & ".{21EC2020-3AEA-1069-A2DD-08002B30309D}", &H100
            Set objFile = objFSO.CreateTextFile(objFolder.Path & ".{21EC2020-3AEA-1069-A2DD-08002B30309D}\Locker.bat")
            objFile.WriteLine "@ECHO OFF"
            objFile.WriteLine "title Folder Private"
            objFile.WriteLine "if EXIST ""Control Panel.{21EC2020-3AEA-1069-A2DD-08002B30309D}"" goto UNLOCK"
            objFile.WriteLine "if NOT EXIST Private goto MDLOCKER"
            objFile.WriteLine ":CONFIRM"
            objFile.WriteLine "echo Apakah Anda yakin ingin mengunci folder? (Y/N)"
            objFile.WriteLine "set/p ""cho=>"""
            objFile.WriteLine "if %cho%==Y goto LOCK"
            objFile.WriteLine "if %cho%==y goto LOCK"
            objFile.WriteLine "if %cho%==n goto END"
            objFile.WriteLine "if %cho%==N goto END"
            objFile.WriteLine "echo Silakan masukkan Y atau N."
            objFile.WriteLine "goto CONFIRM"
            objFile.WriteLine ":LOCK"
            objFile.WriteLine "ren Private ""Control Panel.{21EC2020-3AEA-1069-A2DD-08002B30309D}"""
            objFile.WriteLine "attrib +h +s ""Control Panel.{21EC2020-3AEA-1069-A2DD-08002B30309D}"""
            objFile.WriteLine "echo Folder terkunci."
            objFile.WriteLine "goto END"
            objFile.WriteLine ":UNLOCK"
            objFile.WriteLine "echo Masukkan sandi untuk membuka kunci folder."
            objFile.WriteLine "set/p ""pass=>"""
            objFile.WriteLine "if NOT %pass%== " & strNewName & " goto FAIL"
            objFile.WriteLine "attrib -h -s ""Control Panel.{21EC2020-3AEA-1069-A2DD-08002B30309D}"""
            objFile.WriteLine "ren ""Control Panel.{21EC2020-3AEA-1069-A2DD-08002B30309D}"" Private"
            objFile.WriteLine "echo Folder berhasil dibuka kunci."
            objFile.WriteLine "goto END"
            objFile.WriteLine ":FAIL"
            objFile.WriteLine "echo Sandi salah."
            objFile.WriteLine "goto END"
            objFile.WriteLine ":MDLOCKER"
            objFile.WriteLine "md Private"
            objFile.WriteLine "echo Folder berhasil dibuat."
            objFile.WriteLine "goto END"
            objFile.WriteLine ":END"
            objFile.Close
            Set objFile = Nothing
            WScript.Echo "Folder " & objFolder.Name & " berhasil dikunci dengan sandi " & strNewName & "."
        Else
            WScript.Echo "Sandi tidak valid."
        End If
    Else
        MsgBox "Terjadi kesalahan saat mencoba mengunci folder.", vbCritical, "Kunci Folder"
    end if
Else
    WScript.Echo "Folder tidak ditemukan."
End If
```

## Berhasil

```vb
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.BrowseForFolder(0, "Pilih folder yang akan dikunci:", 0, 0)
If objFolder Is Nothing Then
    Wscript.Quit
End If
strFolder = objFolder.ParentFolder.ParseName(objFolder.Title).Path
Set objFolder = Nothing
Set objShell = Nothing

strPassword = InputBox("Masukkan kata sandi untuk mengunci folder:")
If strPassword = "" Then
    Wscript.Quit
End If

Set objFolder = CreateObject("Scripting.FileSystemObject").GetFolder(strFolder)
strFolderName = objFolder.Name 
Set objFolder = Nothing

Set objShell = CreateObject("WScript.Shell")
strCommand = "cmd /c ren " & Chr(34) & strFolder & Chr(34) & " " & Chr(34) & strFolderName & ".{ED7BA470-8E54-465E-825C-99712043E01C}" & Chr(34)
objShell.Run strCommand, 0, True

Set objShell = Nothing

Set objFolder = CreateObject("Scripting.FileSystemObject").GetFolder(strFolder & ".{ED7BA470-8E54-465E-825C-99712043E01C}")
strFolder = objFolder.Path
Set objFolder = Nothing

Set objShell = CreateObject("WScript.Shell")
strCommand = "cmd /c attrib +h " & Chr(34) & strFolder & Chr(34)
objShell.Run strCommand, 0, True

Set objShell = Nothing

strMessage = "Folder " & strFolderName & " telah dikunci dengan kata sandi " & strPassword & "."
MsgBox strMessage, vbInformation, "Folder Terkunci"
```

