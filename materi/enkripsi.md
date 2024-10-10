# Enkripsi

```vb
'Input teks yang akan dienkripsi
strInput = InputBox("Masukkan teks yang akan dienkripsi:")

'Konversi teks ke kode ASCII dan tambahkan nilai konstanta
For i = 1 To Len(strInput)
    intASCII = Asc(Mid(strInput, i, 1))
    intEncrypted = intASCII + 10 'contoh nilai konstanta: 10
    strEncrypted = strEncrypted & Chr(intEncrypted)
Next

'Tampilkan teks yang telah dienkripsi
MsgBox "Teks yang telah dienkripsi: " & strEncrypted
```

Berikut adalah penjelasan untuk setiap baris pada script tersebut:

```vb
'Input teks yang akan dienkripsi
strInput = InputBox("Masukkan teks yang akan dienkripsi:")
```

Pada baris pertama, kita meminta pengguna untuk memasukkan teks yang akan dienkripsi dengan menggunakan fungsi InputBox. Fungsi ini akan menampilkan kotak dialog yang meminta pengguna untuk memasukkan sebuah teks, lalu nilai teks tersebut akan disimpan dalam variabel strInput.

```vb
'Konversi teks ke kode ASCII dan tambahkan nilai konstanta
For i = 1 To Len(strInput)
    intASCII = Asc(Mid(strInput, i, 1))
    intEncrypted = intASCII + 10 'contoh nilai konstanta: 10
    strEncrypted = strEncrypted & Chr(intEncrypted)
Next
```

Pada baris kedua, kita menggunakan perulangan For untuk mengambil setiap karakter dalam teks yang telah dimasukkan oleh pengguna menggunakan fungsi Len (untuk mendapatkan panjang teks) dan Mid (untuk mengambil karakter ke-i dalam teks).

Fungsi Asc dan Mid digunakan pada baris intASCII = Asc(Mid(strInput, i, 1)) untuk mengonversi sebuah karakter dalam teks menjadi kode ASCII.

Fungsi Mid digunakan untuk mengambil satu karakter dalam teks. Argumen pertama dari fungsi ini adalah teks yang akan diambil karakternya, yaitu strInput. Argumen kedua adalah indeks karakter yang akan diambil, yaitu i. Argumen ketiga menunjukkan berapa karakter yang akan diambil dari indeks tersebut. Pada kasus ini, kita hanya mengambil satu karakter, sehingga argumen ketiga diisi dengan nilai 1.

Kemudian, fungsi Asc digunakan untuk mengonversi karakter tersebut menjadi kode ASCII. Fungsi ini mengembalikan nilai ASCII dari karakter yang diberikan sebagai argumen. Kode ASCII adalah representasi numerik dari sebuah karakter dalam komputer.

Dengan menggunakan fungsi Mid dan Asc, kita dapat mengonversi sebuah karakter dalam teks menjadi kode ASCII. Selanjutnya, kita dapat melakukan operasi matematika, seperti penambahan nilai konstanta, pada kode ASCII tersebut.

Kemudian, pada baris ketiga, kita menggunakan fungsi Asc untuk mengonversi setiap karakter ke kode ASCII, yang kemudian disimpan dalam variabel intASCII. Setelah itu, kita menambahkan nilai konstanta tertentu (dalam contoh ini, nilai konstanta adalah 10) pada setiap kode ASCII menggunakan operator penjumlahan, dan menyimpan hasilnya dalam variabel intEncrypted.

Pada baris keempat, kita mengonversi setiap kode ASCII yang telah dienkripsi kembali menjadi karakter menggunakan fungsi Chr, dan menyimpan hasilnya dalam variabel strEncrypted. Kita juga menggunakan operator penambahan (&) untuk menggabungkan setiap karakter yang telah dienkripsi.


```vb
'Tampilkan teks yang telah dienkripsi
MsgBox "Teks yang telah dienkripsi: " & strEncrypted
```

Pada baris terakhir, kita menampilkan teks yang telah dienkripsi menggunakan fungsi MsgBox, yang menampilkan kotak dialog dengan pesan teks dan tombol OK. Pada fungsi ini, kita menggunakan operator penambahan (&) untuk menggabungkan teks yang menginformasikan bahwa teks yang telah dienkripsi, beserta dengan teks yang telah dienkripsi tersebut.

## Decripsi

```vb
'Input teks yang telah dienkripsi
strEncrypted = InputBox("Masukkan teks yang telah dienkripsi:")

'Konversi teks yang telah dienkripsi ke kode ASCII dan kurangi nilai konstanta
For i = 1 To Len(strEncrypted)
    intEncrypted = Asc(Mid(strEncrypted, i, 1))
    intDecrypted = intEncrypted - 10 'contoh nilai konstanta: 10
    strDecrypted = strDecrypted & Chr(intDecrypted)
Next

'Tampilkan teks yang telah didekripsi
MsgBox "Teks yang telah didekripsi: " & strDecrypted

```

## contoh program enkripsi isi file

```vb
Set fso = CreateObject("Scripting.FileSystemObject")

set cmd = CreateObject("wscript.shell")

Set fso = CreateObject("Scripting.FileSystemObject")
cd = Replace(WScript.ScriptFullName, WScript.ScriptName, "coba.txt")

set file = fso.OpenTextFile(cd, 1)
text = file.ReadAll
file.Close

coba = InputBox("pilih opsi : " & vblf & "1. Enkripsi" & vblf & "2. Dekripsi")
MsgBox coba

For i = 1 To Len(text)
    intASCII = Asc(Mid(text, i, 1))
    if coba = "1" then
        if i = 1 then
            intEncrypted = intASCII + 10 'contoh nilai konstanta: 10
        elseif i = 2 then
            intEncrypted = intASCII + 150 'contoh nilai konstanta: 10
        elseif i = 3 then
            intEncrypted = intASCII + 55 'contoh nilai konstanta: 10
        elseif i = 8 then
            intEncrypted = intASCII + 99 'contoh nilai konstanta: 10
        Else
            intEncrypted = intASCII + 100 'contoh nilai konstanta: 10
        end if
    else
        if i = 1 then
            intEncrypted = intASCII - 10 'contoh nilai konstanta: 10
        elseif i = 2 then
            intEncrypted = intASCII - 150 'contoh nilai konstanta: 10
        elseif i = 3 then
            intEncrypted = intASCII - 55 'contoh nilai konstanta: 10
        elseif i = 8 then
            intEncrypted = intASCII - 99 'contoh nilai konstanta: 10
        Else
            intEncrypted = intASCII - 100 'contoh nilai konstanta: 10
        end if
    end if
    strEncrypted = strEncrypted & Chr(intEncrypted)
Next

set file = fso.OpenTextFile(cd, 2)

file.Write strEncrypted
file.Close

WScript.Echo "Complete."
```

## Contoh lain

```vb
Set fso = CreateObject("Scripting.FileSystemObject")

set cmd = CreateObject("wscript.shell")

Set fso = CreateObject("Scripting.FileSystemObject")
cd = Replace(WScript.ScriptFullName, WScript.ScriptName, "img.jpg")

set file = fso.OpenTextFile(cd, 1)
text = file.ReadAll
file.Close

coba = InputBox("pilih opsi : " & vblf & "1. Enkripsi" & vblf & "2. Dekripsi")
For i = 1 To Len(text)
    intASCII = Asc(Mid(text, i, 1))

    if coba = "1" then
        if i = 1 then
            intEncrypted = intASCII + 10 'contoh nilai konstanta: 10
        elseif i = 2 then
            intEncrypted = intASCII + 150 'contoh nilai konstanta: 10
        elseif i = 3 then
            intEncrypted = intASCII + 55 'contoh nilai konstanta: 10
        elseif i = 8 then
            intEncrypted = intASCII + 99 'contoh nilai konstanta: 10
        Else
            intEncrypted = intASCII + 100 'contoh nilai konstanta: 10
        end if
    else
        if i = 1 then
            intEncrypted = intASCII - 10 'contoh nilai konstanta: 10
        elseif i = 2 then
            intEncrypted = intASCII - 150 'contoh nilai konstanta: 10
        elseif i = 3 then
            intEncrypted = intASCII - 55 'contoh nilai konstanta: 10
        elseif i = 8 then
            intEncrypted = intASCII - 99 'contoh nilai konstanta: 10
        Else
            intEncrypted = intASCII - 100 'contoh nilai konstanta: 10
        end if
    end if

    if intEncrypted >= 255 then
        Do until intEncrypted <= 255
            intEncrypted = intEncrypted - 255
        loop
    end if

    if i mod len(text)/10 = 0 then
        MsgBox i
    end if
    
    strEncrypted = strEncrypted & Chr(intEncrypted)
Next

MsgBox "ok"

set file = fso.OpenTextFile(cd, 2)

file.Write strEncrypted
file.Close

WScript.Echo "Complete."
```

```vb
Set fso = CreateObject("Scripting.FileSystemObject")

set cmd = CreateObject("wscript.shell")

Set fso = CreateObject("Scripting.FileSystemObject")
cd = Replace(WScript.ScriptFullName, WScript.ScriptName, "img.jpg")

set file = fso.OpenTextFile(cd, 1)
text = file.ReadAll
file.Close

' ganti = Replace(text, "ÿ", f)
MsgBox text
ganti = Replace(text, "f", "ÿ")
MsgBox ganti

MsgBox "ok"

set file = fso.OpenTextFile(cd, 2)

file.Write ganti
file.Close

WScript.Echo "Complete."
```