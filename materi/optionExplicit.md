# Option Explicit

**adalah** sebuah penyatann dalam Visual Basic (VB) dan VBScript yang memaksa pengguna untuk mendeklarasikan semua variabel sebelum digunakan dalam program.

Ketika `Option Explicit` digunakan, compiler akan memberikan error jika ada variabel yang tidak dideklarasikan sebelum digunakan.

## Contoh

```vbs
Option Explicit

dim nama, umur
nama = "syiber"
umur = 21

MsgBox "Nama : " & nama & vbLf & "Umur : " & umur
```

## Jika tidak menggunakan

akan menyebabkan bug yang sulit dilacak
