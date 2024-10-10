# Startup Program


## Program yang dijalankan 

Program yang dijalankan ketika kita baru mulai windows / startup

Melihat program yang dijalankan saat startup : 

1. Klik tombol "Start" atau "Mulai" di pojok kiri bawah layar.
2. Ketik "Task Manager" di kotak pencarian dan pilih Task Manager dari hasil pencarian yang muncul.
3. Pilih tab "Startup" di jendela Task Manager.
4. Di sini, Anda dapat melihat daftar program yang diizinkan untuk memulai bersamaan dengan Windows.
5. Jika suatu program dijalankan sebagai administrator saat startup, Anda akan melihat tanda "Enabled" di kolom "Status" dan "High" di kolom "Impact".
6. Anda juga dapat mengklik kanan pada nama program dan memilih "Properties" untuk memeriksa pengaturan eksekusi program.

---

## Membuat Program Startup dengan VBS

```vbs
option Explicit
dim objShell

set objShell = createObject("wscript.shell")

msgbox "selesai"

objShell.run "notepad.exe", 1, True
```

argument ke 2 > 1 > menunjukkan program harus dijalankan secara tersembunyi (hidden)

argument ke 3 > false > menunjukkan bahwa script VBS tidak boleh menunggu program selesai dijalankan.

Selanjutnya : 

1. Simpan file sebagai berkas dengan ekstensi ".vbs". Misalnya, Anda dapat menyimpannya sebagai "startup.vbs".
2. Pindahkan berkas .vbs yang sudah dibuat ke folder Startup di menu Start. Anda bisa akses folder tersebut dengan mengetikkan "shell:startup" pada search bar pada Windows 10.
3. Restart komputer Anda dan program yang dijalankan pada startup akan berjalan secara otomatis.

---

## Penjelasan lengkap mengenai argumen 2 dan 3

Argumen kedua menentukan opsi jendela program yang dijalankan, dengan nilai yang berbeda menunjukkan berbagai opsi tampilan jendela, sebagai berikut:

- 0 : menjalankan program dengan tampilan normal
- 1 : menjalankan program secara tersembunyi atau hidden
- 2 : menjalankan program secara maksimized (penuh layar)
- 3 : menjalankan program secara minimized (dalam taskbar)

argumen ketiga menunjukkan apakah skrip VBS harus menunggu hingga program yang dijalankan saat startup selesai dijalankan..

**True** : Jika Anda ingin skrip VBS menunggu program selesai dijalankan sebelum melanjutkan eksekusi.

**False** : jika Anda ingin skrip VBS melanjutkan eksekusi tanpa menunggu program selesai dijalankan.