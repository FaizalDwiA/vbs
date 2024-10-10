# GetObject

digunakan untuk **mendapatkan referensi** ke sebuah objek OLE yang telah diinstansiasi didalam sistem.

**Objek OLE (Object Linking and Embedding)** : sebuah objek yang dapat digunakan oleh aplikasi lain, seperti Microsoft Excel atau Microsoft Word, dan dapat dipanggil melalui kode VBS.

Fungsi `GetObject` dapat digunakan untuk membuka file / aplikasi OLE, seperti excel atau aplikasi word, dan memanipulasi objek yang terkandung didalamnya.

Fungsi ini memiliki 1 parameter wajib : `PathName` yang merupakan string yang menunjukkan jalur file / program yang ingin dibuka.

Selain itu, `GetObject` juga dapat menerima parameter opsional: `class` dan `ProgID` yang menentukan tipe objek yang ingin diambil.
