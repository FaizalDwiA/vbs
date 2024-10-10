# Creating Array

## Array Basic

```vb
Option Explicit
dim animals(2)

animals(0) = "horse"
animals(1) = "turtle"
animals(2) = "rabbit"

MsgBox animals(1)
```

```vb
names = Array("black", "syiber", "thunder")

MsgBox names(2)
```

## Join

```vb
names = Array("Black", "Syiber", "Thunder", "Dragon")

MsgBox join([names])
```

![1](../asset/img/44/1.webp)

## Pemisah Ketika Join

```vb
names = Array("Black", "Syiber", "Thunder", "Dragon")

MsgBox join(names, ",")
```

![2](../asset/img/44/2.webp)

```vb
names = Array("Black", "Syiber", "Thunder", "Dragon")

MsgBox join(names, vbLf)
```

![3](../asset/img/44/3.webp)

## Menampilkan Array 1 Per 1

```vb
names = Array("Black", "Syiber", "Thunder", "Dragon")

for i = LBound(names) to UBound(names)
    MsgBox names(i)
next
```

## LBound & UBound

> LBound : menampilkan angka array pertama
> UBound : menampilkan angka array terakhir

```vb
names = Array("Black", "Syiber", "Thunder", "Dragon")

MsgBox LBound(names)
MsgBox UBound(names)
```

> LBound : 0
> LBound : 3

## Filter

```vb
Option Explicit
dim names, name

names = Array("Black", "Syiber", "Thunder", "Dragon")

for Each name in Filter(names, "a")
    MsgBox name
Next
```

> akan menampilkan isi array yang didalamnya **terdapat** huruf / kata yang dicari

## False

```vb
Option Explicit
dim names, name

names = Array("Black", "Syiber", "Thunder", "Dragon")

for Each name in Filter(names, "a", False)
    MsgBox name
Next
```

> akan menampilkan isi array yang didalamnya **tidak terdapat** huruf / kata yang dicari

## Mengfilter

```vb
Option Explicit
dim names, name, list

names = Array("Black", "Syiber", "Thunder", "Dragon")

for Each name in Filter(names, "a", False)
    list = list & name & vbLf
Next

MsgBox list
```

![4](../asset/img/44/4.webp)

## Example
