# Kotak Dialog / Message Box

## Basic

```vb
set objShell = CreateObject("wscript.shell")
intButton = objShell.Popup("Ini adalah pesan dalam kotak dialog", 2, "Judul Kotak Dialog", vbInformation + VbMsgBoxSetForeground + vbSystemModal)
```
