VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer, j As Integer

Private Sub Command1_Click()
  Dim Counters(14) As Integer
  Dim Sums(20) As Double
  'Deklarasi pertama membuat sebuah array dengan 15 element,
  'dengan nomor indeks dimulai dari 0 sampai 14. Yang kedua
  'membuat sebuah array dengan 21 element, dengan nomor
  'indeks dimulai dari 0 sampai 20.
  'Default indeks terendah adalah 0 (nol).
  'Berikut salah satu cara untuk mengisi
  'dan mengakses array tsb:
  For i = 0 To 14
    Counters(i) = i
    MsgBox Counters(i)
  Next i
  
  For j = 0 To 20
    Sums(j) = j * 0.2
    MsgBox Sums(j)
  Next j
  
  'Untuk membuat elemen array yang terendah sesuai dengan keinginan, kita harus
  'mengekspresikannya secara eksplisit (seperti sebuah tipe data Long) menggunakan
  'kata kunci "To". Perhatikan contoh di bawah ini:
  Dim Counters1(1 To 15) As Integer
  Dim Sums1(100 To 120) As String
  'Deklarasi yang pertama, jumlah indeks dari Counters1 mulai dari 1 sampai 15,
  'dan jumlah indeks dari Sums1 mulai dari 100 sampai 120.
  'Berikut cara mengisi dan mengakses array tsb:
  For i = 1 To 15
     Counters1(i) = i
     MsgBox Counters1(i)
  Next i
  
  For j = 100 To 120
     Sums1(j) = "String ke-" & j
     MsgBox Sums1(j)
  Next j

End Sub
