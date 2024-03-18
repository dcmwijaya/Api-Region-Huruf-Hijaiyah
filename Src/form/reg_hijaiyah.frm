VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Region_Hijaiyah"
   ClientHeight    =   9450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17985
   DrawWidth       =   3
   LinkTopic       =   "Form2"
   ScaleHeight     =   9450
   ScaleWidth      =   17985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'=================================================================================================='
'Parent'
'=================================================================================================='
Private Sub Form_Load()
a1 = CreateRoundRectRgn(900, 1010, 0, 0, 0, 0)
CombineRgn a1, a1, a1, 1

'=================================================================================================='
'Garis Potong Vertikal'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Enter") = CER'
lv1 = CreateRoundRectRgn(150, 1010, 160, 0, 0, 0)
lv2 = CreateRoundRectRgn(300, 1010, 310, 0, 0, 0)
lv3 = CreateRoundRectRgn(450, 1010, 460, 0, 0, 0)
lv4 = CreateRoundRectRgn(600, 1010, 610, 0, 0, 0)
lv5 = CreateRoundRectRgn(750, 1010, 760, 0, 0, 0)
lv6 = CreateRoundRectRgn(900, 1010, 910, 0, 0, 0)
lv7 = CreateRoundRectRgn(1050, 1010, 1060, 0, 0, 0)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, lv1, 4
CombineRgn a1, a1, lv2, 4
CombineRgn a1, a1, lv3, 4
CombineRgn a1, a1, lv4, 4
CombineRgn a1, a1, lv5, 4
CombineRgn a1, a1, lv6, 4
CombineRgn a1, a1, lv7, 4

'=================================================================================================='
'Garis Potong Horizontal'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
lz1 = CreateRoundRectRgn(0, 120, 900, 130, 0, 0)
lz2 = CreateRoundRectRgn(0, 250, 900, 260, 0, 0)
lz3 = CreateRoundRectRgn(0, 380, 900, 390, 0, 0)
lz4 = CreateRoundRectRgn(0, 510, 900, 520, 0, 0)
lz5 = CreateRoundRectRgn(0, 640, 900, 650, 0, 0)
lz6 = CreateRoundRectRgn(0, 770, 900, 780, 0, 0)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, lz1, 4
CombineRgn a1, a1, lz2, 4
CombineRgn a1, a1, lz3, 4
CombineRgn a1, a1, lz4, 4
CombineRgn a1, a1, lz5, 4
CombineRgn a1, a1, lz6, 4

'=================================================================================================='
'Alif'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
a2 = CreateRoundRectRgn(825, 15, 835, 100, 0, 0)
a3 = CreateEllipticRgn(800, 10, 830, 110)
a4 = CreateEllipticRgn(824, 80, 860, 120)
a5 = CreateEllipticRgn(800, 5, 832, 35)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, a2, 4
CombineRgn a1, a1, a3, 2
CombineRgn a1, a1, a4, 2
CombineRgn a1, a1, a5, 2

'=================================================================================================='
'Ba'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
b1 = CreateEllipticRgn(623, 40, 733, 85)
b2 = CreateEllipticRgn(623, 30, 733, 75)
b3 = CreateEllipticRgn(670, 92, 685, 105)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, b1, 4
CombineRgn a1, a1, b2, 2
CombineRgn a1, a1, b3, 4

'=================================================================================================='
'Ta'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
t1 = CreateEllipticRgn(483, 40, 583, 85)
t2 = CreateEllipticRgn(483, 30, 583, 75)
t3 = CreateEllipticRgn(510, 35, 525, 48)
t4 = CreateEllipticRgn(545, 35, 560, 48)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, t1, 4
CombineRgn a1, a1, t2, 2
CombineRgn a1, a1, t3, 4
CombineRgn a1, a1, t4, 4

'=================================================================================================='
'Tsa'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
ts1 = CreateEllipticRgn(332, 40, 432, 85)
ts2 = CreateEllipticRgn(332, 30, 432, 75)
ts3 = CreateEllipticRgn(360, 35, 375, 48)
ts4 = CreateEllipticRgn(395, 35, 410, 48)
ts5 = CreateEllipticRgn(375, 15, 390, 28)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, ts1, 4
CombineRgn a1, a1, ts2, 2
CombineRgn a1, a1, ts3, 4
CombineRgn a1, a1, ts4, 4
CombineRgn a1, a1, ts5, 4

'=================================================================================================='
'Jim'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
j1 = CreateRoundRectRgn(180, 25, 270, 35, 0, 0)
j2 = CreateEllipticRgn(175, 4, 195, 35)
j3 = CreateEllipticRgn(190, 30, 290, 100)
j4 = CreateEllipticRgn(220, 30, 300, 95)
j5 = CreateEllipticRgn(242, 50, 262, 70)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, j1, 4
CombineRgn a1, a1, j2, 2
CombineRgn a1, a1, j3, 4
CombineRgn a1, a1, j4, 2
CombineRgn a1, a1, j5, 4

'=================================================================================================='
'Ha'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
h1 = CreateRoundRectRgn(25, 25, 120, 35, 0, 0)
h2 = CreateEllipticRgn(15, 4, 40, 35)
h3 = CreateEllipticRgn(30, 30, 140, 100)
h4 = CreateEllipticRgn(60, 30, 144, 97)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, h1, 4
CombineRgn a1, a1, h2, 2
CombineRgn a1, a1, h3, 4
CombineRgn a1, a1, h4, 2

'=================================================================================================='
'Kho'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
kh1 = CreateRoundRectRgn(790, 165, 870, 175, 0, 0)
kh2 = CreateEllipticRgn(760, 150, 800, 178)
kh3 = CreateEllipticRgn(800, 170, 880, 240)
kh4 = CreateEllipticRgn(820, 170, 900, 230)
kh5 = CreateEllipticRgn(832, 140, 850, 155)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, kh1, 4
CombineRgn a1, a1, kh2, 2
CombineRgn a1, a1, kh3, 4
CombineRgn a1, a1, kh4, 2
CombineRgn a1, a1, kh5, 4

'=================================================================================================='
'Dal'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
d1 = CreateEllipticRgn(650, 170, 730, 220)
d2 = CreateEllipticRgn(650, 180, 720, 200)
d3 = CreateEllipticRgn(625, 180, 650, 220)
d4 = CreateEllipticRgn(640, 145, 690, 190)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, d1, 4
CombineRgn a1, a1, d2, 2
CombineRgn a1, a1, d3, 2
CombineRgn a1, a1, d4, 2

'=================================================================================================='
'Dzal'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
dz1 = CreateEllipticRgn(500, 170, 580, 220)
dz2 = CreateEllipticRgn(500, 180, 570, 200)
dz3 = CreateEllipticRgn(500, 180, 520, 190)
dz4 = CreateEllipticRgn(500, 145, 530, 190)
dz5 = CreateEllipticRgn(530, 150, 550, 165)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, dz1, 4
CombineRgn a1, a1, dz2, 2
CombineRgn a1, a1, dz3, 2
CombineRgn a1, a1, dz4, 2
CombineRgn a1, a1, dz5, 4

'=================================================================================================='
'Ro'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
r1 = CreateEllipticRgn(340, 150, 420, 220)
r2 = CreateEllipticRgn(310, 160, 410, 215)
r3 = CreateEllipticRgn(340, 140, 405, 190)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, r1, 4
CombineRgn a1, a1, r2, 2
CombineRgn a1, a1, r3, 2

'=================================================================================================='
'Za'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
z1 = CreateEllipticRgn(190, 150, 270, 220)
z2 = CreateEllipticRgn(170, 160, 260, 215)
z3 = CreateEllipticRgn(160, 140, 255, 190)
z4 = CreateEllipticRgn(230, 150, 245, 170)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, z1, 4
CombineRgn a1, a1, z2, 2
CombineRgn a1, a1, z3, 2
CombineRgn a1, a1, z4, 4

'=================================================================================================='
'Sin'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
s1 = CreateEllipticRgn(90, 150, 130, 200)
s2 = CreateEllipticRgn(90, 130, 130, 190)
s3 = CreateEllipticRgn(60, 160, 100, 200)
s4 = CreateEllipticRgn(70, 130, 100, 190)
s5 = CreateEllipticRgn(10, 160, 85, 230)
s6 = CreateEllipticRgn(10, 140, 80, 220)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, s1, 4
CombineRgn a1, a1, s2, 2
CombineRgn a1, a1, s3, 4
CombineRgn a1, a1, s4, 2
CombineRgn a1, a1, s5, 4
CombineRgn a1, a1, s6, 2

'=================================================================================================='
'Syin'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
sy1 = CreateEllipticRgn(850, 280, 895, 320)
sy2 = CreateEllipticRgn(855, 270, 885, 310)
sy3 = CreateEllipticRgn(830, 285, 860, 325)
sy4 = CreateEllipticRgn(825, 270, 861, 310)
sy5 = CreateEllipticRgn(770, 290, 845, 360)
sy6 = CreateEllipticRgn(775, 280, 830, 350)
sy7 = CreateEllipticRgn(865, 285, 880, 300)
sy8 = CreateEllipticRgn(835, 285, 850, 300)
sy9 = CreateEllipticRgn(850, 270, 865, 285)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, sy1, 4
CombineRgn a1, a1, sy2, 2
CombineRgn a1, a1, sy3, 4
CombineRgn a1, a1, sy4, 2
CombineRgn a1, a1, sy5, 4
CombineRgn a1, a1, sy6, 2
CombineRgn a1, a1, sy7, 4
CombineRgn a1, a1, sy8, 4
CombineRgn a1, a1, sy9, 4

'=================================================================================================='
'Shod'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
sh1 = CreateEllipticRgn(625, 290, 694, 360)
sh2 = CreateEllipticRgn(630, 280, 685, 355)
sh3 = CreateEllipticRgn(680, 285, 745, 325)
sh4 = CreateEllipticRgn(675, 260, 720, 310)
sh5 = CreateEllipticRgn(660, 260, 740, 298)
sh6 = CreateEllipticRgn(710, 305, 735, 315)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, sh1, 4
CombineRgn a1, a1, sh2, 2
CombineRgn a1, a1, sh3, 4
CombineRgn a1, a1, sh4, 2
CombineRgn a1, a1, sh5, 2
CombineRgn a1, a1, sh6, 2

'=================================================================================================='
'Dlod'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
dl1 = CreateEllipticRgn(475, 290, 545, 360)
dl2 = CreateEllipticRgn(481, 280, 536, 355)
dl3 = CreateEllipticRgn(531, 285, 596, 325)
dl4 = CreateEllipticRgn(526, 260, 571, 310)
dl5 = CreateEllipticRgn(511, 260, 591, 298)
dl6 = CreateEllipticRgn(560, 305, 585, 315)
dl7 = CreateEllipticRgn(540, 280, 560, 295)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, dl1, 4
CombineRgn a1, a1, dl2, 2
CombineRgn a1, a1, dl3, 4
CombineRgn a1, a1, dl4, 2
CombineRgn a1, a1, dl5, 2
CombineRgn a1, a1, dl6, 2
CombineRgn a1, a1, dl7, 4

'=================================================================================================='
'Tho'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
th1 = CreateRoundRectRgn(370, 278, 390, 360, 0, 0)
th2 = CreateEllipticRgn(330, 330, 435, 370)
th3 = CreateEllipticRgn(380, 280, 420, 340)
th4 = CreateEllipticRgn(325, 270, 376, 355)
th5 = CreateEllipticRgn(325, 265, 382, 300)
th6 = CreateEllipticRgn(325, 260, 380, 310)
th7 = CreateEllipticRgn(385, 345, 420, 355)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, th1, 4
CombineRgn a1, a1, th2, 4
CombineRgn a1, a1, th3, 2
CombineRgn a1, a1, th4, 2
CombineRgn a1, a1, th5, 2
CombineRgn a1, a1, th6, 2
CombineRgn a1, a1, th7, 2

'=================================================================================================='
'Dzo'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
dz1 = CreateRoundRectRgn(220, 278, 240, 360, 0, 0)
dz2 = CreateEllipticRgn(180, 330, 285, 370)
dz3 = CreateEllipticRgn(230, 280, 270, 340)
dz4 = CreateEllipticRgn(175, 270, 226, 355)
dz5 = CreateEllipticRgn(175, 265, 232, 300)
dz6 = CreateEllipticRgn(175, 260, 230, 310)
dz7 = CreateEllipticRgn(235, 345, 270, 355)
dz8 = CreateEllipticRgn(255, 310, 270, 325)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, dz1, 4
CombineRgn a1, a1, dz2, 4
CombineRgn a1, a1, dz3, 2
CombineRgn a1, a1, dz4, 2
CombineRgn a1, a1, dz5, 2
CombineRgn a1, a1, dz6, 2
CombineRgn a1, a1, dz7, 2
CombineRgn a1, a1, dz8, 4

'=================================================================================================='
'Ain'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
an1 = CreateEllipticRgn(45, 310, 115, 360)
an2 = CreateEllipticRgn(55, 310, 120, 355)
an3 = CreateEllipticRgn(45, 285, 80, 315)
an4 = CreateEllipticRgn(55, 290, 90, 312)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, an1, 4
CombineRgn a1, a1, an2, 2
CombineRgn a1, a1, an3, 4
CombineRgn a1, a1, an4, 2

'=================================================================================================='
'Ghoin'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
gh1 = CreateEllipticRgn(800, 445, 870, 500)
gh2 = CreateEllipticRgn(820, 449, 880, 490)
gh3 = CreateEllipticRgn(800, 420, 840, 457)
gh4 = CreateEllipticRgn(810, 428, 850, 446)
gh5 = CreateEllipticRgn(810, 397, 825, 415)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, gh1, 4
CombineRgn a1, a1, gh2, 2
CombineRgn a1, a1, gh3, 4
CombineRgn a1, a1, gh4, 2
CombineRgn a1, a1, gh5, 4

'=================================================================================================='
'Fa'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
f1 = CreateEllipticRgn(625, 460, 727, 500)
f2 = CreateEllipticRgn(630, 440, 713, 492)
f3 = CreateEllipticRgn(690, 450, 732, 490)
f4 = CreateEllipticRgn(635, 470, 705, 492)
f5 = CreateEllipticRgn(695, 460, 720, 470)
f6 = CreateEllipticRgn(700, 425, 715, 440)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, f1, 4
CombineRgn a1, a1, f2, 2
CombineRgn a1, a1, f3, 4
CombineRgn a1, a1, f4, 2
CombineRgn a1, a1, f5, 2
CombineRgn a1, a1, f6, 4

'=================================================================================================='
'Qof'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
q1 = CreateEllipticRgn(490, 460, 575, 500)
q2 = CreateEllipticRgn(495, 440, 561, 492)
q3 = CreateEllipticRgn(535, 450, 577, 490)
q4 = CreateEllipticRgn(495, 470, 550, 492)
q5 = CreateEllipticRgn(540, 460, 565, 470)
q6 = CreateEllipticRgn(535, 425, 550, 440)
q7 = CreateEllipticRgn(555, 425, 570, 440)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, q1, 4
CombineRgn a1, a1, q2, 2
CombineRgn a1, a1, q3, 4
CombineRgn a1, a1, q4, 2
CombineRgn a1, a1, q5, 2
CombineRgn a1, a1, q6, 4
CombineRgn a1, a1, q7, 4

'=================================================================================================='
'Kaf'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
kf1 = CreateEllipticRgn(345, 479, 420, 505)
kf2 = CreateEllipticRgn(395, 400, 435, 500)
kf3 = CreateEllipticRgn(310, 400, 425, 489)
kf4 = CreateEllipticRgn(360, 460, 390, 485)
kf5 = CreateEllipticRgn(350, 460, 380, 480)
kf6 = CreateEllipticRgn(365, 440, 405, 465)
kf7 = CreateEllipticRgn(375, 440, 415, 470)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, kf1, 4
CombineRgn a1, a1, kf2, 4
CombineRgn a1, a1, kf3, 2
CombineRgn a1, a1, kf4, 4
CombineRgn a1, a1, kf5, 2
CombineRgn a1, a1, kf6, 4
CombineRgn a1, a1, kf7, 2

'=================================================================================================='
'Lam'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
lm1 = CreateEllipticRgn(205, 450, 275, 500)
lm2 = CreateEllipticRgn(215, 400, 275, 490)
lm3 = CreateEllipticRgn(262, 400, 278, 485)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, lm1, 4
CombineRgn a1, a1, lm2, 2
CombineRgn a1, a1, lm3, 4

'=================================================================================================='
'Mim'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
mm1 = CreateEllipticRgn(60, 410, 120, 450)
mm2 = CreateEllipticRgn(60, 425, 75, 500)
mm3 = CreateEllipticRgn(65, 435, 105, 500)
mm4 = CreateEllipticRgn(75, 445, 120, 500)
mm5 = CreateEllipticRgn(75, 415, 100, 425)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, mm1, 4
CombineRgn a1, a1, mm2, 4
CombineRgn a1, a1, mm3, 2
CombineRgn a1, a1, mm4, 2
CombineRgn a1, a1, mm5, 2


'=================================================================================================='
'Nun'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
n1 = CreateEllipticRgn(780, 550, 880, 620)
n2 = CreateEllipticRgn(780, 540, 878, 610)
n3 = CreateEllipticRgn(825, 560, 840, 580)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, n1, 4
CombineRgn a1, a1, n2, 2
CombineRgn a1, a1, n3, 4

'=================================================================================================='
'Wawu'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
w1 = CreateEllipticRgn(640, 550, 727, 615)
w2 = CreateEllipticRgn(630, 540, 713, 610)
w3 = CreateEllipticRgn(680, 550, 725, 590)
w4 = CreateEllipticRgn(685, 560, 715, 575)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, w1, 4
CombineRgn a1, a1, w2, 2
CombineRgn a1, a1, w3, 4
CombineRgn a1, a1, w4, 2

'=================================================================================================='
'Hha'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
hh1 = CreateEllipticRgn(510, 550, 560, 600)
hh2 = CreateEllipticRgn(490, 540, 520, 580)
hh3 = CreateEllipticRgn(500, 580, 550, 640)
hh4 = CreateEllipticRgn(520, 590, 560, 640)
hh5 = CreateEllipticRgn(542, 565, 552, 575)
hh6 = CreateEllipticRgn(527, 565, 537, 575)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, hh1, 4
CombineRgn a1, a1, hh2, 2
CombineRgn a1, a1, hh3, 2
CombineRgn a1, a1, hh4, 2
CombineRgn a1, a1, hh5, 2
CombineRgn a1, a1, hh6, 2

'=================================================================================================='
'Lam Alif'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
la1 = CreateRoundRectRgn(410, 540, 425, 590, 0, 0)
la2 = CreateRoundRectRgn(380, 575, 425, 590, 0, 0)
la3 = CreateEllipticRgn(340, 540, 400, 610)
la4 = CreateEllipticRgn(330, 550, 390, 620)
la5 = CreateEllipticRgn(380, 585, 420, 620)
la6 = CreateEllipticRgn(400, 590, 410, 610)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, la1, 4
CombineRgn a1, a1, la2, 4
CombineRgn a1, a1, la3, 4
CombineRgn a1, a1, la4, 2
CombineRgn a1, a1, la5, 4
CombineRgn a1, a1, la6, 2

'=================================================================================================='
'Hamzah'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
hz1 = CreateEllipticRgn(210, 540, 260, 585)
hz2 = CreateEllipticRgn(220, 545, 262, 586)
hz3 = CreateRoundRectRgn(200, 580, 250, 590, 0, 0)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, hz1, 4
CombineRgn a1, a1, hz2, 2
CombineRgn a1, a1, hz3, 4

'=================================================================================================='
'Yak'
'=================================================================================================='
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips") = CRR'
'("Lebar -> Kanan","Tinggi-Atas","Lebar Kiri -> Kanan","Tinggi-Bawah") = CER'
yk1 = CreateEllipticRgn(20, 550, 135, 610)
yk2 = CreateEllipticRgn(25, 545, 110, 600)
yk3 = CreateEllipticRgn(65, 530, 140, 582)
yk4 = CreateEllipticRgn(80, 535, 150, 575)
yk5 = CreateEllipticRgn(50, 620, 70, 630)
yk6 = CreateEllipticRgn(80, 620, 100, 630)

'"Parent","Parent","Sub","2=muncul/4=disembunyikan"'
CombineRgn a1, a1, yk1, 4
CombineRgn a1, a1, yk2, 2
CombineRgn a1, a1, yk3, 4
CombineRgn a1, a1, yk4, 2
CombineRgn a1, a1, yk5, 4
CombineRgn a1, a1, yk6, 4


'=================================================================================================='
'Menampilkan Data Gabungan Keseluruhan'
'=================================================================================================='
SetWindowRgn Form1.hwnd, a1, True
End Sub

'=================================================================================================='
'Pengaturan Mouse Objek Bebas'
'=================================================================================================='
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage Form1.hwnd, &HA1, 2, 0&
End Sub

