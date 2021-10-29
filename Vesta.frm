VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Photon-pixel coupling: A method for parallel acquisition of electrical signals in scientific investigations"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17430
   LinkTopic       =   "Form1"
   ScaleHeight     =   643
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1162
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicOCRbw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   12360
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   32
      Top             =   480
      Width           =   1215
   End
   Begin VB.PictureBox PicOCR 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   11040
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   31
      Top             =   480
      Width           =   1215
   End
   Begin VB.PictureBox PacT 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   5760
      ScaleHeight     =   135
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   279
      TabIndex        =   28
      Top             =   5520
      Width           =   4215
   End
   Begin VB.CheckBox CheckUP 
      Caption         =   "Show when brightness increases"
      Height          =   255
      Left            =   11160
      TabIndex        =   27
      Top             =   8040
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.PictureBox PerNERV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   11040
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   383
      TabIndex        =   26
      Top             =   6480
      Width           =   5775
      Begin VB.Shape LineCHA 
         Height          =   1095
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
   End
   Begin VB.TextBox SMVal 
      Height          =   285
      Left            =   12840
      TabIndex        =   25
      Text            =   "100"
      Top             =   8400
      Width           =   495
   End
   Begin VB.CheckBox CheckOE 
      Caption         =   "Erase old experiment"
      Height          =   255
      Left            =   11160
      TabIndex        =   24
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CheckBox CheckLB 
      Caption         =   "Plot line or bar"
      Height          =   255
      Left            =   11160
      TabIndex        =   23
      Top             =   8760
      Width           =   1815
   End
   Begin VB.CheckBox CheckSD 
      Caption         =   "Smooth Data by"
      Height          =   255
      Left            =   11160
      TabIndex        =   22
      Top             =   8400
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.PictureBox PerTOT 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   11040
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   383
      TabIndex        =   21
      Top             =   4920
      Width           =   5775
      Begin VB.Shape LineTOT 
         Height          =   1095
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
   End
   Begin VB.PictureBox V4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   14280
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   167
      TabIndex        =   20
      Top             =   3360
      Width           =   2535
      Begin VB.Shape Shape4 
         Height          =   1095
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
   End
   Begin VB.PictureBox V2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   11040
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   167
      TabIndex        =   19
      Top             =   3360
      Width           =   2535
      Begin VB.Shape Shape2 
         Height          =   1095
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
   End
   Begin VB.PictureBox V3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   14280
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   167
      TabIndex        =   18
      Top             =   1800
      Width           =   2535
      Begin VB.Shape Shape3 
         Height          =   1095
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
   End
   Begin VB.PictureBox V1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   11040
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   167
      TabIndex        =   17
      Top             =   1800
      Width           =   2535
      Begin VB.Shape Shape1 
         Height          =   1095
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
   End
   Begin VB.CommandButton Bcor 
      Caption         =   "LED coord."
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton Bcam 
      Caption         =   "Camera"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Stop 
      Caption         =   "Stop"
      Height          =   495
      Left            =   2640
      TabIndex        =   13
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CheckBox CheckGR 
      Caption         =   "Plot grid"
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.PictureBox SumP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   5760
      ScaleHeight     =   135
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   279
      TabIndex        =   10
      Top             =   3120
      Width           =   4215
   End
   Begin VB.TextBox result 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   4440
      Width           =   4815
   End
   Begin VB.TextBox ASS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Text            =   "20"
      Top             =   8160
      Width           =   495
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   7920
      TabIndex        =   5
      Top             =   7920
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   5400
      TabIndex        =   4
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton scanare 
      Caption         =   "Scan Automat"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   8640
      Width           =   1935
   End
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000040&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   240
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   319
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   5760
      ScaleHeight     =   135
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   279
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin VB.PictureBox Center_patt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   240
      Picture         =   "Vesta.frx":0000
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   319
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.Line Line4 
      X1              =   696
      X2              =   696
      Y1              =   24
      Y2              =   624
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   360
      Y1              =   16
      Y2              =   624
   End
   Begin VB.Line Line1 
      X1              =   728
      X2              =   1136
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Label LEDgr 
      BackStyle       =   0  'Transparent
      Caption         =   "Mean LED lighting per lot:"
      Height          =   255
      Left            =   5760
      TabIndex        =   30
      Top             =   5280
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum number of patients in group: 200"
      Height          =   255
      Left            =   5760
      TabIndex        =   29
      Top             =   7560
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LED brightness in individual readings:"
      Height          =   255
      Left            =   5760
      TabIndex        =   16
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label LEDTXT 
      BackStyle       =   0  'Transparent
      Caption         =   "Mean LED lighting brightness the 0 images:"
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Animation step"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "ms"
      Height          =   255
      Index           =   0
      Left            =   3045
      TabIndex        =   7
      Top             =   8160
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'   ________________________________                          ____________________
'  /  Senzori                       \________________________/       v1.00        |
' |                                                                               |
' |            Name:  Senzori                                                     |
' |          Author:  Paul A. Gagniuc                                             |
' |                                                                               |
' |    Date Created:  November 2014                                               |
' |       Tested On:  Windows XP, Windows Vista, Windows 7, Windows 8             |
' |           Email:  paul_gagniuc@acad.ro                                        |
' |             Use:  diabetes prediction                                         |
' |                                                                               |
' |                  _____________________________                                |
' |_________________/                             \_______________________________|
'

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim stopEXP As Boolean
Dim OLDprocent As Variant

Dim v1OLD As Variant
Dim v2OLD As Variant
Dim v3OLD As Variant
Dim v4OLD As Variant

Dim d() As Double
Dim MAS() As Double
Dim C1() As Double
Dim C2() As Double
Dim C3() As Double
Dim C4() As Double
Dim iter As Integer


Dim LED(1 To 200) As String
Dim Matrix(0 To 20, 0 To 10) As String
Dim MatrixTOT() As String
Dim pTOT As Variant

Private Sub Bcam_Click()
Pic2.Visible = False
Center_patt.Visible = True
End Sub

Private Sub Bcor_Click()
Pic2.Visible = True
Center_patt.Visible = False
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
'Text1.Text = Dir1.Path
End Sub

Private Sub File1_Click()
'MsgBox Dir1.Path & "\" & File1.FileName
Center_patt.Picture = LoadPicture(Dir1.Path & "\" & File1.FileName)
End Sub

Private Sub Form_Load()


ReDim MatrixTOT(0 To 20, 0 To 10, 0 To 200) As String

Dir1.Path = App.Path & "\pacienti" '
File1.Path = Dir1.Path

pTOT = 0

Call draw_scale(PerTOT, 1)
Call draw_scale(V1, 1)
Call draw_scale(V2, 1)
Call draw_scale(V3, 1)
Call draw_scale(V4, 1)

Call draw_scale(PerNERV, 1)


LED(1) = "15,35"
LED(2) = "29,35"
LED(3) = "44,35"
LED(4) = "58,32"
LED(5) = "74,31"
LED(6) = "87,30"
LED(7) = "100,29"
LED(8) = "116,30"
LED(9) = "136,28"
LED(10) = "154,29"
LED(11) = "165,29"
LED(12) = "179,27"
LED(13) = "195,27"
LED(14) = "211,26"
LED(15) = "226,27"
LED(16) = "239,28"
LED(17) = "259,32"
LED(18) = "270,28"
LED(19) = "284,26"
LED(20) = "301,27"
LED(21) = "13,48"
LED(22) = "27,48"
LED(23) = "44,47"
LED(24) = "59,49"
LED(25) = "74,49"
LED(26) = "88,47"
LED(27) = "102,47"
LED(28) = "117,47"
LED(29) = "135,48"
LED(30) = "153,48"
LED(31) = "165,47"
LED(32) = "179,47"
LED(33) = "198,45"
LED(34) = "213,44"
LED(35) = "225,43"
LED(36) = "242,43"
LED(37) = "256,44"
LED(38) = "273,44"
LED(39) = "287,44"
LED(40) = "300,42"
LED(41) = "13,65"
LED(42) = "30,66"
LED(43) = "46,63"
LED(44) = "59,64"
LED(45) = "75,63"
LED(46) = "89,63"
LED(47) = "104,61"
LED(48) = "119,62"
LED(49) = "134,61"
LED(50) = "149,61"
LED(51) = "165,61"
LED(52) = "184,61"
LED(53) = "197,61"
LED(54) = "210,61"
LED(55) = "228,60"
LED(56) = "242,59"
LED(57) = "257,59"
LED(58) = "275,59"
LED(59) = "285,60"
LED(60) = "300,58"
LED(61) = "13,79"
LED(62) = "27,80"
LED(63) = "44,79"
LED(64) = "58,79"
LED(65) = "74,76"
LED(66) = "87,76"
LED(67) = "106,75"
LED(68) = "118,76"
LED(69) = "133,79"
LED(70) = "151,78"
LED(71) = "168,78"
LED(72) = "183,78"
LED(73) = "197,75"
LED(74) = "214,79"
LED(75) = "231,75"
LED(76) = "247,79"
LED(77) = "260,79"
LED(78) = "277,77"
LED(79) = "289,75"
LED(80) = "300,75"
LED(81) = "13,96"
LED(82) = "28,96"
LED(83) = "41,95"
LED(84) = "57,96"
LED(85) = "75,96"
LED(86) = "88,95"
LED(87) = "107,93"
LED(88) = "121,93"
LED(89) = "135,93"
LED(90) = "152,92"
LED(91) = "169,95"
LED(92) = "183,92"
LED(93) = "203,93"
LED(94) = "215,93"
LED(95) = "229,93"
LED(96) = "245,94"
LED(97) = "259,93"
LED(98) = "275,92"
LED(99) = "289,92"
LED(100) = "304,91"
LED(101) = "15,108"
LED(102) = "28,109"
LED(103) = "40,110"
LED(104) = "60,111"
LED(105) = "74,110"
LED(106) = "86,109"
LED(107) = "107,111"
LED(108) = "122,109"
LED(109) = "136,109"
LED(110) = "150,109"
LED(111) = "169,109"
LED(112) = "181,109"
LED(113) = "200,109"
LED(114) = "215,113"
LED(115) = "231,110"
LED(116) = "248,106"
LED(117) = "262,108"
LED(118) = "272,107"
LED(119) = "290,107"
LED(120) = "304,107"
LED(121) = "16,125"
LED(122) = "29,125"
LED(123) = "42,125"
LED(124) = "59,127"
LED(125) = "77,126"
LED(126) = "92,126"
LED(127) = "104,127"
LED(128) = "122,126"
LED(129) = "138,125"
LED(130) = "150,126"
LED(131) = "169,125"
LED(132) = "185,126"
LED(133) = "200,127"
LED(134) = "217,127"
LED(135) = "230,126"
LED(136) = "246,125"
LED(137) = "259,123"
LED(138) = "276,122"
LED(139) = "289,123"
LED(140) = "308,123"
LED(141) = "16,144"
LED(142) = "28,140"
LED(143) = "42,142"
LED(144) = "64,141"
LED(145) = "76,143"
LED(146) = "89,142"
LED(147) = "108,142"
LED(148) = "119,144"
LED(149) = "134,142"
LED(150) = "152,142"
LED(151) = "168,140"
LED(152) = "183,141"
LED(153) = "199,141"
LED(154) = "214,141"
LED(155) = "231,140"
LED(156) = "244,141"
LED(157) = "261,139"
LED(158) = "277,138"
LED(159) = "289,137"
LED(160) = "305,134"
LED(161) = "16,153"
LED(162) = "28,154"
LED(163) = "44,156"
LED(164) = "58,156"
LED(165) = "72,157"
LED(166) = "87,157"
LED(167) = "104,155"
LED(168) = "120,155"
LED(169) = "136,155"
LED(170) = "151,156"
LED(171) = "167,156"
LED(172) = "182,156"
LED(173) = "198,155"
LED(174) = "215,154"
LED(175) = "229,153"
LED(176) = "244,153"
LED(177) = "260,154"
LED(178) = "274,151"
LED(179) = "289,151"
LED(180) = "303,151"
LED(181) = "15,168"
LED(182) = "28,165"
LED(183) = "44,169"
LED(184) = "60,171"
LED(185) = "72,170"
LED(186) = "85,170"
LED(187) = "104,171"
LED(188) = "119,169"
LED(189) = "138,172"
LED(190) = "153,173"
LED(191) = "170,173"
LED(192) = "184,172"
LED(193) = "199,170"
LED(194) = "216,171"
LED(195) = "229,170"
LED(196) = "246,172"
LED(197) = "261,172"
LED(198) = "274,174"
LED(199) = "288,168"
LED(200) = "302,167"


End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
stopEXP = True
End
End Sub

Private Sub scanare_Click()

stopEXP = False
OLDprocent = 0

v1OLD = 0
v2OLD = 0
v3OLD = 0
v4OLD = 0


ReDim MAS(File1.ListCount) As Double
ReDim C1(File1.ListCount) As Double
ReDim C2(File1.ListCount) As Double
ReDim C3(File1.ListCount) As Double
ReDim C4(File1.ListCount) As Double


If CheckOE.Value = 1 Then
    PerTOT.Cls
    V1.Cls
    V2.Cls
    V3.Cls
    V4.Cls
End If


For i = 1 To File1.ListCount

    If stopEXP = True Then GoTo 1
    
    p = Dir1.Path & "\" & File1.List(i)

    If InStr(p, ".") > 0 Then

        eJPG = Split(p, ".")(UBound(Split(p, ".")))
        
        If eJPG = "jpg" Or eJPG = "JPG" Then
            Center_patt.Picture = LoadPicture(p)
            Call StartSc(i)
            DoEvents
            Sleep (CLng(ASS.Text))
        End If
    End If



Next i
'-------------------------------------------------------------------------------------
' Media matricelor la acelasi pacient
ib = 20
jb = 10
Row = (SumP.ScaleHeight / jb)
col = (SumP.ScaleWidth / ib)
For j = 0 To jb - 1 'Rows

    For i = 0 To ib - 1 'Cols
        k = Val(Matrix(i, j)) / File1.ListCount
        SumP.Line (col * i, Row * j)-(col * (i + 1), Row * (j + 1)), RGB(k, k, k), BF
        
        If CheckGR.Value = 1 Then
            SumP.Line (col * i, 0)-(col * i, SumP.ScaleHeight), vbBlack, B
            SumP.Line (0, Row * j)-(SumP.ScaleWidth, Row * j), vbBlack, B
        End If
    Next i

Next j


For j = 0 To jb - 1 'Rows
    For i = 0 To ib - 1 'Cols
        MatrixTOT(i, j, pTOT) = Val(Matrix(i, j)) / File1.ListCount ' e var k
        Matrix(i, j) = 0
    Next i
Next j

pTOT = pTOT + 1
'-------------------------------------------------------------------------------------
' Media matricelor intre pacienti
For j = 0 To jb - 1 'Rows
    For i = 0 To ib - 1 'Cols
        
        For s = 0 To pTOT '- 1
            w = Val(w) + Round(Val(MatrixTOT(i, j, s)))
        Next s
        
        k = w / pTOT
        w = 0
        
        PacT.Line (col * i, Row * j)-(col * (i + 1), Row * (j + 1)), RGB(k, k, k), BF
        
        If CheckGR.Value = 1 Then
            PacT.Line (col * i, 0)-(col * i, PacT.ScaleHeight), vbBlack, B
            PacT.Line (0, Row * j)-(PacT.ScaleWidth, Row * j), vbBlack, B
        End If
        
    Next i
Next j
'-------------------------------------------------------------------------------------
LEDgr.Caption = "Mean LED lighting per lot: " & pTOT
'-------------------------------------------------------------------------------------


LEDTXT.Caption = "Mean LED brightness between the " & File1.ListCount & " images:"


If CheckSD.Value = 1 Then
    Call Abraziv(PerTOT, MAS)
    Call Abraziv(V1, C1)
    Call Abraziv(V2, C2)
    Call Abraziv(V3, C3)
    Call Abraziv(V4, C4)
    

End If


Call UP_or_DOWN(PerNERV, MAS)

result.Text = result.Text & "Numar imagini analizate: " & File1.ListCount & vbCrLf
1:

End Sub

Private Sub Stop_Click()
stopEXP = True
End Sub




Function StartSc(ByVal fil As Integer)



ib = 20
jb = 10

g = 0

Pic1.Cls


Row = (Pic1.ScaleHeight / jb)
col = (Pic1.ScaleWidth / ib)

RowOCR = (PicOCR.ScaleHeight / jb)
ColOCR = (PicOCR.ScaleWidth / ib)

vRow = (V1.ScaleHeight / jb)
vCol = (V1.ScaleWidth / ib)

    tY = PerTOT.ScaleHeight / Val(vbWhite)
    tX = PerTOT.ScaleWidth / File1.ListCount


    tYV = V1.ScaleHeight / Val(vbWhite)
    tXV = V1.ScaleWidth / File1.ListCount

For j = 0 To jb - 1 'Rows

    For i = 0 To ib - 1 'Cols
    
    
        g = g + 1
    
        X = Val(Split(LED(g), ",")(0))
        Y = Val(Split(LED(g), ",")(1))


    
        a = Center_patt.Point(X, Y)
        h = Center_patt.Point(X, Y) / (vbWhite / 255)


        B = B + a
    
        Matrix(i, j) = Val(Matrix(i, j)) + Val(h)
        
        Pic1.Line (col * i, Row * j)-(col * (i + 1), Row * (j + 1)), a, BF
        
        '--------------------------------------------------------------------------------
        PicOCR.Line (ColOCR * i, RowOCR * j)-(ColOCR * (i + 1), RowOCR * (j + 1)), a, BF
        
        
        If a > (RGB(255, 255, 255) / 2) Then aOCR = vbWhite Else aOCR = vbBlack
        PicOCRbw.Line (ColOCR * i, RowOCR * j)-(ColOCR * (i + 1), RowOCR * (j + 1)), aOCR, BF
        '--------------------------------------------------------------------------------
        
        If CheckGR.Value = 1 Then
            Pic1.Line (col * i, 0)-(col * i, Pic1.ScaleHeight), vbBlack, B
            Pic1.Line (0, Row * j)-(Pic1.ScaleWidth, Row * j), vbBlack, B
        End If

        If j < 5 And i < 10 Then
            V1a = V1a + a
        End If

        If j > 5 And i < 10 Then
            V2a = V2a + a
        End If

        If j < 5 And i > 10 Then
            V3a = V3a + a
        End If

        If j > 5 And i > 10 Then
            V4a = V4a + a
        End If



    Next i

Next j



B = B / 200 ' media tuturor ledurilor

If CheckLB.Value = 1 Then
    PerTOT.Line (tX * (fil - 1), PerTOT.ScaleHeight - (tY * OLDprocent))-(tX * fil, PerTOT.ScaleHeight - (tY * B)), vbBlack
Else
    PerTOT.Line (tX * (fil - 1), PerTOT.ScaleHeight)-(tX * fil, PerTOT.ScaleHeight - (tY * B)), vbBlack, BF
End If
OLDprocent = B
MAS(fil) = B



v1NEW = V1a / 50
If CheckLB.Value = 1 Then
    V1.Line (tXV * (fil - 1), V1.ScaleHeight - (tYV * v1OLD))-(tXV * fil, V1.ScaleHeight - (tYV * v1NEW)), vbBlack
Else
    V1.Line (tXV * (fil - 1), V1.ScaleHeight)-(tXV * fil, V1.ScaleHeight - (tYV * v1NEW)), vbBlack, BF
End If
v1OLD = v1NEW
C1(fil) = v1NEW

v2NEW = V2a / 50
If CheckLB.Value = 1 Then
    V2.Line (tXV * (fil - 1), V2.ScaleHeight - (tYV * v2OLD))-(tXV * fil, V2.ScaleHeight - (tYV * v2NEW)), vbBlack
Else
    V2.Line (tXV * (fil - 1), V2.ScaleHeight)-(tXV * fil, V2.ScaleHeight - (tYV * v2NEW)), vbBlack, BF
End If
v2OLD = v2NEW
C2(fil) = v2NEW

v3NEW = V3a / 50
If CheckLB.Value = 1 Then
    V3.Line (tXV * (fil - 1), V3.ScaleHeight - (tYV * v3OLD))-(tXV * fil, V3.ScaleHeight - (tYV * v3NEW)), vbBlack
Else
    V3.Line (tXV * (fil - 1), V3.ScaleHeight)-(tXV * fil, V3.ScaleHeight - (tYV * v3NEW)), vbBlack, BF
End If
v3OLD = v3NEW
C3(fil) = v3NEW

v4NEW = V4a / 50
If CheckLB.Value = 1 Then
    V4.Line (tXV * (fil - 1), V4.ScaleHeight - (tYV * v4OLD))-(tXV * fil, V4.ScaleHeight - (tYV * v4NEW)), vbBlack
Else
    V4.Line (tXV * (fil - 1), V4.ScaleHeight)-(tXV * fil, V4.ScaleHeight - (tYV * v4NEW)), vbBlack, BF
End If
v4OLD = v4NEW
C4(fil) = v4NEW





If Pic2.Visible = True Then
    For i = 1 To 200

        X = Split(LED(i), ",")(0)
        Y = Split(LED(i), ",")(1)

        a = Center_patt.Point(X, Y)

        Pic2.Circle (X, Y), 5, a

    Next i
End If

DoEvents


End Function



Function Abraziv(ByRef pb As PictureBox, ByRef dat() As Double)

For q = 1 To Val(SMVal.Text)
    iter = iter + 1
    dat() = smoothData1(dat)
Next q



tY = pb.ScaleHeight / Val(vbWhite)
tX = pb.ScaleWidth / File1.ListCount
pb.Cls

OLDc = 0 'dat(i) 'pb.ScaleHeight

For i = 1 To UBound(dat)

    If CheckLB.Value = 1 Then
        pb.Line (tX * (i - 1), pb.ScaleHeight - (tY * OLDc))-(tX * i, pb.ScaleHeight - (tY * dat(i))), vbRed ', B
    Else
        pb.Line (tX * i, pb.ScaleHeight)-(tX * i + 1, pb.ScaleHeight - (tY * dat(i))), vbRed, BF
    End If
    
    If CheckUP.Value = 1 And OLDc < dat(i) Then
        pb.Line (tX * i, pb.ScaleHeight)-(tX * i + 1, pb.ScaleHeight - 10), RGB(45, 77, 88), BF
    End If

    OLDc = dat(i)

Next i

End Function

Private Function smoothData1(dat)
'Arithmetic
For n = LBound(dat) + 2 To UBound(dat) - 2
dat(n) = (dat(n - 1) + dat(n + 1)) / 2
Next
smoothData1 = dat
End Function

Private Sub showData(pb As PictureBox, dat, col As Long)
For X = LBound(dat) + 1 To UBound(dat)
pb.Line (X - 1, pb.ScaleHeight - dat(X - 1))-(X, pb.ScaleHeight - dat(X)), col
Next
End Sub




Function UP_or_DOWN(ByRef pb As PictureBox, ByRef dd() As Double)

ReDim d(1 To UBound(dd) + 1) As Double

For i = 2 To UBound(dd) - 1

    If dd(i - 1) > dd(i) Then
        d(i) = dd(i - 1) - dd(i)
    Else
        d(i) = dd(i) - dd(i - 1)
    End If

    If MAXd < d(i) Then
        MAXd = d(i)
    End If

Next i


tY = pb.ScaleHeight / MAXd
tX = pb.ScaleWidth / File1.ListCount
pb.Cls

OLDc = 0 'd(i) 'pb.ScaleHeight

For i = 1 To UBound(d)

    If CheckLB.Value = 1 Then
        pb.Line (tX * (i - 1), pb.ScaleHeight - (tY * OLDc))-(tX * i, pb.ScaleHeight - (tY * d(i))), vbRed   ', B
    Else
        pb.Line (tX * i, pb.ScaleHeight)-(tX * i + 1, pb.ScaleHeight - (tY * d(i))), vbRed, BF
    End If

    OLDc = d(i)

Next i

End Function



Function draw_scale(ByRef pict As PictureBox, ByVal k_stat As Integer)
Dim zx, qx, zy, qy As Variant
Dim sp As Variant
Dim i As Integer

'Form1.Cls

'X axis on pict OBJ
'-------------------------------------
sp = pict.ScaleWidth / k_stat
For i = 0 To k_stat

    zx = pict.Left + (sp * i)
    qx = zx
    zy = pict.Top + pict.ScaleHeight
    qy = pict.Top + pict.ScaleHeight + 6

    If k_stat < 10 Then
        Form1.CurrentX = zx - 6
        Form1.CurrentY = qy
        Form1.Print i & "h"
    End If

    Form1.Line (zx, zy)-(qx, qy), &H808080

Next i
'-------------------------------------

'Y axis on pict OBJ
'-------------------------------------
    zx = pict.Left - 6
    qx = pict.Left
    zy = pict.Top
    qy = zy
    Form1.Line (zx, zy)-(qx, qy), &H808080
    Form1.CurrentX = zx - 20
    Form1.CurrentY = qy - 6
    Form1.Print "100"

    zx = pict.Left - 6
    qx = pict.Left
    zy = pict.Top + pict.ScaleHeight
    qy = zy
    Form1.Line (zx, zy)-(qx, qy), &H808080
    Form1.CurrentX = zx - 7
    Form1.CurrentY = qy - 6
    Form1.Print "0"
'-------------------------------------

End Function




Function Grafic_apeleaza(ByRef X As Single, ByRef Pic As PictureBox, ByRef lin As Shape)
    lin.Left = X
    q = (File1.ListCount / PerTOT.ScaleWidth) * X
    Center_patt.Picture = LoadPicture(Dir1.Path & "\" & File1.List(Round(q)))
    Call Mouse_scan
End Function


Private Sub PerNERV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Grafic_apeleaza(X, PerNERV, LineCHA)
    LineCHA.Visible = True
    LineTOT.Visible = False
    Shape1.Visible = False
    Shape2.Visible = False
    Shape3.Visible = False
    Shape4.Visible = False
End Sub

Private Sub PerTOT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Grafic_apeleaza(X, PerTOT, LineTOT)
    LineCHA.Visible = False
    LineTOT.Visible = True
    Shape1.Visible = False
    Shape2.Visible = False
    Shape3.Visible = False
    Shape4.Visible = False
End Sub


Private Sub V1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Grafic_apeleaza(X, V1, Shape1)
    LineCHA.Visible = False
    LineTOT.Visible = False
    Shape1.Visible = True
    Shape2.Visible = False
    Shape3.Visible = False
    Shape4.Visible = False

    aC = Pic1.ScaleHeight / 2
    bC = Pic1.ScaleWidth / 2
    Pic1.Line (0, 0)-(bC, aC), vbRed, B
End Sub

Private Sub V2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Grafic_apeleaza(X, V2, Shape2)
    LineCHA.Visible = False
    LineTOT.Visible = False
    Shape1.Visible = False
    Shape2.Visible = True
    Shape3.Visible = False
    Shape4.Visible = False

    aC = Pic1.ScaleHeight / 2
    bC = Pic1.ScaleWidth / 2
    Pic1.Line (0, aC)-(bC, Pic1.ScaleHeight), vbRed, B
End Sub

Private Sub V3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Grafic_apeleaza(X, V3, Shape3)
    LineCHA.Visible = False
    LineTOT.Visible = False
    Shape1.Visible = False
    Shape2.Visible = False
    Shape3.Visible = True
    Shape4.Visible = False

    aC = Pic1.ScaleHeight / 2
    bC = Pic1.ScaleWidth / 2
    Pic1.Line (bC, 0)-(Pic1.ScaleWidth, aC), vbRed, B
End Sub

Private Sub V4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Grafic_apeleaza(X, V4, Shape4)
    LineCHA.Visible = False
    LineTOT.Visible = False
    Shape1.Visible = False
    Shape2.Visible = False
    Shape3.Visible = False
    Shape4.Visible = True

    aC = Pic1.ScaleHeight / 2
    bC = Pic1.ScaleWidth / 2
    Pic1.Line (bC, aC)-(Pic1.ScaleWidth, Pic1.ScaleHeight), vbRed, B
End Sub


Function Mouse_scan()

ib = 20
jb = 10

g = 0

Pic1.Cls

Row = (Pic1.ScaleHeight / jb)
col = (Pic1.ScaleWidth / ib)


    tY = PerTOT.ScaleHeight / Val(vbWhite)
    tX = PerTOT.ScaleWidth / File1.ListCount


For j = 0 To jb - 1 'Rows

    For i = 0 To ib - 1 'Cols
    
    
    g = g + 1
    
    X = Val(Split(LED(g), ",")(0))
    Y = Val(Split(LED(g), ",")(1))


    
    a = Center_patt.Point(X, Y)
    

        Pic1.Line (col * i, Row * j)-(col * (i + 1), Row * (j + 1)), a, BF

        If CheckGR.Value = 1 Then
            Pic1.Line (col * i, 0)-(col * i, Pic1.ScaleHeight), vbBlack, B
            Pic1.Line (0, Row * j)-(Pic1.ScaleWidth, Row * j), vbBlack, B
        End If
        
    Next i

Next j

If Pic2.Visible = True Then
    For i = 1 To 200

        X = Split(LED(i), ",")(0)
        Y = Split(LED(i), ",")(1)

        a = Center_patt.Point(X, Y)

        Pic2.Circle (X, Y), 5, a

    Next i
End If

DoEvents
End Function

