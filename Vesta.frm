VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Vesta Project - The main data extraction software."
   ClientHeight    =   10020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16995
   LinkTopic       =   "Form1"
   ScaleHeight     =   668
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1133
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1935
      Left            =   10200
      TabIndex        =   32
      Top             =   240
      Width           =   3255
      Begin VB.CheckBox CheckSD 
         Caption         =   "Smooth Data by"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox CheckLB 
         Caption         =   "Plot line or bar"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CheckBox CheckOE 
         Caption         =   "Erase old experiment"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox SMVal 
         Height          =   285
         Left            =   1920
         TabIndex        =   34
         Text            =   "100"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox CheckUP 
         Caption         =   "Show when brightness increases"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Value           =   1  'Checked
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data for NN"
      Height          =   1935
      Left            =   13560
      TabIndex        =   26
      Top             =   240
      Width           =   3135
      Begin VB.PictureBox PicOCR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         ScaleHeight     =   39
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   79
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.PictureBox PicOCRbw 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1680
         ScaleHeight     =   39
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   79
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "This section can be used for the neural network. (in the optical character recognition version)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   40
         Top             =   1080
         Width           =   2655
      End
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
      TabIndex        =   23
      Top             =   5520
      Width           =   4215
   End
   Begin VB.PictureBox PerNERV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   10680
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   383
      TabIndex        =   22
      Top             =   8280
      Width           =   5775
      Begin VB.Shape LineCHA 
         Height          =   1095
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
   End
   Begin VB.PictureBox PerTOT 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   10680
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   383
      TabIndex        =   21
      Top             =   6480
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
      Left            =   13920
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   167
      TabIndex        =   20
      Top             =   4440
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
      Left            =   10680
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   167
      TabIndex        =   19
      Top             =   4440
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
      Left            =   13920
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   167
      TabIndex        =   18
      Top             =   2760
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
      Left            =   10680
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   167
      TabIndex        =   17
      Top             =   2760
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
      Caption         =   "LED coordinates"
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton Bcam 
      Caption         =   "Images from video camera"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Stop 
      Caption         =   "Stop"
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   9240
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
      Height          =   4095
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "Vesta.frx":0000
      Top             =   4440
      Width           =   5055
   End
   Begin VB.TextBox ASS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Text            =   "20"
      Top             =   8760
      Width           =   495
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   8040
      TabIndex        =   5
      Top             =   7920
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   5880
      TabIndex        =   4
      Top             =   7920
      Width           =   2175
   End
   Begin VB.CommandButton scanare 
      Caption         =   "Scan"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   9240
      Width           =   1935
   End
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   360
      Picture         =   "Vesta.frx":00CA
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
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   360
      Picture         =   "Vesta.frx":1F59
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   319
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      X1              =   16
      X2              =   352
      Y1              =   576
      Y2              =   576
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   368
      X2              =   368
      Y1              =   16
      Y2              =   648
   End
   Begin VB.Label Label9 
      Caption         =   "Up vs. Down - average brightness over time:"
      Height          =   255
      Left            =   10680
      TabIndex        =   41
      Top             =   8040
      Width           =   3855
   End
   Begin VB.Label Label7 
      Caption         =   "Front side brightness  (bottom - right)"
      Height          =   255
      Left            =   13920
      TabIndex        =   39
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Back side brightness  (bottom - left)"
      Height          =   255
      Left            =   10680
      TabIndex        =   38
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Average brightness over time:"
      Height          =   255
      Left            =   10680
      TabIndex        =   31
      Top             =   6240
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Front side brightness  (top - right)"
      Height          =   255
      Left            =   13920
      TabIndex        =   30
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Back side brightness (top - left)"
      Height          =   255
      Left            =   10680
      TabIndex        =   29
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   3  'Dot
      Height          =   7575
      Left            =   10200
      Top             =   2280
      Width           =   6495
   End
   Begin VB.Label LEDgr 
      BackStyle       =   0  'Transparent
      Caption         =   "Mean LED lighting per lot:"
      Height          =   255
      Left            =   5760
      TabIndex        =   25
      Top             =   5280
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum number of patients in group: 200"
      Height          =   255
      Left            =   5760
      TabIndex        =   24
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
      Caption         =   "Mean LED lighting brightness 0 images:"
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
      Left            =   1680
      TabIndex        =   8
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "ms"
      Height          =   255
      Index           =   0
      Left            =   3285
      TabIndex        =   7
      Top             =   8760
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'   ________________________________                          ____________________
'  /  Sensors                       \________________________/       v1.00        |
' |                                                                               |
' |            Name:  Sensors                                                     |
' |          Author:  Dr. Paul A. Gagniuc                                         |
' |                                                                               |
' |    Date Created:  November 2016                                               |
' |       Tested On:  Windows XP, Windows Vista, Windows 7, Windows 8, Windows 10 |
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
Dim Matrix(0 To 20, 0 To 10) As String 'matrice medie intre sabloanele pacientului
Dim Mat(0 To 20, 0 To 10) As String 'matrice led cate una
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


Private Sub Command1_Click()
    numeDIRE = Split(Dir1.Path, "\")(UBound(Split(Dir1.Path, "\")))
    MsgBox numeDIRE
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

    Dir1.Path = App.Path & "\one_subject_example" '
    File1.Path = Dir1.Path
    
    pTOT = 0
    
    Call draw_scale(PerTOT, 1)
    Call draw_scale(V1, 1)
    Call draw_scale(V2, 1)
    Call draw_scale(V3, 1)
    Call draw_scale(V4, 1)
    
    Call draw_scale(PerNERV, 1)
    
    LED(1) = "16,48"
    LED(2) = "29,47"
    LED(3) = "42,47"
    LED(4) = "60,44"
    LED(5) = "75,43"
    LED(6) = "91,45"
    LED(7) = "105,45"
    LED(8) = "122,43"
    LED(9) = "137,43"
    LED(10) = "152,43"
    LED(11) = "169,43"
    LED(12) = "183,42"
    LED(13) = "198,42"
    LED(14) = "212,42"
    LED(15) = "229,42"
    LED(16) = "242,40"
    LED(17) = "258,41"
    LED(18) = "273,40"
    LED(19) = "287,40"
    LED(20) = "302,41"
    LED(21) = "15,60"
    LED(22) = "30,61"
    LED(23) = "43,61"
    LED(24) = "63,61"
    LED(25) = "76,61"
    LED(26) = "89,60"
    LED(27) = "105,61"
    LED(28) = "119,61"
    LED(29) = "137,59"
    LED(30) = "152,60"
    LED(31) = "169,60"
    LED(32) = "183,59"
    LED(33) = "198,61"
    LED(34) = "213,58"
    LED(35) = "230,57"
    LED(36) = "244,57"
    LED(37) = "258,58"
    LED(38) = "275,58"
    LED(39) = "287,57"
    LED(40) = "303,58"
    LED(41) = "16,77"
    LED(42) = "30,81"
    LED(43) = "44,79"
    LED(44) = "61,78"
    LED(45) = "75,77"
    LED(46) = "91,77"
    LED(47) = "107,77"
    LED(48) = "122,77"
    LED(49) = "136,77"
    LED(50) = "153,77"
    LED(51) = "169,75"
    LED(52) = "185,75"
    LED(53) = "201,75"
    LED(54) = "213,75"
    LED(55) = "230,75"
    LED(56) = "244,75"
    LED(57) = "260,75"
    LED(58) = "276,73"
    LED(59) = "288,72"
    LED(60) = "304,72"
    LED(61) = "18,92"
    LED(62) = "29,91"
    LED(63) = "46,92"
    LED(64) = "61,92"
    LED(65) = "77,91"
    LED(66) = "91,91"
    LED(67) = "108,91"
    LED(68) = "122,88"
    LED(69) = "138,93"
    LED(70) = "154,91"
    LED(71) = "171,91"
    LED(72) = "185,91"
    LED(73) = "200,91"
    LED(74) = "216,92"
    LED(75) = "232,89"
    LED(76) = "248,90"
    LED(77) = "261,89"
    LED(78) = "279,91"
    LED(79) = "291,88"
    LED(80) = "306,87"
    LED(81) = "16,108"
    LED(82) = "30,106"
    LED(83) = "46,107"
    LED(84) = "59,107"
    LED(85) = "76,106"
    LED(86) = "91,106"
    LED(87) = "106,104"
    LED(88) = "120,106"
    LED(89) = "139,106"
    LED(90) = "154,107"
    LED(91) = "170,108"
    LED(92) = "187,106"
    LED(93) = "206,106"
    LED(94) = "215,106"
    LED(95) = "231,106"
    LED(96) = "248,106"
    LED(97) = "261,106"
    LED(98) = "277,105"
    LED(99) = "291,105"
    LED(100) = "307,105"
    LED(101) = "15,122"
    LED(102) = "29,122"
    LED(103) = "45,122"
    LED(104) = "64,121"
    LED(105) = "76,122"
    LED(106) = "90,122"
    LED(107) = "109,123"
    LED(108) = "121,123"
    LED(109) = "140,123"
    LED(110) = "151,122"
    LED(111) = "172,122"
    LED(112) = "183,121"
    LED(113) = "200,119"
    LED(114) = "217,120"
    LED(115) = "232,120"
    LED(116) = "246,120"
    LED(117) = "260,120"
    LED(118) = "278,120"
    LED(119) = "289,120"
    LED(120) = "305,120"
    LED(121) = "15,139"
    LED(122) = "31,140"
    LED(123) = "42,139"
    LED(124) = "61,139"
    LED(125) = "78,139"
    LED(126) = "93,139"
    LED(127) = "106,141"
    LED(128) = "122,140"
    LED(129) = "139,140"
    LED(130) = "154,137"
    LED(131) = "168,139"
    LED(132) = "184,139"
    LED(133) = "199,139"
    LED(134) = "219,138"
    LED(135) = "231,141"
    LED(136) = "250,139"
    LED(137) = "263,137"
    LED(138) = "277,137"
    LED(139) = "292,134"
    LED(140) = "305,135"
    LED(141) = "16,153"
    LED(142) = "29,153"
    LED(143) = "42,154"
    LED(144) = "64,155"
    LED(145) = "76,156"
    LED(146) = "94,156"
    LED(147) = "105,156"
    LED(148) = "122,156"
    LED(149) = "139,154"
    LED(150) = "155,156"
    LED(151) = "169,154"
    LED(152) = "183,154"
    LED(153) = "198,156"
    LED(154) = "216,156"
    LED(155) = "235,154"
    LED(156) = "245,154"
    LED(157) = "265,154"
    LED(158) = "277,152"
    LED(159) = "291,149"
    LED(160) = "308,149"
    LED(161) = "17,168"
    LED(162) = "32,166"
    LED(163) = "43,167"
    LED(164) = "62,168"
    LED(165) = "74,169"
    LED(166) = "89,170"
    LED(167) = "105,170"
    LED(168) = "122,170"
    LED(169) = "138,169"
    LED(170) = "155,170"
    LED(171) = "172,170"
    LED(172) = "184,170"
    LED(173) = "201,170"
    LED(174) = "217,171"
    LED(175) = "231,167"
    LED(176) = "248,171"
    LED(177) = "262,168"
    LED(178) = "275,167"
    LED(179) = "288,164"
    LED(180) = "303,162"
    LED(181) = "18,182"
    LED(182) = "29,183"
    LED(183) = "45,184"
    LED(184) = "61,182"
    LED(185) = "75,181"
    LED(186) = "91,183"
    LED(187) = "107,182"
    LED(188) = "121,186"
    LED(189) = "138,186"
    LED(190) = "155,186"
    LED(191) = "170,186"
    LED(192) = "186,186"
    LED(193) = "202,186"
    LED(194) = "216,185"
    LED(195) = "230,184"
    LED(196) = "250,184"
    LED(197) = "263,184"
    LED(198) = "274,181"
    LED(199) = "290,180"
    LED(200) = "304,178"
    
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
   ' Medie matrici la acelasi pacient
    ib = 20
    jb = 10
    Row = (SumP.ScaleHeight / jb)
    col = (SumP.ScaleWidth / ib)
    
    For j = 0 To jb - 1 'Rows
    
        For i = 0 To ib - 1 'Cols
            k = Val(Matrix(i, j)) / File1.ListCount
            'MsgBox Val(Matrix(i, j))
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
   ' Medie matrici intre pacienti
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
            
            Mat(i, j) = Val(h)
            
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
    
    qp = 0
    
    For j = 0 To jb - 1 'Rows
    
        For i = 0 To ib - 1 'Cols
        
            If qp < Val(Mat(i, j)) Then qp = Val(Mat(i, j))
        
        Next i
    
    Next j
    
    'sTXT_EXCEL = DrowMatrixForEXCEL(20, 10, Mat, "(P)", "Copy/Paste to EXCEL:")
    'result.Text = result.Text & sTXT_EXCEL
    'MsgBox qp
    
    For j = 0 To jb - 1 'Rows
    
        For i = 0 To ib - 1 'Cols
        
            Mat(i, j) = (100 / qp) * Val(Mat(i, j))
            'MsgBox (qp / 100)
        Next i
    
    Next j
    
    sTXT_EXCEL = DrowMatrixForEXCEL(20, 10, Mat, "", "")
    
    numeDIRE = Split(Dir1.Path, "\")(UBound(Split(Dir1.Path, "\")))
    
    'Start append text to file
    FileNum = FreeFile
    Open App.Path & "\" & numeDIRE & ".txt" For Append As FileNum
        Print #FileNum, sTXT_EXCEL
    Close FileNum

    'result.Text = result.Text & sTXT_EXCEL
    
    B = B / 200 'media tuturor ledurilor
    
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


Function DrowMatrixForEXCEL(ib, jb, ByVal M As Variant, ByVal model As String, ByVal msg As String) As String

    '------ Show Matrix in Text OBJ -------------------------------------------
    For j = 0 To jb - 1 'cols
    
        For i = 0 To ib - 1 'Rows
    
    
        
        If M(i, j) <> "" Then v = Round(M(i, j), 1) Else v = 0 '
            
            If j = jb Then o = "" Else o = Chr(9)
            ct = ct & v & o
            
        Next i
    
    ct = ct & Chr(9) & j & vbCrLf
    
    Next j
    '--------------------------------------------------------------------------
    DrowMatrixForEXCEL = ct & vbCrLf '& " M[" & Val(jb) & "," & Val(ib) & "]" & vbCrLf & ct & vbCrLf & vbCrLf
    '--------------------------------------------------------------------------

End Function

