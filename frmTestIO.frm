VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTestIO 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin dp6.XPButton XPButton1 
      Height          =   945
      Index           =   2
      Left            =   60
      TabIndex        =   239
      Top             =   10530
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1667
      TxtText         =   "Test I/O"
      TxtTop          =   35
      TxtLeft         =   45
      BTYPE           =   3
      IMGTOP          =   5
      IMGLEFT         =   5
      ICONA           =   "..\bitmap\icone\enet1.ico"
      ImgW            =   50
      ImgH            =   20
      ImgAllarga      =   0   'False
      TX              =   "      "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin dp6.XPButton XPButton1 
      Height          =   945
      Index           =   1
      Left            =   60
      TabIndex        =   238
      Top             =   10530
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1667
      TxtText         =   "Logical scheme"
      TxtTop          =   35
      TxtLeft         =   25
      BTYPE           =   3
      IMGTOP          =   5
      IMGLEFT         =   5
      ICONA           =   "..\bitmap\icone\RSLGX0052.ico"
      ImgW            =   50
      ImgH            =   20
      ImgAllarga      =   0   'False
      TX              =   "      "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin dp6.XPButton XPButton1 
      Height          =   945
      Index           =   0
      Left            =   13050
      TabIndex        =   237
      Top             =   10530
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1667
      TxtText         =   "Alarm info"
      TxtTop          =   35
      TxtLeft         =   50
      BTYPE           =   3
      IMGTOP          =   5
      IMGLEFT         =   5
      ICONA           =   "..\bitmap\MSGBOX02.ICO"
      ImgW            =   50
      ImgH            =   20
      ImgAllarga      =   0   'False
      TX              =   "      "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10485
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   18494
      _Version        =   393216
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   882
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Packpipe"
      TabPicture(0)   =   "frmTestIO.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Blowing"
      TabPicture(1)   =   "frmTestIO.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Shape1(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Shape1(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Shape1(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Shape1(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Shape1(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Shape1(5)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Shape1(6)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Shape1(7)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Shape1(8)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Shape1(9)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Shape1(10)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Shape1(11)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Shape1(12)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Shape1(13)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Shape1(14)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Shape1(15)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Shape1(16)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Shape1(17)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Shape1(18)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Shape1(19)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Shape1(20)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Shape1(21)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Shape1(22)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Shape1(23)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Shape1(24)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Shape1(25)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Shape1(26)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Shape1(27)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Shape1(28)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Shape1(29)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Shape1(30)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Shape1(31)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Shape1(32)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Shape1(33)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Shape1(34)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Shape1(35)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Shape1(36)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Shape1(37)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Shape1(38)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Shape1(39)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Shape1(40)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Shape1(41)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Shape1(42)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Shape1(43)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Shape1(44)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "Shape1(45)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "Shape1(46)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "Shape1(47)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "Shape1(48)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "Shape1(49)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "Shape1(50)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "Shape1(51)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "Shape1(52)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "Shape1(53)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "Shape1(54)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "Shape1(55)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "FrameLav"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "Ingresso(0)"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "Ingresso(1)"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "Ingresso(2)"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "Ingresso(3)"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "TimerLocale"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "Ingresso(4)"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "Ingresso(5)"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "Ingresso(6)"
      Tab(1).Control(64).Enabled=   0   'False
      Tab(1).Control(65)=   "Ingresso(7)"
      Tab(1).Control(65).Enabled=   0   'False
      Tab(1).Control(66)=   "Ingresso(8)"
      Tab(1).Control(66).Enabled=   0   'False
      Tab(1).Control(67)=   "Ingresso(9)"
      Tab(1).Control(67).Enabled=   0   'False
      Tab(1).Control(68)=   "Ingresso(10)"
      Tab(1).Control(68).Enabled=   0   'False
      Tab(1).Control(69)=   "Ingresso(11)"
      Tab(1).Control(69).Enabled=   0   'False
      Tab(1).Control(70)=   "Ingresso(12)"
      Tab(1).Control(70).Enabled=   0   'False
      Tab(1).Control(71)=   "Ingresso(13)"
      Tab(1).Control(71).Enabled=   0   'False
      Tab(1).Control(72)=   "Ingresso(14)"
      Tab(1).Control(72).Enabled=   0   'False
      Tab(1).Control(73)=   "Ingresso(15)"
      Tab(1).Control(73).Enabled=   0   'False
      Tab(1).Control(74)=   "Ingresso(16)"
      Tab(1).Control(74).Enabled=   0   'False
      Tab(1).Control(75)=   "Ingresso(17)"
      Tab(1).Control(75).Enabled=   0   'False
      Tab(1).Control(76)=   "Ingresso(18)"
      Tab(1).Control(76).Enabled=   0   'False
      Tab(1).Control(77)=   "Ingresso(19)"
      Tab(1).Control(77).Enabled=   0   'False
      Tab(1).Control(78)=   "Ingresso(20)"
      Tab(1).Control(78).Enabled=   0   'False
      Tab(1).Control(79)=   "Ingresso(21)"
      Tab(1).Control(79).Enabled=   0   'False
      Tab(1).Control(80)=   "Ingresso(22)"
      Tab(1).Control(80).Enabled=   0   'False
      Tab(1).Control(81)=   "Ingresso(23)"
      Tab(1).Control(81).Enabled=   0   'False
      Tab(1).Control(82)=   "Ingresso(24)"
      Tab(1).Control(82).Enabled=   0   'False
      Tab(1).Control(83)=   "Ingresso(25)"
      Tab(1).Control(83).Enabled=   0   'False
      Tab(1).Control(84)=   "Ingresso(26)"
      Tab(1).Control(84).Enabled=   0   'False
      Tab(1).Control(85)=   "Ingresso(27)"
      Tab(1).Control(85).Enabled=   0   'False
      Tab(1).Control(86)=   "Ingresso(28)"
      Tab(1).Control(86).Enabled=   0   'False
      Tab(1).Control(87)=   "Ingresso(29)"
      Tab(1).Control(87).Enabled=   0   'False
      Tab(1).Control(88)=   "Ingresso(30)"
      Tab(1).Control(88).Enabled=   0   'False
      Tab(1).Control(89)=   "Ingresso(31)"
      Tab(1).Control(89).Enabled=   0   'False
      Tab(1).Control(90)=   "Ingresso(32)"
      Tab(1).Control(90).Enabled=   0   'False
      Tab(1).Control(91)=   "Ingresso(33)"
      Tab(1).Control(91).Enabled=   0   'False
      Tab(1).Control(92)=   "Ingresso(34)"
      Tab(1).Control(92).Enabled=   0   'False
      Tab(1).Control(93)=   "Ingresso(35)"
      Tab(1).Control(93).Enabled=   0   'False
      Tab(1).Control(94)=   "Ingresso(36)"
      Tab(1).Control(94).Enabled=   0   'False
      Tab(1).Control(95)=   "Ingresso(37)"
      Tab(1).Control(95).Enabled=   0   'False
      Tab(1).Control(96)=   "Ingresso(38)"
      Tab(1).Control(96).Enabled=   0   'False
      Tab(1).Control(97)=   "Ingresso(39)"
      Tab(1).Control(97).Enabled=   0   'False
      Tab(1).Control(98)=   "Ingresso(40)"
      Tab(1).Control(98).Enabled=   0   'False
      Tab(1).Control(99)=   "Ingresso(41)"
      Tab(1).Control(99).Enabled=   0   'False
      Tab(1).Control(100)=   "Ingresso(42)"
      Tab(1).Control(100).Enabled=   0   'False
      Tab(1).Control(101)=   "Ingresso(43)"
      Tab(1).Control(101).Enabled=   0   'False
      Tab(1).Control(102)=   "Ingresso(44)"
      Tab(1).Control(102).Enabled=   0   'False
      Tab(1).Control(103)=   "Ingresso(45)"
      Tab(1).Control(103).Enabled=   0   'False
      Tab(1).Control(104)=   "Ingresso(46)"
      Tab(1).Control(104).Enabled=   0   'False
      Tab(1).Control(105)=   "Ingresso(47)"
      Tab(1).Control(105).Enabled=   0   'False
      Tab(1).Control(106)=   "Ingresso(48)"
      Tab(1).Control(106).Enabled=   0   'False
      Tab(1).Control(107)=   "Ingresso(49)"
      Tab(1).Control(107).Enabled=   0   'False
      Tab(1).Control(108)=   "Ingresso(50)"
      Tab(1).Control(108).Enabled=   0   'False
      Tab(1).Control(109)=   "Ingresso(51)"
      Tab(1).Control(109).Enabled=   0   'False
      Tab(1).Control(110)=   "Ingresso(52)"
      Tab(1).Control(110).Enabled=   0   'False
      Tab(1).Control(111)=   "Ingresso(53)"
      Tab(1).Control(111).Enabled=   0   'False
      Tab(1).Control(112)=   "Ingresso(54)"
      Tab(1).Control(112).Enabled=   0   'False
      Tab(1).Control(113)=   "Ingresso(55)"
      Tab(1).Control(113).Enabled=   0   'False
      Tab(1).Control(114)=   "Ingresso(56)"
      Tab(1).Control(114).Enabled=   0   'False
      Tab(1).Control(115)=   "Ingresso(57)"
      Tab(1).Control(115).Enabled=   0   'False
      Tab(1).Control(116)=   "Ingresso(58)"
      Tab(1).Control(116).Enabled=   0   'False
      Tab(1).Control(117)=   "Ingresso(59)"
      Tab(1).Control(117).Enabled=   0   'False
      Tab(1).Control(118)=   "Ingresso(60)"
      Tab(1).Control(118).Enabled=   0   'False
      Tab(1).Control(119)=   "Ingresso(61)"
      Tab(1).Control(119).Enabled=   0   'False
      Tab(1).Control(120)=   "Ingresso(62)"
      Tab(1).Control(120).Enabled=   0   'False
      Tab(1).Control(121)=   "Ingresso(63)"
      Tab(1).Control(121).Enabled=   0   'False
      Tab(1).Control(122)=   "Ingresso(64)"
      Tab(1).Control(122).Enabled=   0   'False
      Tab(1).Control(123)=   "Ingresso(65)"
      Tab(1).Control(123).Enabled=   0   'False
      Tab(1).Control(124)=   "Ingresso(66)"
      Tab(1).Control(124).Enabled=   0   'False
      Tab(1).Control(125)=   "Ingresso(67)"
      Tab(1).Control(125).Enabled=   0   'False
      Tab(1).Control(126)=   "Ingresso(68)"
      Tab(1).Control(126).Enabled=   0   'False
      Tab(1).Control(127)=   "Ingresso(69)"
      Tab(1).Control(127).Enabled=   0   'False
      Tab(1).Control(128)=   "Ingresso(70)"
      Tab(1).Control(128).Enabled=   0   'False
      Tab(1).Control(129)=   "Ingresso(71)"
      Tab(1).Control(129).Enabled=   0   'False
      Tab(1).Control(130)=   "Ingresso(72)"
      Tab(1).Control(130).Enabled=   0   'False
      Tab(1).Control(131)=   "Ingresso(73)"
      Tab(1).Control(131).Enabled=   0   'False
      Tab(1).Control(132)=   "Ingresso(74)"
      Tab(1).Control(132).Enabled=   0   'False
      Tab(1).Control(133)=   "Ingresso(75)"
      Tab(1).Control(133).Enabled=   0   'False
      Tab(1).Control(134)=   "Ingresso(76)"
      Tab(1).Control(134).Enabled=   0   'False
      Tab(1).Control(135)=   "Ingresso(77)"
      Tab(1).Control(135).Enabled=   0   'False
      Tab(1).Control(136)=   "Ingresso(78)"
      Tab(1).Control(136).Enabled=   0   'False
      Tab(1).Control(137)=   "Ingresso(79)"
      Tab(1).Control(137).Enabled=   0   'False
      Tab(1).Control(138)=   "Ingresso(80)"
      Tab(1).Control(138).Enabled=   0   'False
      Tab(1).Control(139)=   "Ingresso(81)"
      Tab(1).Control(139).Enabled=   0   'False
      Tab(1).Control(140)=   "Ingresso(82)"
      Tab(1).Control(140).Enabled=   0   'False
      Tab(1).Control(141)=   "Ingresso(83)"
      Tab(1).Control(141).Enabled=   0   'False
      Tab(1).Control(142)=   "Ingresso(84)"
      Tab(1).Control(142).Enabled=   0   'False
      Tab(1).Control(143)=   "Ingresso(85)"
      Tab(1).Control(143).Enabled=   0   'False
      Tab(1).Control(144)=   "Ingresso(86)"
      Tab(1).Control(144).Enabled=   0   'False
      Tab(1).Control(145)=   "Ingresso(87)"
      Tab(1).Control(145).Enabled=   0   'False
      Tab(1).Control(146)=   "Ingresso(88)"
      Tab(1).Control(146).Enabled=   0   'False
      Tab(1).Control(147)=   "Ingresso(89)"
      Tab(1).Control(147).Enabled=   0   'False
      Tab(1).Control(148)=   "Ingresso(90)"
      Tab(1).Control(148).Enabled=   0   'False
      Tab(1).Control(149)=   "Ingresso(91)"
      Tab(1).Control(149).Enabled=   0   'False
      Tab(1).Control(150)=   "Ingresso(92)"
      Tab(1).Control(150).Enabled=   0   'False
      Tab(1).Control(151)=   "Ingresso(93)"
      Tab(1).Control(151).Enabled=   0   'False
      Tab(1).Control(152)=   "Ingresso(94)"
      Tab(1).Control(152).Enabled=   0   'False
      Tab(1).Control(153)=   "Ingresso(95)"
      Tab(1).Control(153).Enabled=   0   'False
      Tab(1).Control(154)=   "Ingresso(96)"
      Tab(1).Control(154).Enabled=   0   'False
      Tab(1).Control(155)=   "Ingresso(97)"
      Tab(1).Control(155).Enabled=   0   'False
      Tab(1).Control(156)=   "Ingresso(98)"
      Tab(1).Control(156).Enabled=   0   'False
      Tab(1).Control(157)=   "Ingresso(99)"
      Tab(1).Control(157).Enabled=   0   'False
      Tab(1).Control(158)=   "Ingresso(100)"
      Tab(1).Control(158).Enabled=   0   'False
      Tab(1).Control(159)=   "Ingresso(101)"
      Tab(1).Control(159).Enabled=   0   'False
      Tab(1).Control(160)=   "Ingresso(102)"
      Tab(1).Control(160).Enabled=   0   'False
      Tab(1).Control(161)=   "Ingresso(103)"
      Tab(1).Control(161).Enabled=   0   'False
      Tab(1).Control(162)=   "Ingresso(104)"
      Tab(1).Control(162).Enabled=   0   'False
      Tab(1).Control(163)=   "Ingresso(105)"
      Tab(1).Control(163).Enabled=   0   'False
      Tab(1).Control(164)=   "Ingresso(106)"
      Tab(1).Control(164).Enabled=   0   'False
      Tab(1).Control(165)=   "Ingresso(107)"
      Tab(1).Control(165).Enabled=   0   'False
      Tab(1).Control(166)=   "Ingresso(108)"
      Tab(1).Control(166).Enabled=   0   'False
      Tab(1).Control(167)=   "Ingresso(109)"
      Tab(1).Control(167).Enabled=   0   'False
      Tab(1).Control(168)=   "Ingresso(110)"
      Tab(1).Control(168).Enabled=   0   'False
      Tab(1).Control(169)=   "Ingresso(111)"
      Tab(1).Control(169).Enabled=   0   'False
      Tab(1).Control(170)=   "Ingresso(112)"
      Tab(1).Control(170).Enabled=   0   'False
      Tab(1).Control(171)=   "Ingresso(113)"
      Tab(1).Control(171).Enabled=   0   'False
      Tab(1).Control(172)=   "Ingresso(114)"
      Tab(1).Control(172).Enabled=   0   'False
      Tab(1).Control(173)=   "Ingresso(115)"
      Tab(1).Control(173).Enabled=   0   'False
      Tab(1).Control(174)=   "Ingresso(116)"
      Tab(1).Control(174).Enabled=   0   'False
      Tab(1).Control(175)=   "Ingresso(117)"
      Tab(1).Control(175).Enabled=   0   'False
      Tab(1).Control(176)=   "Ingresso(118)"
      Tab(1).Control(176).Enabled=   0   'False
      Tab(1).Control(177)=   "Ingresso(119)"
      Tab(1).Control(177).Enabled=   0   'False
      Tab(1).Control(178)=   "Ingresso(120)"
      Tab(1).Control(178).Enabled=   0   'False
      Tab(1).Control(179)=   "Ingresso(121)"
      Tab(1).Control(179).Enabled=   0   'False
      Tab(1).Control(180)=   "Ingresso(122)"
      Tab(1).Control(180).Enabled=   0   'False
      Tab(1).Control(181)=   "Ingresso(123)"
      Tab(1).Control(181).Enabled=   0   'False
      Tab(1).Control(182)=   "Ingresso(124)"
      Tab(1).Control(182).Enabled=   0   'False
      Tab(1).Control(183)=   "Ingresso(125)"
      Tab(1).Control(183).Enabled=   0   'False
      Tab(1).Control(184)=   "Ingresso(126)"
      Tab(1).Control(184).Enabled=   0   'False
      Tab(1).Control(185)=   "Ingresso(127)"
      Tab(1).Control(185).Enabled=   0   'False
      Tab(1).Control(186)=   "Ingresso(128)"
      Tab(1).Control(186).Enabled=   0   'False
      Tab(1).Control(187)=   "Ingresso(129)"
      Tab(1).Control(187).Enabled=   0   'False
      Tab(1).Control(188)=   "Ingresso(130)"
      Tab(1).Control(188).Enabled=   0   'False
      Tab(1).Control(189)=   "Ingresso(131)"
      Tab(1).Control(189).Enabled=   0   'False
      Tab(1).Control(190)=   "Ingresso(132)"
      Tab(1).Control(190).Enabled=   0   'False
      Tab(1).Control(191)=   "Ingresso(133)"
      Tab(1).Control(191).Enabled=   0   'False
      Tab(1).Control(192)=   "Ingresso(134)"
      Tab(1).Control(192).Enabled=   0   'False
      Tab(1).Control(193)=   "Ingresso(135)"
      Tab(1).Control(193).Enabled=   0   'False
      Tab(1).Control(194)=   "Ingresso(136)"
      Tab(1).Control(194).Enabled=   0   'False
      Tab(1).Control(195)=   "Ingresso(137)"
      Tab(1).Control(195).Enabled=   0   'False
      Tab(1).Control(196)=   "Ingresso(138)"
      Tab(1).Control(196).Enabled=   0   'False
      Tab(1).Control(197)=   "Ingresso(139)"
      Tab(1).Control(197).Enabled=   0   'False
      Tab(1).Control(198)=   "Ingresso(140)"
      Tab(1).Control(198).Enabled=   0   'False
      Tab(1).Control(199)=   "Ingresso(141)"
      Tab(1).Control(199).Enabled=   0   'False
      Tab(1).Control(200)=   "Ingresso(142)"
      Tab(1).Control(200).Enabled=   0   'False
      Tab(1).Control(201)=   "Ingresso(143)"
      Tab(1).Control(201).Enabled=   0   'False
      Tab(1).Control(202)=   "Ingresso(144)"
      Tab(1).Control(202).Enabled=   0   'False
      Tab(1).Control(203)=   "Ingresso(145)"
      Tab(1).Control(203).Enabled=   0   'False
      Tab(1).Control(204)=   "Ingresso(146)"
      Tab(1).Control(204).Enabled=   0   'False
      Tab(1).Control(205)=   "Ingresso(147)"
      Tab(1).Control(205).Enabled=   0   'False
      Tab(1).Control(206)=   "Ingresso(148)"
      Tab(1).Control(206).Enabled=   0   'False
      Tab(1).Control(207)=   "Ingresso(149)"
      Tab(1).Control(207).Enabled=   0   'False
      Tab(1).Control(208)=   "Ingresso(150)"
      Tab(1).Control(208).Enabled=   0   'False
      Tab(1).Control(209)=   "Ingresso(151)"
      Tab(1).Control(209).Enabled=   0   'False
      Tab(1).Control(210)=   "Ingresso(152)"
      Tab(1).Control(210).Enabled=   0   'False
      Tab(1).Control(211)=   "Ingresso(153)"
      Tab(1).Control(211).Enabled=   0   'False
      Tab(1).Control(212)=   "Ingresso(154)"
      Tab(1).Control(212).Enabled=   0   'False
      Tab(1).Control(213)=   "Ingresso(155)"
      Tab(1).Control(213).Enabled=   0   'False
      Tab(1).Control(214)=   "Ingresso(156)"
      Tab(1).Control(214).Enabled=   0   'False
      Tab(1).Control(215)=   "Ingresso(157)"
      Tab(1).Control(215).Enabled=   0   'False
      Tab(1).Control(216)=   "Ingresso(158)"
      Tab(1).Control(216).Enabled=   0   'False
      Tab(1).Control(217)=   "Ingresso(159)"
      Tab(1).Control(217).Enabled=   0   'False
      Tab(1).Control(218)=   "Ingresso(160)"
      Tab(1).Control(218).Enabled=   0   'False
      Tab(1).Control(219)=   "Ingresso(161)"
      Tab(1).Control(219).Enabled=   0   'False
      Tab(1).Control(220)=   "Ingresso(162)"
      Tab(1).Control(220).Enabled=   0   'False
      Tab(1).Control(221)=   "Ingresso(163)"
      Tab(1).Control(221).Enabled=   0   'False
      Tab(1).Control(222)=   "Ingresso(164)"
      Tab(1).Control(222).Enabled=   0   'False
      Tab(1).Control(223)=   "Ingresso(165)"
      Tab(1).Control(223).Enabled=   0   'False
      Tab(1).Control(224)=   "Ingresso(166)"
      Tab(1).Control(224).Enabled=   0   'False
      Tab(1).Control(225)=   "Ingresso(167)"
      Tab(1).Control(225).Enabled=   0   'False
      Tab(1).Control(226)=   "Ingresso(168)"
      Tab(1).Control(226).Enabled=   0   'False
      Tab(1).Control(227)=   "Ingresso(169)"
      Tab(1).Control(227).Enabled=   0   'False
      Tab(1).Control(228)=   "Ingresso(170)"
      Tab(1).Control(228).Enabled=   0   'False
      Tab(1).Control(229)=   "Ingresso(171)"
      Tab(1).Control(229).Enabled=   0   'False
      Tab(1).Control(230)=   "Ingresso(172)"
      Tab(1).Control(230).Enabled=   0   'False
      Tab(1).Control(231)=   "Ingresso(173)"
      Tab(1).Control(231).Enabled=   0   'False
      Tab(1).Control(232)=   "Ingresso(174)"
      Tab(1).Control(232).Enabled=   0   'False
      Tab(1).Control(233)=   "Ingresso(175)"
      Tab(1).Control(233).Enabled=   0   'False
      Tab(1).Control(234)=   "Ingresso(176)"
      Tab(1).Control(234).Enabled=   0   'False
      Tab(1).Control(235)=   "Ingresso(177)"
      Tab(1).Control(235).Enabled=   0   'False
      Tab(1).Control(236)=   "Ingresso(178)"
      Tab(1).Control(236).Enabled=   0   'False
      Tab(1).Control(237)=   "Ingresso(179)"
      Tab(1).Control(237).Enabled=   0   'False
      Tab(1).Control(238)=   "Ingresso(180)"
      Tab(1).Control(238).Enabled=   0   'False
      Tab(1).Control(239)=   "Ingresso(181)"
      Tab(1).Control(239).Enabled=   0   'False
      Tab(1).Control(240)=   "Ingresso(182)"
      Tab(1).Control(240).Enabled=   0   'False
      Tab(1).Control(241)=   "Ingresso(183)"
      Tab(1).Control(241).Enabled=   0   'False
      Tab(1).Control(242)=   "Ingresso(184)"
      Tab(1).Control(242).Enabled=   0   'False
      Tab(1).Control(243)=   "Ingresso(185)"
      Tab(1).Control(243).Enabled=   0   'False
      Tab(1).Control(244)=   "Ingresso(186)"
      Tab(1).Control(244).Enabled=   0   'False
      Tab(1).Control(245)=   "Ingresso(187)"
      Tab(1).Control(245).Enabled=   0   'False
      Tab(1).Control(246)=   "Ingresso(188)"
      Tab(1).Control(246).Enabled=   0   'False
      Tab(1).Control(247)=   "Ingresso(189)"
      Tab(1).Control(247).Enabled=   0   'False
      Tab(1).Control(248)=   "Ingresso(190)"
      Tab(1).Control(248).Enabled=   0   'False
      Tab(1).Control(249)=   "Ingresso(191)"
      Tab(1).Control(249).Enabled=   0   'False
      Tab(1).Control(250)=   "Ingresso(192)"
      Tab(1).Control(250).Enabled=   0   'False
      Tab(1).Control(251)=   "Ingresso(193)"
      Tab(1).Control(251).Enabled=   0   'False
      Tab(1).Control(252)=   "Ingresso(194)"
      Tab(1).Control(252).Enabled=   0   'False
      Tab(1).Control(253)=   "Ingresso(195)"
      Tab(1).Control(253).Enabled=   0   'False
      Tab(1).Control(254)=   "Ingresso(196)"
      Tab(1).Control(254).Enabled=   0   'False
      Tab(1).Control(255)=   "Ingresso(197)"
      Tab(1).Control(255).Enabled=   0   'False
      Tab(1).Control(256)=   "Ingresso(198)"
      Tab(1).Control(256).Enabled=   0   'False
      Tab(1).Control(257)=   "Ingresso(199)"
      Tab(1).Control(257).Enabled=   0   'False
      Tab(1).Control(258)=   "Ingresso(200)"
      Tab(1).Control(258).Enabled=   0   'False
      Tab(1).Control(259)=   "Ingresso(201)"
      Tab(1).Control(259).Enabled=   0   'False
      Tab(1).Control(260)=   "Ingresso(202)"
      Tab(1).Control(260).Enabled=   0   'False
      Tab(1).Control(261)=   "Ingresso(203)"
      Tab(1).Control(261).Enabled=   0   'False
      Tab(1).Control(262)=   "Ingresso(204)"
      Tab(1).Control(262).Enabled=   0   'False
      Tab(1).Control(263)=   "Ingresso(205)"
      Tab(1).Control(263).Enabled=   0   'False
      Tab(1).Control(264)=   "Ingresso(206)"
      Tab(1).Control(264).Enabled=   0   'False
      Tab(1).Control(265)=   "Ingresso(207)"
      Tab(1).Control(265).Enabled=   0   'False
      Tab(1).Control(266)=   "Ingresso(208)"
      Tab(1).Control(266).Enabled=   0   'False
      Tab(1).Control(267)=   "Ingresso(209)"
      Tab(1).Control(267).Enabled=   0   'False
      Tab(1).Control(268)=   "Ingresso(210)"
      Tab(1).Control(268).Enabled=   0   'False
      Tab(1).Control(269)=   "Ingresso(211)"
      Tab(1).Control(269).Enabled=   0   'False
      Tab(1).Control(270)=   "Ingresso(212)"
      Tab(1).Control(270).Enabled=   0   'False
      Tab(1).Control(271)=   "Ingresso(213)"
      Tab(1).Control(271).Enabled=   0   'False
      Tab(1).Control(272)=   "Ingresso(214)"
      Tab(1).Control(272).Enabled=   0   'False
      Tab(1).Control(273)=   "Ingresso(215)"
      Tab(1).Control(273).Enabled=   0   'False
      Tab(1).Control(274)=   "Ingresso(216)"
      Tab(1).Control(274).Enabled=   0   'False
      Tab(1).Control(275)=   "Ingresso(217)"
      Tab(1).Control(275).Enabled=   0   'False
      Tab(1).Control(276)=   "Ingresso(218)"
      Tab(1).Control(276).Enabled=   0   'False
      Tab(1).Control(277)=   "Ingresso(219)"
      Tab(1).Control(277).Enabled=   0   'False
      Tab(1).Control(278)=   "Ingresso(220)"
      Tab(1).Control(278).Enabled=   0   'False
      Tab(1).Control(279)=   "Ingresso(221)"
      Tab(1).Control(279).Enabled=   0   'False
      Tab(1).Control(280)=   "Ingresso(222)"
      Tab(1).Control(280).Enabled=   0   'False
      Tab(1).Control(281)=   "Ingresso(223)"
      Tab(1).Control(281).Enabled=   0   'False
      Tab(1).ControlCount=   282
      TabCaption(2)   =   "Entry"
      TabPicture(2)   =   "frmTestIO.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Threading"
      TabPicture(3)   =   "frmTestIO.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Walkingbeam"
      TabPicture(4)   =   "frmTestIO.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Storage"
      TabPicture(5)   =   "frmTestIO.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   223
         Left            =   14220
         TabIndex        =   224
         Top             =   6420
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   222
         Left            =   14220
         TabIndex        =   223
         Top             =   6780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   221
         Left            =   14220
         TabIndex        =   222
         Top             =   7110
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   220
         Left            =   14220
         TabIndex        =   221
         Top             =   7440
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   219
         Left            =   13140
         TabIndex        =   220
         Top             =   6420
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   218
         Left            =   13140
         TabIndex        =   219
         Top             =   6780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   217
         Left            =   13140
         TabIndex        =   218
         Top             =   7110
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   216
         Left            =   13140
         TabIndex        =   217
         Top             =   7440
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   215
         Left            =   12060
         TabIndex        =   216
         Top             =   6420
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   214
         Left            =   12060
         TabIndex        =   215
         Top             =   6780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   213
         Left            =   12060
         TabIndex        =   214
         Top             =   7110
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   212
         Left            =   12060
         TabIndex        =   213
         Top             =   7440
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   211
         Left            =   10980
         TabIndex        =   212
         Top             =   6420
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   210
         Left            =   10980
         TabIndex        =   211
         Top             =   6780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   209
         Left            =   10980
         TabIndex        =   210
         Top             =   7110
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   208
         Left            =   10980
         TabIndex        =   209
         Top             =   7440
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   207
         Left            =   9900
         TabIndex        =   208
         Top             =   6420
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   206
         Left            =   9900
         TabIndex        =   207
         Top             =   6780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   205
         Left            =   9900
         TabIndex        =   206
         Top             =   7110
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   204
         Left            =   9900
         TabIndex        =   205
         Top             =   7440
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   203
         Left            =   8820
         TabIndex        =   204
         Top             =   6420
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   202
         Left            =   8820
         TabIndex        =   203
         Top             =   6780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   201
         Left            =   8820
         TabIndex        =   202
         Top             =   7110
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   200
         Left            =   8820
         TabIndex        =   201
         Top             =   7440
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   199
         Left            =   7740
         TabIndex        =   200
         Top             =   6420
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   198
         Left            =   7740
         TabIndex        =   199
         Top             =   6780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   197
         Left            =   7740
         TabIndex        =   198
         Top             =   7110
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   196
         Left            =   7740
         TabIndex        =   197
         Top             =   7440
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   195
         Left            =   6660
         TabIndex        =   196
         Top             =   6420
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   194
         Left            =   6660
         TabIndex        =   195
         Top             =   6780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   193
         Left            =   6660
         TabIndex        =   194
         Top             =   7110
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   192
         Left            =   6660
         TabIndex        =   193
         Top             =   7440
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   191
         Left            =   5580
         TabIndex        =   192
         Top             =   6420
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   190
         Left            =   5580
         TabIndex        =   191
         Top             =   6780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   189
         Left            =   5580
         TabIndex        =   190
         Top             =   7110
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   188
         Left            =   5580
         TabIndex        =   189
         Top             =   7440
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   187
         Left            =   4500
         TabIndex        =   188
         Top             =   6420
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   186
         Left            =   4500
         TabIndex        =   187
         Top             =   6780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   185
         Left            =   4500
         TabIndex        =   186
         Top             =   7110
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   184
         Left            =   4500
         TabIndex        =   185
         Top             =   7440
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   183
         Left            =   3420
         TabIndex        =   184
         Top             =   6420
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   182
         Left            =   3420
         TabIndex        =   183
         Top             =   6780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   181
         Left            =   3420
         TabIndex        =   182
         Top             =   7110
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   180
         Left            =   3420
         TabIndex        =   181
         Top             =   7440
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   179
         Left            =   2340
         TabIndex        =   180
         Top             =   6420
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   178
         Left            =   2340
         TabIndex        =   179
         Top             =   6780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   177
         Left            =   2340
         TabIndex        =   178
         Top             =   7110
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   176
         Left            =   2340
         TabIndex        =   177
         Top             =   7440
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   175
         Left            =   1260
         TabIndex        =   176
         Top             =   6420
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   174
         Left            =   1260
         TabIndex        =   175
         Top             =   6780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   173
         Left            =   1260
         TabIndex        =   174
         Top             =   7110
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   172
         Left            =   1260
         TabIndex        =   173
         Top             =   7440
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   171
         Left            =   180
         TabIndex        =   172
         Top             =   6420
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   170
         Left            =   180
         TabIndex        =   171
         Top             =   6780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   169
         Left            =   180
         TabIndex        =   170
         Top             =   7110
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   168
         Left            =   180
         TabIndex        =   169
         Top             =   7440
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   167
         Left            =   14220
         TabIndex        =   168
         Top             =   4800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   166
         Left            =   14220
         TabIndex        =   167
         Top             =   5160
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   165
         Left            =   14220
         TabIndex        =   166
         Top             =   5490
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   164
         Left            =   14220
         TabIndex        =   165
         Top             =   5820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   163
         Left            =   13140
         TabIndex        =   164
         Top             =   4800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   162
         Left            =   13140
         TabIndex        =   163
         Top             =   5160
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   161
         Left            =   13140
         TabIndex        =   162
         Top             =   5490
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   160
         Left            =   13140
         TabIndex        =   161
         Top             =   5820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   159
         Left            =   12060
         TabIndex        =   160
         Top             =   4800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   158
         Left            =   12060
         TabIndex        =   159
         Top             =   5160
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   157
         Left            =   12060
         TabIndex        =   158
         Top             =   5490
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   156
         Left            =   12060
         TabIndex        =   157
         Top             =   5820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   155
         Left            =   10980
         TabIndex        =   156
         Top             =   4800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   154
         Left            =   10980
         TabIndex        =   155
         Top             =   5160
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   153
         Left            =   10980
         TabIndex        =   154
         Top             =   5490
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   152
         Left            =   10980
         TabIndex        =   153
         Top             =   5820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   151
         Left            =   9900
         TabIndex        =   152
         Top             =   4800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   150
         Left            =   9900
         TabIndex        =   151
         Top             =   5160
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   149
         Left            =   9900
         TabIndex        =   150
         Top             =   5490
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   148
         Left            =   9900
         TabIndex        =   149
         Top             =   5820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   147
         Left            =   8820
         TabIndex        =   148
         Top             =   4800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   146
         Left            =   8820
         TabIndex        =   147
         Top             =   5160
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   145
         Left            =   8820
         TabIndex        =   146
         Top             =   5490
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   144
         Left            =   8820
         TabIndex        =   145
         Top             =   5820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   143
         Left            =   7740
         TabIndex        =   144
         Top             =   4800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   142
         Left            =   7740
         TabIndex        =   143
         Top             =   5160
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   141
         Left            =   7740
         TabIndex        =   142
         Top             =   5490
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   140
         Left            =   7740
         TabIndex        =   141
         Top             =   5820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   139
         Left            =   6660
         TabIndex        =   140
         Top             =   4800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   138
         Left            =   6660
         TabIndex        =   139
         Top             =   5160
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   137
         Left            =   6660
         TabIndex        =   138
         Top             =   5490
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   136
         Left            =   6660
         TabIndex        =   137
         Top             =   5820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   135
         Left            =   5580
         TabIndex        =   136
         Top             =   4800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   134
         Left            =   5580
         TabIndex        =   135
         Top             =   5160
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   133
         Left            =   5580
         TabIndex        =   134
         Top             =   5490
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   132
         Left            =   5580
         TabIndex        =   133
         Top             =   5820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   131
         Left            =   4500
         TabIndex        =   132
         Top             =   4800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   130
         Left            =   4500
         TabIndex        =   131
         Top             =   5160
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   129
         Left            =   4500
         TabIndex        =   130
         Top             =   5490
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   128
         Left            =   4500
         TabIndex        =   129
         Top             =   5820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   127
         Left            =   3420
         TabIndex        =   128
         Top             =   4800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   126
         Left            =   3420
         TabIndex        =   127
         Top             =   5160
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   125
         Left            =   3420
         TabIndex        =   126
         Top             =   5490
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   124
         Left            =   3420
         TabIndex        =   125
         Top             =   5820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   123
         Left            =   2340
         TabIndex        =   124
         Top             =   4800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   122
         Left            =   2340
         TabIndex        =   123
         Top             =   5160
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   121
         Left            =   2340
         TabIndex        =   122
         Top             =   5490
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   120
         Left            =   2340
         TabIndex        =   121
         Top             =   5820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   119
         Left            =   1260
         TabIndex        =   120
         Top             =   4800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   118
         Left            =   1260
         TabIndex        =   119
         Top             =   5160
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   117
         Left            =   1260
         TabIndex        =   118
         Top             =   5490
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   116
         Left            =   1260
         TabIndex        =   117
         Top             =   5820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   115
         Left            =   180
         TabIndex        =   116
         Top             =   4800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   114
         Left            =   180
         TabIndex        =   115
         Top             =   5160
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   113
         Left            =   180
         TabIndex        =   114
         Top             =   5490
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   112
         Left            =   180
         TabIndex        =   113
         Top             =   5820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   111
         Left            =   14250
         TabIndex        =   112
         Top             =   2460
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   110
         Left            =   14250
         TabIndex        =   111
         Top             =   2820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   109
         Left            =   14250
         TabIndex        =   110
         Top             =   3150
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   108
         Left            =   14250
         TabIndex        =   109
         Top             =   3480
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   107
         Left            =   13170
         TabIndex        =   108
         Top             =   2460
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   106
         Left            =   13170
         TabIndex        =   107
         Top             =   2820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   105
         Left            =   13170
         TabIndex        =   106
         Top             =   3150
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   104
         Left            =   13170
         TabIndex        =   105
         Top             =   3480
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   103
         Left            =   12090
         TabIndex        =   104
         Top             =   2460
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   102
         Left            =   12090
         TabIndex        =   103
         Top             =   2820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   101
         Left            =   12090
         TabIndex        =   102
         Top             =   3150
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   100
         Left            =   12090
         TabIndex        =   101
         Top             =   3480
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   99
         Left            =   11010
         TabIndex        =   100
         Top             =   2460
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   98
         Left            =   11010
         TabIndex        =   99
         Top             =   2820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   97
         Left            =   11010
         TabIndex        =   98
         Top             =   3150
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   96
         Left            =   11010
         TabIndex        =   97
         Top             =   3480
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   95
         Left            =   9930
         TabIndex        =   96
         Top             =   2460
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   94
         Left            =   9930
         TabIndex        =   95
         Top             =   2820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   93
         Left            =   9930
         TabIndex        =   94
         Top             =   3150
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   92
         Left            =   9930
         TabIndex        =   93
         Top             =   3480
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   91
         Left            =   8850
         TabIndex        =   92
         Top             =   2460
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   90
         Left            =   8850
         TabIndex        =   91
         Top             =   2820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   89
         Left            =   8850
         TabIndex        =   90
         Top             =   3150
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   88
         Left            =   8850
         TabIndex        =   89
         Top             =   3480
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   87
         Left            =   7770
         TabIndex        =   88
         Top             =   2460
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   86
         Left            =   7770
         TabIndex        =   87
         Top             =   2820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   85
         Left            =   7770
         TabIndex        =   86
         Top             =   3150
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   84
         Left            =   7770
         TabIndex        =   85
         Top             =   3480
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   83
         Left            =   6690
         TabIndex        =   84
         Top             =   2460
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   82
         Left            =   6690
         TabIndex        =   83
         Top             =   2820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   81
         Left            =   6690
         TabIndex        =   82
         Top             =   3150
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   80
         Left            =   6690
         TabIndex        =   81
         Top             =   3480
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   79
         Left            =   5610
         TabIndex        =   80
         Top             =   2460
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   78
         Left            =   5610
         TabIndex        =   79
         Top             =   2820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   77
         Left            =   5610
         TabIndex        =   78
         Top             =   3150
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   76
         Left            =   5610
         TabIndex        =   77
         Top             =   3480
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   75
         Left            =   4530
         TabIndex        =   76
         Top             =   2460
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   74
         Left            =   4530
         TabIndex        =   75
         Top             =   2820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   73
         Left            =   4530
         TabIndex        =   74
         Top             =   3150
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   72
         Left            =   4530
         TabIndex        =   73
         Top             =   3480
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   71
         Left            =   3450
         TabIndex        =   72
         Top             =   2460
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   70
         Left            =   3450
         TabIndex        =   71
         Top             =   2820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   69
         Left            =   3450
         TabIndex        =   70
         Top             =   3150
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   68
         Left            =   3450
         TabIndex        =   69
         Top             =   3480
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   67
         Left            =   2370
         TabIndex        =   68
         Top             =   2460
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   66
         Left            =   2370
         TabIndex        =   67
         Top             =   2820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   65
         Left            =   2370
         TabIndex        =   66
         Top             =   3150
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   64
         Left            =   2370
         TabIndex        =   65
         Top             =   3480
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   63
         Left            =   1290
         TabIndex        =   64
         Top             =   2460
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   62
         Left            =   1290
         TabIndex        =   63
         Top             =   2820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   61
         Left            =   1290
         TabIndex        =   62
         Top             =   3150
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   60
         Left            =   1290
         TabIndex        =   61
         Top             =   3480
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   59
         Left            =   210
         TabIndex        =   60
         Top             =   2460
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   58
         Left            =   210
         TabIndex        =   59
         Top             =   2820
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   57
         Left            =   210
         TabIndex        =   58
         Top             =   3150
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   56
         Left            =   210
         TabIndex        =   57
         Top             =   3480
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   55
         Left            =   14250
         TabIndex        =   56
         Top             =   780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   54
         Left            =   14250
         TabIndex        =   55
         Top             =   1140
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   53
         Left            =   14250
         TabIndex        =   54
         Top             =   1470
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   52
         Left            =   14250
         TabIndex        =   53
         Top             =   1800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   51
         Left            =   13170
         TabIndex        =   52
         Top             =   780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   50
         Left            =   13170
         TabIndex        =   51
         Top             =   1140
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   49
         Left            =   13170
         TabIndex        =   50
         Top             =   1470
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   48
         Left            =   13170
         TabIndex        =   49
         Top             =   1800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   47
         Left            =   12090
         TabIndex        =   48
         Top             =   780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   46
         Left            =   12090
         TabIndex        =   47
         Top             =   1140
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   45
         Left            =   12090
         TabIndex        =   46
         Top             =   1470
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   44
         Left            =   12090
         TabIndex        =   45
         Top             =   1800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   43
         Left            =   11010
         TabIndex        =   44
         Top             =   780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   42
         Left            =   11010
         TabIndex        =   43
         Top             =   1140
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   41
         Left            =   11010
         TabIndex        =   42
         Top             =   1470
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   40
         Left            =   11010
         TabIndex        =   41
         Top             =   1800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   39
         Left            =   9930
         TabIndex        =   40
         Top             =   780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   38
         Left            =   9930
         TabIndex        =   39
         Top             =   1140
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   37
         Left            =   9930
         TabIndex        =   38
         Top             =   1470
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   36
         Left            =   9930
         TabIndex        =   37
         Top             =   1800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   35
         Left            =   8850
         TabIndex        =   36
         Top             =   780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   34
         Left            =   8850
         TabIndex        =   35
         Top             =   1140
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   33
         Left            =   8850
         TabIndex        =   34
         Top             =   1470
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   32
         Left            =   8850
         TabIndex        =   33
         Top             =   1800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   31
         Left            =   7770
         TabIndex        =   32
         Top             =   780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   30
         Left            =   7770
         TabIndex        =   31
         Top             =   1140
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   29
         Left            =   7770
         TabIndex        =   30
         Top             =   1470
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   28
         Left            =   7770
         TabIndex        =   29
         Top             =   1800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   27
         Left            =   6690
         TabIndex        =   28
         Top             =   780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   26
         Left            =   6690
         TabIndex        =   27
         Top             =   1140
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   25
         Left            =   6690
         TabIndex        =   26
         Top             =   1470
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   24
         Left            =   6690
         TabIndex        =   25
         Top             =   1800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   23
         Left            =   5610
         TabIndex        =   24
         Top             =   780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   22
         Left            =   5610
         TabIndex        =   23
         Top             =   1140
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   21
         Left            =   5610
         TabIndex        =   22
         Top             =   1470
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   20
         Left            =   5610
         TabIndex        =   21
         Top             =   1800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   19
         Left            =   4530
         TabIndex        =   20
         Top             =   780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   18
         Left            =   4530
         TabIndex        =   19
         Top             =   1140
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   17
         Left            =   4530
         TabIndex        =   18
         Top             =   1470
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   16
         Left            =   4530
         TabIndex        =   17
         Top             =   1800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   15
         Left            =   3450
         TabIndex        =   16
         Top             =   780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   14
         Left            =   3450
         TabIndex        =   15
         Top             =   1140
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   13
         Left            =   3450
         TabIndex        =   14
         Top             =   1470
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   12
         Left            =   3450
         TabIndex        =   13
         Top             =   1800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   11
         Left            =   2370
         TabIndex        =   12
         Top             =   780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   10
         Left            =   2370
         TabIndex        =   11
         Top             =   1140
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   9
         Left            =   2370
         TabIndex        =   10
         Top             =   1470
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   8
         Left            =   2370
         TabIndex        =   9
         Top             =   1800
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   7
         Left            =   1290
         TabIndex        =   8
         Top             =   780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   6
         Left            =   1290
         TabIndex        =   7
         Top             =   1140
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   5
         Left            =   1290
         TabIndex        =   6
         Top             =   1470
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   4
         Left            =   1320
         TabIndex        =   5
         Top             =   1800
         Width           =   885
      End
      Begin VB.Timer TimerLocale 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   9120
         Top             =   1020
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.3"
         Height          =   465
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   780
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.2"
         Height          =   465
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1140
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.1"
         Height          =   465
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   1470
         Width           =   885
      End
      Begin VB.CheckBox Ingresso 
         Caption         =   "E5.0"
         Height          =   465
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   1800
         Width           =   885
      End
      Begin VB.Frame FrameLav 
         Height          =   7755
         Left            =   120
         TabIndex        =   225
         Top             =   750
         Width           =   15105
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "BLOW"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Index           =   10
            Left            =   8610
            TabIndex        =   236
            Top             =   1770
            Width           =   795
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "A18.3      locking top"
            Height          =   675
            Index           =   6
            Left            =   2880
            TabIndex        =   235
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "E8.1  nozzle backward"
            Height          =   285
            Index           =   9
            Left            =   510
            TabIndex        =   234
            Top             =   3480
            Width           =   2655
         End
         Begin VB.Shape Shape3 
            FillStyle       =   0  'Solid
            Height          =   195
            Index           =   2
            Left            =   4440
            Shape           =   3  'Circle
            Top             =   1560
            Width           =   165
         End
         Begin VB.Shape Shape3 
            FillStyle       =   0  'Solid
            Height          =   195
            Index           =   1
            Left            =   4440
            Shape           =   3  'Circle
            Top             =   1740
            Width           =   165
         End
         Begin VB.Shape Shape3 
            FillStyle       =   0  'Solid
            Height          =   195
            Index           =   0
            Left            =   6630
            Shape           =   3  'Circle
            Top             =   2070
            Width           =   165
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "A19.2      nozzle   foward"
            Height          =   675
            Index           =   8
            Left            =   5490
            TabIndex        =   233
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "A20.2   Blowing enable"
            Height          =   675
            Index           =   7
            Left            =   7680
            TabIndex        =   232
            Top             =   1710
            Width           =   795
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            Height          =   705
            Left            =   8490
            Shape           =   3  'Circle
            Top             =   1530
            Width           =   735
         End
         Begin VB.Line Line3 
            Index           =   4
            X1              =   7620
            X2              =   8550
            Y1              =   1890
            Y2              =   1890
         End
         Begin VB.Image Image1 
            Height          =   930
            Index           =   2
            Left            =   6750
            Picture         =   "frmTestIO.frx":00A8
            Top             =   1410
            Width           =   885
         End
         Begin VB.Line Line5 
            Index           =   3
            X1              =   6060
            X2              =   480
            Y1              =   3690
            Y2              =   3690
         End
         Begin VB.Line Line4 
            Index           =   3
            X1              =   6060
            X2              =   6060
            Y1              =   2160
            Y2              =   3690
         End
         Begin VB.Line Line3 
            Index           =   3
            X1              =   6060
            X2              =   6810
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "E7.2   Locking bottom"
            Height          =   255
            Index           =   5
            Left            =   450
            TabIndex        =   231
            Top             =   2850
            Width           =   2595
         End
         Begin VB.Line Line5 
            Index           =   2
            X1              =   4320
            X2              =   450
            Y1              =   3090
            Y2              =   3090
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "E7.1   Locking top"
            Height          =   255
            Index           =   4
            Left            =   450
            TabIndex        =   230
            Top             =   2520
            Width           =   2355
         End
         Begin VB.Line Line5 
            Index           =   1
            X1              =   4080
            X2              =   450
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "E5.4   mobile template top"
            Height          =   255
            Index           =   3
            Left            =   450
            TabIndex        =   229
            Top             =   2190
            Width           =   2235
         End
         Begin VB.Line Line5 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   2
            Index           =   0
            X1              =   3870
            X2              =   450
            Y1              =   2430
            Y2              =   2430
         End
         Begin VB.Line Line4 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   2
            Index           =   2
            X1              =   3870
            X2              =   3870
            Y1              =   1470
            Y2              =   2430
         End
         Begin VB.Line Line4 
            Index           =   1
            X1              =   4080
            X2              =   4080
            Y1              =   1650
            Y2              =   2760
         End
         Begin VB.Line Line4 
            Index           =   0
            X1              =   4320
            X2              =   4320
            Y1              =   1830
            Y2              =   3090
         End
         Begin VB.Image Image1 
            Height          =   930
            Index           =   0
            Left            =   4560
            Picture         =   "frmTestIO.frx":043E
            Top             =   1080
            Width           =   885
         End
         Begin VB.Line Line3 
            Index           =   2
            X1              =   4320
            X2              =   4620
            Y1              =   1830
            Y2              =   1830
         End
         Begin VB.Line Line3 
            Index           =   1
            X1              =   4080
            X2              =   4620
            Y1              =   1650
            Y2              =   1650
         End
         Begin VB.Line Line3 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   2
            Index           =   0
            X1              =   3870
            X2              =   4620
            Y1              =   1470
            Y2              =   1470
         End
         Begin VB.Image Image1 
            Height          =   930
            Index           =   1
            Left            =   1920
            Picture         =   "frmTestIO.frx":07D4
            Top             =   810
            Width           =   885
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "28.1  Blowing on"
            Height          =   255
            Index           =   2
            Left            =   450
            TabIndex        =   228
            Top             =   1380
            Width           =   1455
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   2
            Index           =   2
            X1              =   450
            X2              =   1890
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "28.2  Tube present"
            Height          =   435
            Index           =   1
            Left            =   450
            TabIndex        =   227
            Top             =   1050
            Width           =   1605
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "E7.0   Step"
            Height          =   255
            Index           =   0
            Left            =   450
            TabIndex        =   226
            Top             =   690
            Width           =   1245
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   2
            Index           =   0
            X1              =   450
            X2              =   1890
            Y1              =   930
            Y2              =   930
         End
         Begin VB.Line Line2 
            X1              =   2340
            X2              =   4590
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   2
            Index           =   1
            X1              =   420
            X2              =   2010
            Y1              =   1290
            Y2              =   1290
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   5340
            X2              =   6780
            Y1              =   1530
            Y2              =   1530
         End
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   55
         Left            =   14160
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   54
         Left            =   13080
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   53
         Left            =   12000
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   52
         Left            =   10920
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   51
         Left            =   9840
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   50
         Left            =   8760
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   49
         Left            =   7680
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   48
         Left            =   6600
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   47
         Left            =   5520
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   46
         Left            =   4440
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   45
         Left            =   3360
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   44
         Left            =   2280
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   43
         Left            =   1200
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   42
         Left            =   120
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   41
         Left            =   14160
         Top             =   4770
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   40
         Left            =   13080
         Top             =   4770
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   39
         Left            =   12000
         Top             =   4770
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   38
         Left            =   10920
         Top             =   4770
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   37
         Left            =   9840
         Top             =   4770
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   36
         Left            =   8760
         Top             =   4770
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   35
         Left            =   7680
         Top             =   4770
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   34
         Left            =   6600
         Top             =   4770
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   33
         Left            =   5520
         Top             =   4770
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   32
         Left            =   4440
         Top             =   4770
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   31
         Left            =   3360
         Top             =   4770
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   30
         Left            =   2280
         Top             =   4770
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   29
         Left            =   1200
         Top             =   4770
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   28
         Left            =   120
         Top             =   4770
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   27
         Left            =   14190
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   26
         Left            =   13110
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   25
         Left            =   12030
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   24
         Left            =   10950
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   23
         Left            =   9870
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   22
         Left            =   8790
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   21
         Left            =   7710
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   20
         Left            =   6630
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   19
         Left            =   5550
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   18
         Left            =   4470
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   17
         Left            =   3390
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   16
         Left            =   2310
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   15
         Left            =   1230
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   14
         Left            =   150
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   13
         Left            =   14190
         Top             =   750
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   12
         Left            =   13110
         Top             =   750
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   11
         Left            =   12030
         Top             =   750
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   10
         Left            =   10950
         Top             =   750
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   9
         Left            =   9870
         Top             =   750
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   8
         Left            =   8790
         Top             =   750
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   7
         Left            =   7710
         Top             =   750
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   6
         Left            =   6630
         Top             =   750
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   5
         Left            =   5550
         Top             =   750
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   4
         Left            =   4470
         Top             =   750
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   3
         Left            =   3390
         Top             =   750
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   2
         Left            =   2310
         Top             =   750
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   1
         Left            =   1230
         Top             =   750
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Index           =   0
         Left            =   150
         Top             =   750
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmTestIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' DB locale di test
Private DBTestIN As DBClass
Private DBTestOUT As DBClass
' variabile
Private m_numPagine As Integer
Private m_PaginaCorrente As Integer
Private m_Percorso As String
Public NomeFile As String

Private Sub Form_Activate()
  TimerLocale.Enabled = True
  FrameLav.Visible = False
  XPButton1(2).Visible = False
  XPButton1(1).Visible = True
End Sub

Private Sub Form_Load()
   Dim PathFile As String
   Dim a() As String
   Dim i As Integer
    
    TimerLocale.Enabled = False
    
    '====================================================================
    '                          carica il db di test
    '====================================================================
    '
    Dim IDitem, ItemVar, StringaCollegamento, TmpValore
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Const S7_nome_collegamento = "S7:[S7_connection_name1|VFD1|CP_L2_1:]"
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    Set DBTestIN = New DBClass
    Set DBTestOUT = New DBClass
    
    DBTestIN.DB_ID = "DBTestIN"
    DBTestOUT.DB_ID = "DBTestOUT"
    DBTestIN.Server = Attivo * Abs(frmKernel.ServerOpcOn) + Simula * Abs(frmKernel.SimulaON): frmKernel.ServerDisattivo = frmKernel.SimulaON Or (frmKernel.ServerOpcOn = False)
    DBTestOUT.Server = Attivo * Abs(frmKernel.ServerOpcOn) + Simula * Abs(frmKernel.SimulaON): frmKernel.ServerDisattivo = frmKernel.SimulaON Or (frmKernel.ServerOpcOn = False)
    
    Dim StrIO(0 To 1) As String
    Dim k As Integer
    
    ' query di ricerca in mappa dati
    StrIO(0) = "SELECT MappaDati.DBItem, MappaDati.Valore From MappaDati WHERE (((MappaDati.DBItem) Like 'E%') AND ((MappaDati.Gruppo)='Soffiatura') AND ((MappaDati.Attivato)=True)) OR (((MappaDati.DBItem) Like 'E%') AND ((MappaDati.Gruppo)='Entrata soffiatura') AND ((MappaDati.Attivato)=True));"
    StrIO(1) = "SELECT MappaDati.DBItem, MappaDati.Valore From MappaDati WHERE (((MappaDati.DBItem) Like 'A%') AND ((MappaDati.Gruppo)='Soffiatura') AND ((MappaDati.Attivato)=True)) OR (((MappaDati.DBItem) Like 'A%') AND ((MappaDati.Gruppo)='Entrata soffiatura') AND ((MappaDati.Attivato)=True));"
    On Error Resume Next
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\Plc.mdb;Persist Security Info=False"
    cn.Open
    For k = 0 To UBound(StrIO)
        With rs
           .Open StrIO(k), cn, adOpenKeyset, adLockReadOnly, adCmdText
           .MoveFirst
           While .EOF = False
              IDitem = "DBTest"
              ItemVar = .Fields("DBItem")
              StringaCollegamento = S7_nome_collegamento & .Fields("DBItem") & ",1"
              TmpValore = Val(IIf(IsNull(.Fields("Valore")) = False, .Fields("Valore"), 0))
              Select Case Left(ItemVar, 1)
              Case "E"
                  DBTestIN.Init IDitem & "IN", StringaCollegamento, ItemVar, TmpValore
              Case "A"
                  DBTestOUT.Init IDitem & "OUT", StringaCollegamento, ItemVar, TmpValore
             End Select
             .MoveNext
           Wend
           .Close
           Set .ActiveConnection = Nothing
        End With
    Next
    
    Set rs = Nothing
    Set cn = Nothing
    
    '===========  carica gli ingressi===============
    
    Dim h As Integer
    
    For h = 0 To 111
       If h < DBTestIN.NumItems Then
          Ingresso(h).Caption = DBTestIN.NomeItem(h + 1)
          Ingresso(h).Enabled = True
       Else
          Ingresso(h).Caption = ""
          Ingresso(h).Enabled = False
       End If
    Next
    
    For h = 112 To 223
       If h - 112 < DBTestOUT.NumItems Then
          Ingresso(h).Caption = DBTestOUT.NomeItem(h - 111)
          Ingresso(h).Enabled = True
       Else
          Ingresso(h).Caption = ""
          Ingresso(h).Enabled = False
       End If
    Next

End Sub
Private Sub Label2_Click()
    Hide
    Unload Me
End Sub

Private Sub TimerLocale_Timer()
    '===========  carica gli ingressi===============
'
    Dim h As Integer
'
    For h = 0 To 111
       If h < DBTestIN.NumItems Then
          Ingresso(h).value = Abs(DBTestIN.Item(h))
       End If
    Next
'
    For h = 112 To 223
       If h - 112 < DBTestOUT.NumItems Then
          Ingresso(h).value = Abs(DBTestOUT.Item(h))
       End If
    Next

'   Ingresso(0).Value = Abs(DBTestIN.Bit(5, 0, PlcIN))
'   Ingresso(1).Value = Abs(DBTestIN.Bit(5, 1, PlcIN))
'   Ingresso(2).Value = Abs(DBTestIN.Bit(5, 2, PlcIN))
'   Ingresso(3).Value = Abs(DBTestIN.Bit(5, 3, PlcIN))
End Sub

Private Sub XPButton1_Click(Index As Integer)
   Select Case Index
   Case 0
         TimerLocale.Enabled = False
         Set DBTestIN = Nothing
         Set DBTestOUT = Nothing
         Unload Me
   Case 1
         FrameLav.Visible = True
         FrameLav.ZOrder
         XPButton1(2).Visible = True
         XPButton1(1).Visible = False
   Case 2
         FrameLav.Visible = False
         FrameLav.ZOrder
         XPButton1(2).Visible = False
         XPButton1(1).Visible = True
   End Select
End Sub
