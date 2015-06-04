VERSION 5.00
Object = "{3020F9BD-8F7D-4F6D-9B07-88E6EE6A69F6}#1.0#0"; "BarraOpzioni.ocx"
Begin VB.MDIForm MainMDIForm 
   BackColor       =   &H8000000C&
   ClientHeight    =   9240
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14835
   Icon            =   "MainMDIForm.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   14775
      TabIndex        =   0
      Top             =   7965
      Visible         =   0   'False
      Width           =   14835
      Begin BarraOpzioni.Barra ù 
         Height          =   1095
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   1931
         PulOrdiniPos    =   1200
         PulOrdiniAbi    =   -1  'True
         PulOrdiniVis    =   -1  'True
         PulStoricoPos   =   2400
         PulstoricoAbi   =   -1  'True
         PulStoricoVis   =   -1  'True
         PulMappaPos     =   3600
         PulMappaAbi     =   -1  'True
         PulMappaVis     =   -1  'True
         PulPaccoPos     =   4800
         PulPaccoAbi     =   -1  'True
         PulPaccoVis     =   -1  'True
         PulreggePos     =   6000
         PulReggeAbi     =   -1  'True
         PulReggeVis     =   -1  'True
         PulPesaPos      =   7200
         PulPesaAbi      =   -1  'True
         PulPesaVis      =   -1  'True
         PulAllarmiPos   =   8400
         PulAllarmiAbi   =   -1  'True
         PulAllarmiVis   =   -1  'True
         PulServicePos   =   9600
         PulServiceAbi   =   -1  'True
         PulServiceVis   =   -1  'True
      End
   End
End
Attribute VB_Name = "MainMDIForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
