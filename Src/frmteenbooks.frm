VERSION 5.00
Begin VB.Form frmteenbooks 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Teen Books"
   ClientHeight    =   8475
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgLegendseries 
      Height          =   1695
      Left            =   6960
      Picture         =   "frmteenbooks.frx":0000
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Image imgCinder 
      Height          =   1695
      Left            =   5280
      Picture         =   "frmteenbooks.frx":23A8
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Image imgAngelfall 
      Height          =   1695
      Left            =   3600
      Picture         =   "frmteenbooks.frx":272BA
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Image imghungergames 
      Height          =   1695
      Left            =   1920
      Picture         =   "frmteenbooks.frx":4C3DC
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Image imgDarkestMinds 
      Height          =   1695
      Left            =   240
      Picture         =   "frmteenbooks.frx":712DE
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label lbldystopian 
      Caption         =   "Dystopian Novels"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Image imgEsperenza 
      Height          =   1815
      Left            =   9480
      Picture         =   "frmteenbooks.frx":74C72
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Image imgCP2 
      Height          =   1815
      Left            =   7800
      Picture         =   "frmteenbooks.frx":77562
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label lblHistorical 
      Caption         =   "Historical Novels"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   4
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Image imgTFIOS 
      Height          =   1695
      Left            =   6240
      Picture         =   "frmteenbooks.frx":7A121
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Image imgBrokenhearts 
      Height          =   1695
      Left            =   4680
      Picture         =   "frmteenbooks.frx":7C99C
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label lblRealisticFiction 
      Caption         =   "Realistic Books"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Image imgGraffitiMoon 
      Height          =   1695
      Left            =   3240
      Picture         =   "frmteenbooks.frx":7E9A4
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Image imgHitchhiker 
      Height          =   1695
      Left            =   1680
      Picture         =   "frmteenbooks.frx":A3AC6
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Image imgtomorrowcode 
      Height          =   1695
      Left            =   120
      Picture         =   "frmteenbooks.frx":A61B0
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label lblscifi 
      Caption         =   "Science Ficiton Novels"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Image imgTMI 
      Height          =   1935
      Left            =   9120
      Picture         =   "frmteenbooks.frx":CB2EA
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Image imgLordofRings 
      Height          =   1935
      Left            =   7320
      Picture         =   "frmteenbooks.frx":CEACA
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Image imgVampireAcademy 
      Height          =   1935
      Left            =   5520
      Picture         =   "frmteenbooks.frx":CFE6B
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Image imgGraceling 
      Height          =   1935
      Left            =   3720
      Picture         =   "frmteenbooks.frx":D1DDE
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblFantasybooks 
      Caption         =   "Fantasy Books"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Image imgUnspoken 
      Height          =   1920
      Left            =   1920
      Picture         =   "frmteenbooks.frx":D46B9
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1650
   End
   Begin VB.Image imgUnitedSpy 
      Height          =   1935
      Left            =   120
      Picture         =   "frmteenbooks.frx":F053B
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Image imglogobarpart2 
      Height          =   1335
      Left            =   8760
      Picture         =   "frmteenbooks.frx":115675
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2295
   End
   Begin VB.Image imglogobar 
      Height          =   1305
      Left            =   0
      Picture         =   "frmteenbooks.frx":11734B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8805
   End
   Begin VB.Label lblMysteryBooks 
      Caption         =   "Mystery Books"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Menu mnuBacktomenu 
      Caption         =   "&Back to Menu"
   End
End
Attribute VB_Name = "frmteenbooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lorraine
'Teen Books Form
'Jan.19,2014
'This form displays the teen books.None of them are available for checkout.

Private Sub mnuBacktomenu_Click()
    'brings the user back to the home page
    Unload frmteenbooks
    frmhomepage.Show
End Sub
