VERSION 5.00
Begin VB.Form frmhomepage 
   Caption         =   "Home Page"
   ClientHeight    =   5955
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcoupon 
      Caption         =   "COUPON!!!!!"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdsales 
      Caption         =   "Sales"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdsupplies 
      Caption         =   "Supplies"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdbooks 
      Caption         =   "Books"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblerrormessage 
      BackColor       =   &H00808000&
      Caption         =   $"frmhomepage.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   5280
      Width           =   8175
   End
   Begin VB.Image imgSales 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   5280
      Picture         =   "frmhomepage.frx":00AC
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Image imgsupplies 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   2880
      Picture         =   "frmhomepage.frx":2504E
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Image imgbooks 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   480
      Picture         =   "frmhomepage.frx":4A2B4
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Image imglogobarpart2 
      Height          =   1575
      Left            =   6720
      Picture         =   "frmhomepage.frx":4BD7D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1935
   End
   Begin VB.Image imglogobar 
      Height          =   1545
      Left            =   0
      Picture         =   "frmhomepage.frx":4DA53
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6765
   End
   Begin VB.Menu mnuBooks 
      Caption         =   "&Books"
      Begin VB.Menu mnuChildren 
         Caption         =   "&Children"
      End
      Begin VB.Menu mnuTeen 
         Caption         =   "&Teen"
      End
      Begin VB.Menu mnuNonFicitonandClassics 
         Caption         =   "&Non-fiction and Classics"
      End
   End
   Begin VB.Menu mnuSupplies 
      Caption         =   "&Supplies"
      Begin VB.Menu mnuBookmarks 
         Caption         =   "&Bookmarks"
      End
      Begin VB.Menu mnuSchoolSupplies 
         Caption         =   "&School Supplies"
      End
   End
   Begin VB.Menu mnuSales 
      Caption         =   "&Sales"
   End
End
Attribute VB_Name = "frmhomepage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lorraine Li
'Final Project
'Jan.11,2014
'This is the home page where the user can receive the coupon code
'and visit the books, supplies, and sales forms
Private Sub cmdbooks_Click()
    Unload frmhomepage
    frmbookspage.Show
End Sub

Private Sub cmdcoupon_Click()
    'displays the coupon code
    MsgBox ("Limited Time Offer: 10 % off entire purchase. Remember the code for purchase at the end. Code: TH4NKS 4 3H0PP1N6 4T THE 3OOK N00K")
End Sub

Private Sub cmdexit_Click()
    'exits the entire program
    Unload frmhomepage
End Sub

Private Sub cmdsales_Click()
    Unload frmhomepage
    frmSales.Show
End Sub

Private Sub cmdsupplies_Click()
    Unload frmhomepage
    frmsupplies.Show
End Sub

Private Sub Form_Load()
    'changes the background colour
    frmhomepage.BackColor = vbMagenta
End Sub

Private Sub mnuChildren_Click()
    Unload frmhomepage
    frmchildrenbooks.Show
End Sub
Private Sub mnuNonFicitonandClassics_Click()
    MsgBox ("Sorry this page is currently inavailable. Sorry for the Inconvenience.")
End Sub

Private Sub mnuSales_Click()
    Unload frmhomepage
    frmSales.Show
End Sub

Private Sub mnuTeen_Click()
    Unload frmhomepage
    frmteenbooks.Show
End Sub
