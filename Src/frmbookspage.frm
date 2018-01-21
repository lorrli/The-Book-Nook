VERSION 5.00
Begin VB.Form frmbookspage 
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraLorrainesPicks 
      Caption         =   "Lorraine's Picks"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4455
      Left            =   2280
      TabIndex        =   4
      Top             =   1800
      Width           =   7095
      Begin VB.Timer tmrbookswitch 
         Interval        =   1800
         Left            =   3240
         Top             =   360
      End
      Begin VB.Label lblLtitles3 
         Caption         =   "The Akhenaten Adventure by P.B. Kerr (Children of the Lamp #1)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   4680
         TabIndex        =   23
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label lblLtitles3 
         Caption         =   "The Tomorrow Code by Brian Falkner"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   4680
         TabIndex        =   22
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblLtitles3 
         Caption         =   "Legend Series by Marie Lu"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   4680
         TabIndex        =   21
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblLtitles3 
         Caption         =   "Percy Jackson Series by  Rick Riordan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   4680
         TabIndex        =   20
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblLtitles3 
         Caption         =   "Inkheart by Cornelia Funke"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   4680
         TabIndex        =   19
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblLtitles3 
         Caption         =   "Unspoken by Sarah Rees Brennan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   4680
         TabIndex        =   18
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label1 
         Height          =   495
         Left            =   4680
         TabIndex        =   17
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblLtitles2 
         Caption         =   "Clockwork Princess by Cassandra Clare"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   2400
         TabIndex        =   16
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblLtitles2 
         Caption         =   "National Geographic Kids Almanac 2014"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   2400
         TabIndex        =   15
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblLtitles2 
         Caption         =   "      Cinder by                 Marissa Meyer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   2400
         TabIndex        =   14
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblLtitles2 
         Caption         =   "The Girl Who Could Fly by Victoria Forester"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   2400
         TabIndex        =   13
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label lblLtitles2 
         Caption         =   "Warriors Series by Erin Hunter"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   2400
         TabIndex        =   12
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblLtitles2 
         Caption         =   "A Wrinkle in Time by Madeleine L'Engle"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   2400
         TabIndex        =   11
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblLtitles 
         Caption         =   "The Hobbit by J.R.R.Tolkien"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label lblLtitles 
         Caption         =   "United We Spy by Ally Carter"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label lblLtitles 
         Caption         =   "Angelfall by Susan Ee"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label lblLtitles 
         Caption         =   "Hunger Games by Suzanne Collins"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label lblLtitles 
         Caption         =   "Graffiti Moon by Cath Crowley"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label lblLtitles 
         Caption         =   "Harry Potter Series by J.K.Rowling"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Image imgLbooks3 
         Height          =   2895
         Index           =   5
         Left            =   4680
         Picture         =   "frmbookspage.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2175
      End
      Begin VB.Image imgLbooks3 
         Height          =   2895
         Index           =   4
         Left            =   4680
         Picture         =   "frmbookspage.frx":2513A
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2175
      End
      Begin VB.Image imgLbooks3 
         Height          =   2895
         Index           =   3
         Left            =   4680
         Picture         =   "frmbookspage.frx":49E24
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2175
      End
      Begin VB.Image imgLbooks3 
         Height          =   2895
         Index           =   2
         Left            =   4680
         Picture         =   "frmbookspage.frx":4C1CC
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2175
      End
      Begin VB.Image imgLbooks3 
         Height          =   2895
         Index           =   1
         Left            =   4680
         Picture         =   "frmbookspage.frx":4E068
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2175
      End
      Begin VB.Image imgLbooks3 
         Height          =   2895
         Index           =   0
         Left            =   4680
         Picture         =   "frmbookspage.frx":73002
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2175
      End
      Begin VB.Image imgLbooks1 
         Height          =   2895
         Index           =   5
         Left            =   120
         Picture         =   "frmbookspage.frx":8EE84
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2055
      End
      Begin VB.Image imgLbooks1 
         Height          =   2895
         Index           =   4
         Left            =   120
         Picture         =   "frmbookspage.frx":B3FBA
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2055
      End
      Begin VB.Image imgLbooks2 
         Height          =   2895
         Index           =   5
         Left            =   2400
         Picture         =   "frmbookspage.frx":D90F4
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2055
      End
      Begin VB.Image imgLbooks2 
         Height          =   2895
         Index           =   4
         Left            =   2400
         Picture         =   "frmbookspage.frx":DBCB3
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2055
      End
      Begin VB.Image imgLbooks2 
         Height          =   2865
         Index           =   3
         Left            =   2400
         Picture         =   "frmbookspage.frx":100DED
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2010
      End
      Begin VB.Image imgLbooks2 
         Height          =   2895
         Index           =   2
         Left            =   2400
         Picture         =   "frmbookspage.frx":125CFF
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2055
      End
      Begin VB.Image imgLbooks2 
         Height          =   2895
         Index           =   1
         Left            =   2400
         Picture         =   "frmbookspage.frx":14AC01
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2055
      End
      Begin VB.Image imgLbooks2 
         Height          =   2895
         Index           =   0
         Left            =   2400
         Picture         =   "frmbookspage.frx":16FE43
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2040
      End
      Begin VB.Image imgLbooks1 
         Height          =   2895
         Index           =   3
         Left            =   120
         Picture         =   "frmbookspage.frx":194D55
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2055
      End
      Begin VB.Image imgLbooks1 
         Height          =   2895
         Index           =   2
         Left            =   120
         Picture         =   "frmbookspage.frx":1B9E77
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2055
      End
      Begin VB.Image imgLbooks1 
         Height          =   2775
         Index           =   0
         Left            =   120
         Picture         =   "frmbookspage.frx":1DED79
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1935
      End
      Begin VB.Image imgLbooks1 
         Height          =   2895
         Index           =   1
         Left            =   120
         Picture         =   "frmbookspage.frx":1E0ED6
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdNonfictionandClassics 
      Caption         =   "Non-fiction and Classics"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MaskColor       =   &H000040C0&
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdTeen 
      Caption         =   "Teen"
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
      Left            =   240
      TabIndex        =   1
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdChildren 
      Caption         =   "Children "
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
      Left            =   240
      TabIndex        =   0
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Image imglogobarpart2 
      Height          =   1575
      Left            =   7320
      Picture         =   "frmbookspage.frx":205FF8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
   Begin VB.Image imglogobar 
      Height          =   1545
      Left            =   0
      Picture         =   "frmbookspage.frx":207CCE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7365
   End
End
Attribute VB_Name = "frmbookspage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lorraine Li
'Jan.12,2014
'Book Home form
Option Explicit
'This form is where the user can access the different age groups of books
Private intIncrement As Integer

Private Sub cmdback_Click()
    Unload frmbookspage
    frmhomepage.Show
End Sub

Private Sub cmdChildren_Click()
    Unload frmbookspage
    frmchildrenbooks.Show
End Sub

Private Sub cmdNonfictionandClassics_Click()
    MsgBox ("Sorry this page is currently inavailable. Sorry for the Inconvenience.")
End Sub

Private Sub cmdTeen_Click()
    Unload frmbookspage
    frmteenbooks.Show
End Sub

Private Sub Form_Load()
    'changes the background colour
    frmbookspage.BackColor = vbGreen
    fraLorrainesPicks.BackColor = vbWhite
    'displays all the images being shown in the Lorraine's Picks frame
    imgLbooks1(0).Visible = True
    imgLbooks1(1).Visible = False
    imgLbooks1(2).Visible = False
    imgLbooks1(3).Visible = False
    imgLbooks1(4).Visible = False
    imgLbooks1(5).Visible = False
    
    imgLbooks2(0).Visible = True
    imgLbooks2(1).Visible = False
    imgLbooks2(2).Visible = False
    imgLbooks2(3).Visible = False
    imgLbooks2(4).Visible = False
    imgLbooks2(5).Visible = False
    
    imgLbooks3(0).Visible = True
    imgLbooks3(1).Visible = False
    imgLbooks3(2).Visible = False
    imgLbooks3(3).Visible = False
    imgLbooks3(4).Visible = False
    imgLbooks3(5).Visible = False
    
    lblLtitles(0).Visible = True
    lblLtitles(1).Visible = False
    lblLtitles(2).Visible = False
    lblLtitles(3).Visible = False
    lblLtitles(4).Visible = False
    lblLtitles(5).Visible = False
    
    lblLtitles2(0).Visible = True
    lblLtitles2(1).Visible = False
    lblLtitles2(2).Visible = False
    lblLtitles2(3).Visible = False
    lblLtitles2(4).Visible = False
    lblLtitles2(5).Visible = False
    
    lblLtitles3(0).Visible = True
    lblLtitles3(1).Visible = False
    lblLtitles3(2).Visible = False
    lblLtitles3(3).Visible = False
    lblLtitles3(4).Visible = False
    lblLtitles3(5).Visible = False
    
    intIncrement = 0
End Sub

Private Sub tmrbookswitch_Timer()
    
    imgLbooks1(intIncrement).Visible = False
    imgLbooks2(intIncrement).Visible = False
    imgLbooks3(intIncrement).Visible = False
    lblLtitles(intIncrement).Visible = False
    lblLtitles2(intIncrement).Visible = False
    lblLtitles3(intIncrement).Visible = False
    'sets the timer to switch between the pictures
    If intIncrement = 5 Then
        intIncrement = 0
    Else
        intIncrement = intIncrement + 1
    End If
    imgLbooks1(intIncrement).Visible = True
    imgLbooks2(intIncrement).Visible = True
    imgLbooks3(intIncrement).Visible = True
    lblLtitles(intIncrement).Visible = True
    lblLtitles2(intIncrement).Visible = True
    lblLtitles3(intIncrement).Visible = True
End Sub
