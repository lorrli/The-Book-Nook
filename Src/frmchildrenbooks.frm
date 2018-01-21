VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmchildrenbooks 
   BackColor       =   &H0000C000&
   Caption         =   "Children Books"
   ClientHeight    =   10740
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   10740
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcheckout 
      Caption         =   "Click to Checkout"
      Height          =   375
      Left            =   8760
      TabIndex        =   11
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame framoreKidsbooks 
      Caption         =   "More Books"
      Height          =   5895
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   10335
      Begin VB.Image imgChildrenofLamp 
         Height          =   2055
         Left            =   4080
         Picture         =   "frmchildrenbooks.frx":0000
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Image imgholes 
         Height          =   2055
         Left            =   2160
         Picture         =   "frmchildrenbooks.frx":24CEA
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblmoreadventure 
         Caption         =   "Adventure"
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
         Left            =   2160
         TabIndex        =   23
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Image imgSistersGrimm 
         Height          =   1935
         Left            =   8160
         Picture         =   "frmchildrenbooks.frx":2685D
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgCharlieFactory 
         Height          =   1935
         Left            =   6120
         Picture         =   "frmchildrenbooks.frx":290E2
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgtaleofdesperaux 
         Height          =   1935
         Left            =   4080
         Picture         =   "frmchildrenbooks.frx":2CD31
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgMatilda 
         Height          =   1935
         Left            =   2040
         Picture         =   "frmchildrenbooks.frx":2EC76
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgPercyJackson 
         Height          =   1935
         Left            =   120
         Picture         =   "frmchildrenbooks.frx":32E55
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblmorefantasy 
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
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblmoreRealistic 
         Caption         =   "Realistic Fiction"
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
         TabIndex        =   8
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Image imgWinnDixie 
         Height          =   2100
         Left            =   120
         Picture         =   "frmchildrenbooks.frx":34CF1
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdshowmore 
      Caption         =   "Show More Books"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Frame frabookdescriptions 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   10335
      Begin VB.OptionButton optSoftcover 
         Caption         =   "Softcover"
         Height          =   255
         Left            =   3840
         TabIndex        =   20
         Top             =   3240
         Width           =   1335
      End
      Begin VB.OptionButton optHardcover 
         Caption         =   "Hardcover"
         Height          =   195
         Left            =   3840
         TabIndex        =   19
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CheckBox chkaddcart 
         Caption         =   "Check to Add to Cart"
         DataSource      =   "imgBooks"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   3120
         Width           =   1215
      End
      Begin VB.ComboBox cboAdventure 
         Height          =   315
         ItemData        =   "frmchildrenbooks.frx":36E5D
         Left            =   240
         List            =   "frmchildrenbooks.frx":36E67
         TabIndex        =   15
         Text            =   "--------select--------"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.ComboBox cboRealistic 
         Height          =   315
         ItemData        =   "frmchildrenbooks.frx":36E8F
         Left            =   240
         List            =   "frmchildrenbooks.frx":36E99
         TabIndex        =   14
         Text            =   "--------select--------"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.ComboBox cboMystery 
         Height          =   315
         ItemData        =   "frmchildrenbooks.frx":36EBC
         Left            =   240
         List            =   "frmchildrenbooks.frx":36EC6
         TabIndex        =   13
         Text            =   "--------select--------"
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cboFantasy 
         Height          =   315
         ItemData        =   "frmchildrenbooks.frx":36EE7
         Left            =   240
         List            =   "frmchildrenbooks.frx":36EF7
         TabIndex        =   12
         Text            =   "--------select--------"
         Top             =   1200
         Width           =   1815
      End
      Begin MSComctlLib.ImageList ilbookpics 
         Left            =   1800
         Top             =   2880
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   120
         ImageHeight     =   183
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmchildrenbooks.frx":36F37
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmchildrenbooks.frx":3843F
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmchildrenbooks.frx":5D30D
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmchildrenbooks.frx":5F47A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmchildrenbooks.frx":845C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmchildrenbooks.frx":A94D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmchildrenbooks.frx":CE47C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmchildrenbooks.frx":E10CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmchildrenbooks.frx":E39E1
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmchildrenbooks.frx":108903
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblSC 
         Caption         =   "Softcover price = $"
         Height          =   495
         Left            =   7800
         TabIndex        =   22
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label lblHC 
         Caption         =   "Hardcover price = $"
         Height          =   495
         Left            =   5520
         TabIndex        =   21
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label lblSCprice 
         Height          =   495
         Left            =   9240
         TabIndex        =   18
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblBookTitle 
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         Top             =   240
         Width           =   4215
      End
      Begin VB.Image imgBooks 
         Height          =   2535
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblHCprice 
         Height          =   495
         Left            =   6960
         TabIndex        =   10
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblAdventurebooks 
         Caption         =   "Adventure Books"
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
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   1935
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
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblbookdescriptions 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   4560
         TabIndex        =   5
         Top             =   480
         Width           =   5535
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
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1695
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
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Line lineseperator 
      X1              =   0
      X2              =   10560
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Image imglogobarpart2 
      Height          =   1335
      Left            =   8520
      Picture         =   "frmchildrenbooks.frx":12DB55
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
   Begin VB.Image imglogobar 
      Height          =   1305
      Left            =   0
      Picture         =   "frmchildrenbooks.frx":12F82B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8685
   End
   Begin VB.Menu mnubacktomenu 
      Caption         =   "&Back to Menu"
   End
End
Attribute VB_Name = "frmchildrenbooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Lorraine Li
'Jan.12,2014
'Children's books form
Option Explicit
'This form displays all the children books available for checkout
'It also displays several books under the more books button that can not be checked out

Private Sub cboMystery_Click()
    'displays the pictures, item names, item price, and item description in this combobox
    imgBooks.Picture = ilbookpics.ListImages.Item(cboMystery.ListIndex + 1).Picture
    lblBookTitle.Caption = strTitleList(cboMystery.ListIndex)
    lblHCprice = dblHPriceList(cboMystery.ListIndex)
    lblSCprice = dblSPriceList(cboMystery.ListIndex)
    lblbookdescriptions.Caption = strDescription(cboMystery.ListIndex)
    chkaddcart.Visible = True
    chkaddcart.Value = vbUnchecked
End Sub
Private Sub cboFantasy_Click()
    'displays the pictures, item names, item price, and item description in this combobox
    imgBooks.Picture = ilbookpics.ListImages.Item(cboFantasy.ListIndex + 3).Picture
    lblBookTitle.Caption = strTitleList(cboFantasy.ListIndex + 2)
    lblHCprice = dblHPriceList(cboFantasy.ListIndex + 2)
    lblSCprice = dblSPriceList(cboFantasy.ListIndex + 2)
    lblbookdescriptions.Caption = strDescription(cboFantasy.ListIndex + 2)
    chkaddcart.Visible = True
    chkaddcart.Value = vbUnchecked
End Sub
Private Sub cboRealistic_click()
    'displays the pictures, item names, item price, and item description in this combobox
    imgBooks.Picture = ilbookpics.ListImages.Item(cboRealistic.ListIndex + 7).Picture
    lblBookTitle.Caption = strTitleList(cboRealistic.ListIndex + 6)
    lblHCprice = dblHPriceList(cboRealistic.ListIndex + 6)
    lblSCprice = dblSPriceList(cboRealistic.ListIndex + 6)
    lblbookdescriptions.Caption = strDescription(cboRealistic.ListIndex + 6)
    chkaddcart.Visible = True
    chkaddcart.Value = vbUnchecked
End Sub
Private Sub cboAdventure_Click()
    'displays the pictures, item names, item price, and item description in this combobox
    imgBooks.Picture = ilbookpics.ListImages.Item(cboAdventure.ListIndex + 9).Picture
    lblBookTitle.Caption = strTitleList(cboAdventure.ListIndex + 8)
    lblHCprice = dblHPriceList(cboAdventure.ListIndex + 8)
    lblSCprice = dblSPriceList(cboAdventure.ListIndex + 8)
    lblbookdescriptions.Caption = strDescription(cboAdventure.ListIndex + 8)
    chkaddcart.Visible = True
    chkaddcart.Value = vbUnchecked
End Sub

Private Sub chkaddcart_Click()
    'prevents the user from buying more than 3 items
    'assigns some variables with values to be used in the shopping cart and checkout form
    If chkaddcart.Value = vbChecked Then
        dblbuybooks = dblbuybooks + 1
        If dblbuybooks >= 4 Then
            MsgBox ("Sorry you already have three books in your cart.Please proceed to checkout. Or go back to the home page to add supplies or sale items to your cart.")
            cboMystery.Enabled = False
            cboFantasy.Enabled = False
            cboRealistic.Enabled = False
            cboAdventure.Enabled = False
            dblbuybooks = 3
        Else
            strchkoutbooktitle(dblbuybooks - 1) = lblBookTitle.Caption
            MsgBox ("You have added " & strchkoutbooktitle(dblbuybooks - 1) & " to your cart.")
            If intcount = 1 Then
                dblchkoutprice(dblbuybooks - 1) = lblHCprice.Caption
                strchkoutbooktype(dblbuybooks - 1) = "hardcover"
            ElseIf intcount = 2 Then
                dblchkoutprice(dblbuybooks - 1) = lblSCprice.Caption
                strchkoutbooktype(dblbuybooks - 1) = "softcover"
            Else
                MsgBox ("Please select a book type.")
            End If
        End If
    Else
    
    End If
End Sub
Private Sub cmdcheckout_Click()
    'brings user to shopping cart form if they have chosen at least one book
    If dblbuybooks > 0 And dblbuybooks < 4 Then
        MsgBox ("You have " & dblbuybooks & " books for checkout")
        Unload frmchildrenbooks
        frmShoppingCart.Show
    Else
        MsgBox ("You have " & dblbuybooks & " books for checkout.Please select at least one item.")
    End If
End Sub
Private Sub cmdshowmore_Click()
    'displays more books
    framoreKidsbooks.Visible = True
End Sub
Private Sub Form_Load()
    'hides the extra books, the checkbox, and the option buttons
    framoreKidsbooks.Visible = False
    chkaddcart.Visible = False
    chkaddcart.Value = vbUnchecked
    optSoftcover.Value = True
    optHardcover.Value = False
End Sub

Private Sub mnuBacktomenu_Click()
    Unload frmchildrenbooks
    frmhomepage.Show
End Sub

Private Sub optHardcover_Click()
    'gives the variable a value
    intcount = 1
End Sub

Private Sub optSoftcover_Click()
    'gives the variable a value
    intcount = 2
End Sub
