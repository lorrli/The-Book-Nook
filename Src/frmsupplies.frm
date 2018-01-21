VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsupplies 
   Caption         =   "Supplies"
   ClientHeight    =   7605
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcheckout 
      Caption         =   "Click to Checkout"
      Height          =   375
      Left            =   9240
      TabIndex        =   9
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Frame frasupplies 
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   10815
      Begin VB.CheckBox chkaddcart 
         Caption         =   "Check to Add to Cart"
         DataSource      =   "imgBooks"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   3480
         Width           =   1215
      End
      Begin VB.ComboBox cboagenda 
         Height          =   315
         ItemData        =   "frmsupplies.frx":0000
         Left            =   240
         List            =   "frmsupplies.frx":000D
         TabIndex        =   7
         Text            =   "select agenda"
         Top             =   2760
         Width           =   2055
      End
      Begin VB.ComboBox cbonotebooks 
         Height          =   315
         ItemData        =   "frmsupplies.frx":0073
         Left            =   240
         List            =   "frmsupplies.frx":0083
         TabIndex        =   6
         Text            =   "select notebooks"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.ComboBox cbobookmarks 
         Height          =   315
         ItemData        =   "frmsupplies.frx":00ED
         Left            =   240
         List            =   "frmsupplies.frx":00FA
         TabIndex        =   2
         Text            =   "select bookmarks"
         Top             =   720
         Width           =   2055
      End
      Begin MSComctlLib.ImageList ilsuppliespics 
         Left            =   2400
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   215
         ImageHeight     =   235
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmsupplies.frx":013E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmsupplies.frx":28C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmsupplies.frx":10CE8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmsupplies.frx":12383
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmsupplies.frx":14D33
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmsupplies.frx":17486
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmsupplies.frx":1946E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmsupplies.frx":1BCBE
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmsupplies.frx":1C9E5
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmsupplies.frx":1E5B4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblsupplyname 
         Height          =   375
         Left            =   5880
         TabIndex        =   12
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblprice 
         Height          =   375
         Left            =   7680
         TabIndex        =   11
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label lblpricedes 
         Caption         =   "PRICE = $"
         Height          =   255
         Left            =   6840
         TabIndex        =   10
         Top             =   3840
         Width           =   855
      End
      Begin VB.Image imgsupplies 
         Height          =   2535
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblAgenda 
         Caption         =   "Agendas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lblNotebooks 
         Caption         =   "Notebooks"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblbookmarks 
         Caption         =   "Bookmarks"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblsupplydes 
         Height          =   3015
         Left            =   6000
         TabIndex        =   1
         Top             =   840
         Width           =   4575
      End
   End
   Begin VB.Image imglogobarpart2 
      Height          =   1815
      Left            =   9000
      Picture         =   "frmsupplies.frx":200E1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2055
   End
   Begin VB.Image imglogobar 
      Height          =   1785
      Left            =   0
      Picture         =   "frmsupplies.frx":21DB7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9045
   End
   Begin VB.Menu mnuBacktoMenu 
      Caption         =   "&Back to Menu"
   End
End
Attribute VB_Name = "frmsupplies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lorraine Li
'Jan.18,2014
'Supplies form
Option Explicit
'This form displays the supplies (bookmarks,agendas,notebooks) available for checkout.
'displays the pictures, item names, item price, and item description in this combobox
Private Sub cboagenda_Click()
    imgsupplies.Picture = ilsuppliespics.ListImages.Item(cboagenda.ListIndex + 8).Picture
    lblsupplyname.Caption = strSupplyNameList(cboagenda.ListIndex + 7)
    lblprice = dblsupplypriceList(cboagenda.ListIndex + 7)
    lblsupplydes.Caption = strSupplyDes(cboagenda.ListIndex + 7)
    chkaddcart.Visible = True
    chkaddcart.Value = vbUnchecked
End Sub
'displays the pictures, item names, item price, and item description in this combobox
Private Sub cbonotebooks_Click()
    imgsupplies.Picture = ilsuppliespics.ListImages.Item(cbonotebooks.ListIndex + 4).Picture
    lblsupplyname.Caption = strSupplyNameList(cbonotebooks.ListIndex + 3)
    lblprice = dblsupplypriceList(cbonotebooks.ListIndex + 3)
    lblsupplydes.Caption = strSupplyDes(cbonotebooks.ListIndex + 3)
    chkaddcart.Visible = True
    chkaddcart.Value = vbUnchecked
End Sub
'displays the pictures, item names, item price, and item description in this combobox
Private Sub cbobookmarks_Click()
    imgsupplies.Picture = ilsuppliespics.ListImages.Item(cbobookmarks.ListIndex + 1).Picture
    lblsupplyname.Caption = strSupplyNameList(cbobookmarks.ListIndex)
    lblprice = dblsupplypriceList(cbobookmarks.ListIndex)
    lblsupplydes.Caption = strSupplyDes(cbobookmarks.ListIndex)
    chkaddcart.Visible = True
    chkaddcart.Value = vbUnchecked
End Sub
Private Sub Form_Load()
    'changes the background colour
    frmsupplies.BackColor = vbCyan
    'hides the checkbox and leaves it unchecked
    chkaddcart.Visible = False
    chkaddcart.Value = vbUnchecked
End Sub
Private Sub chkaddcart_Click()
    'prevents the user from buying more than 3 items
    'assigns some variables with values to be used in the shopping cart and checkout form
    If chkaddcart.Value = vbChecked Then
        dblbuysupplies = dblbuysupplies + 1
        If dblbuysupplies >= 4 Then
            MsgBox ("Sorry you already have three supplies in your cart.Please proceed to checkout. Or go back to home to add books or sale items to your cart.")
            cboagenda.Enabled = False
            cbobookmarks.Enabled = False
            cbonotebooks.Enabled = False
            dblbuysupplies = 3
        Else
            strchkoutitemname(dblbuysupplies - 1) = lblsupplyname.Caption
            dblchkoutitemprice(dblbuysupplies - 1) = lblprice
            MsgBox ("You have added " & strchkoutitemname(dblbuysupplies - 1) & " to your cart.")
            
        End If
    Else
    
    End If
End Sub
Private Sub cmdcheckout_Click()
    'displays the shopping cart form if they have chosen at least one item for checkout
    If dblbuysupplies > 0 And dblbuysupplies < 4 Then
        MsgBox ("You have " & dblbuysupplies & " items for checkout")
        Unload frmsupplies
        frmShoppingCart.Show
    Else
        MsgBox ("You have " & dblbuysupplies & " items for checkout.Please select at least one item.")
    End If
End Sub
Private Sub mnuBacktomenu_Click()
    'brings the user back to the main menu
    Unload frmsupplies
    frmhomepage.Show
End Sub
