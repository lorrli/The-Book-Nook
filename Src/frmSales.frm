VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSales 
   BackColor       =   &H00FF00FF&
   Caption         =   "Sales"
   ClientHeight    =   8205
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcheckout 
      Caption         =   "Click to Checkout"
      Height          =   375
      Left            =   8880
      TabIndex        =   8
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Frame fraSales 
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   10215
      Begin MSComctlLib.ImageList ilsalepics 
         Left            =   480
         Top             =   2520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   201
         ImageHeight     =   251
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSales.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSales.frx":25086
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSales.frx":DD848
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSales.frx":19600A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkaddcart 
         Caption         =   "Check to Add to Cart"
         DataSource      =   "imgBooks"
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   2160
         TabIndex        =   3
         Top             =   5160
         Width           =   1815
      End
      Begin VB.ComboBox cbosales 
         Height          =   315
         ItemData        =   "frmSales.frx":1DFD54
         Left            =   120
         List            =   "frmSales.frx":1DFD64
         TabIndex        =   2
         Text            =   "--------select--------"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblprice 
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblpricedes 
         Caption         =   "PRICE = $"
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label lblpackdes 
         Height          =   5175
         Left            =   5280
         TabIndex        =   5
         Top             =   120
         Width           =   4815
      End
      Begin VB.Label lblpackname 
         Height          =   495
         Left            =   2040
         TabIndex        =   4
         Top             =   3600
         Width           =   3135
      End
      Begin VB.Image imgsales 
         Height          =   3255
         Left            =   2040
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblsales 
         Caption         =   "Items on Sale"
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
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Image imglogobarpart2 
      Height          =   1815
      Left            =   8880
      Picture         =   "frmSales.frx":1DFDE2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1815
   End
   Begin VB.Image imglogobar 
      Height          =   1785
      Left            =   0
      Picture         =   "frmSales.frx":1E1AB8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8955
   End
   Begin VB.Menu mnubacktomenu 
      Caption         =   "&Back to Menu"
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lorraine Li
'Jan.18,2014
'Supplies form
Option Explicit
'This form allows the user to purchase items on sale
Private Sub cbosales_Click()
    'displays the pictures, item names, item price, and item description in this combobox
    imgsales.Picture = ilsalepics.ListImages.Item(cbosales.ListIndex + 1).Picture
    lblpackname.Caption = strPackNameList(cbosales.ListIndex)
    lblprice = dblpackpricelist(cbosales.ListIndex)
    lblpackdes.Caption = strPackdes(cbosales.ListIndex)
    chkaddcart.Visible = True
    chkaddcart.Value = vbUnchecked
End Sub
Private Sub chkaddcart_Click()
    'prevents the user from buying more than 3 items
    'assigns some variables with values to be used in the shopping cart and checkout form
    If chkaddcart.Value = vbChecked Then
        dblbuysales = dblbuysales + 1
        If dblbuysales >= 4 Then
            MsgBox ("Sorry you already have three packages in your cart.Please proceed to checkout. Or go back to home to add books or supply items to your cart.")
            cbosales.Enabled = False
            dblbuysales = 3
        Else
            strchkoutpackname(dblbuysales - 1) = lblpackname.Caption
            dblchkoutpackprice(dblbuysales - 1) = lblprice
            MsgBox ("You have added " & strchkoutpackname(dblbuysales - 1) & " to your cart.")
            
        End If
    Else
    
    End If
End Sub
Private Sub cmdcheckout_Click()
    'brings user to shopping cart form
    If dblbuysales > 0 And dblbuysales < 4 Then
        MsgBox ("You have " & dblbuysales & " items for checkout")
        Unload frmSales
        frmShoppingCart.Show
    Else
        MsgBox ("You have " & dblbuysales & " items for checkout. Please select at least one item.")
    End If
End Sub

Private Sub Form_Load()
    'hides the extra books, the checkbox, and the option buttons
    chkaddcart.Visible = False
    chkaddcart.Value = vbUnchecked
End Sub

Private Sub mnuBacktomenu_Click()
    'brings user back to home page
    Unload frmsupplies
    frmhomepage.Show
End Sub
    
    
