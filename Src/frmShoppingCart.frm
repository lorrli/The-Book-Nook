VERSION 5.00
Begin VB.Form frmShoppingCart 
   Caption         =   "Shopping Cart"
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   10545
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraprice 
      BackColor       =   &H000080FF&
      Height          =   2655
      Left            =   0
      TabIndex        =   55
      Top             =   8040
      Width           =   10215
      Begin VB.CommandButton cmdcontinue 
         Caption         =   "Click here to proceed to checkout."
         Height          =   1215
         Left            =   8640
         TabIndex        =   67
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdcalculateprice 
         Caption         =   "Calculate Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   66
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lbltotal 
         BackColor       =   &H000080FF&
         Height          =   375
         Left            =   2280
         TabIndex        =   65
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblcoupon 
         BackColor       =   &H000080FF&
         Caption         =   "Coupon Deduction:"
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
         TabIndex        =   64
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblcouponmessage 
         BackColor       =   &H000080FF&
         Caption         =   "10% of Subtotal = $"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   63
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblcouponamount 
         BackColor       =   &H000080FF&
         Height          =   375
         Left            =   4080
         TabIndex        =   62
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblshippingfee 
         BackColor       =   &H000080FF&
         Caption         =   "Shipping Fee: $5"
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
         Left            =   120
         TabIndex        =   61
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblsubtotalcaption 
         BackColor       =   &H000080FF&
         Caption         =   "Subtotal: $"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   60
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblsubtotal 
         BackColor       =   &H000080FF&
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
         Left            =   2160
         TabIndex        =   59
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lbltaxcaption 
         BackColor       =   &H000080FF&
         Caption         =   "Tax: $"
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
         Left            =   1200
         TabIndex        =   58
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lbltax 
         BackColor       =   &H000080FF&
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
         Left            =   2160
         TabIndex        =   57
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lbltotalcaption 
         BackColor       =   &H000080FF&
         Caption         =   "Total: $"
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
         Left            =   1200
         TabIndex        =   56
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdcouponredemption 
      Caption         =   "REDEEM YOUR COUPON HERE"
      Height          =   735
      Left            =   8640
      TabIndex        =   54
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Frame fraSaleorders 
      Caption         =   "Sales Details"
      Height          =   1935
      Left            =   120
      TabIndex        =   41
      Top             =   5880
      Width           =   8415
      Begin VB.ComboBox cboquantity9 
         Height          =   315
         ItemData        =   "frmShoppingCart.frx":0000
         Left            =   3360
         List            =   "frmShoppingCart.frx":0013
         TabIndex        =   50
         Text            =   "1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox cboquantity8 
         Height          =   315
         ItemData        =   "frmShoppingCart.frx":0026
         Left            =   3360
         List            =   "frmShoppingCart.frx":0039
         TabIndex        =   49
         Text            =   "1"
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox cboquantity7 
         Height          =   315
         ItemData        =   "frmShoppingCart.frx":004C
         Left            =   3360
         List            =   "frmShoppingCart.frx":005F
         TabIndex        =   48
         Text            =   "1"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblprice9 
         Height          =   255
         Left            =   4320
         TabIndex        =   53
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lblprice8 
         Height          =   255
         Left            =   4320
         TabIndex        =   52
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblprice7 
         Height          =   255
         Left            =   4320
         TabIndex        =   51
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblpack3 
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label lblpack2 
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label lblpack1 
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lblpackpricecaption 
         BackColor       =   &H00FF8080&
         Caption         =   "Price"
         Height          =   255
         Left            =   4800
         TabIndex        =   44
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblpackquantity 
         BackColor       =   &H00FF8080&
         Caption         =   "Quantity"
         Height          =   255
         Left            =   3360
         TabIndex        =   43
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblpackcaption 
         BackColor       =   &H00FF8080&
         Caption         =   "Item Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdmenu 
      Caption         =   "Click here to go back to menu to add more items to cart."
      Height          =   1095
      Left            =   8640
      TabIndex        =   40
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame frasupplyorders 
      Caption         =   "Supplies Details"
      Height          =   1815
      Left            =   120
      TabIndex        =   27
      Top             =   4080
      Width           =   8415
      Begin VB.ComboBox cboquantity6 
         Height          =   315
         ItemData        =   "frmShoppingCart.frx":0072
         Left            =   3360
         List            =   "frmShoppingCart.frx":0085
         TabIndex        =   33
         Text            =   "1"
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox cboquantity5 
         Height          =   315
         ItemData        =   "frmShoppingCart.frx":0098
         Left            =   3360
         List            =   "frmShoppingCart.frx":00AB
         TabIndex        =   32
         Text            =   "1"
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cboquantity4 
         Height          =   315
         ItemData        =   "frmShoppingCart.frx":00BE
         Left            =   3360
         List            =   "frmShoppingCart.frx":00D1
         TabIndex        =   31
         Text            =   "1"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblprice6 
         Height          =   255
         Left            =   4440
         TabIndex        =   39
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lblprice5 
         Height          =   255
         Left            =   4440
         TabIndex        =   38
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblprice4 
         Height          =   255
         Left            =   4440
         TabIndex        =   37
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblitem3 
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label lblitem2 
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblitem1 
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label lblsupplyPriceCaption 
         BackColor       =   &H00FF8080&
         Caption         =   "Price"
         Height          =   255
         Left            =   4800
         TabIndex        =   30
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblsupplyquantity 
         BackColor       =   &H00FF8080&
         Caption         =   "Quantity"
         Height          =   255
         Left            =   3360
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblItemCaption 
         BackColor       =   &H00FF8080&
         Caption         =   "Item Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame frabookorders 
      Caption         =   "Book Details "
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8415
      Begin VB.Frame frabooktype3 
         Height          =   855
         Left            =   4200
         TabIndex        =   24
         Top             =   2160
         Width           =   1335
         Begin VB.OptionButton optEbook3 
            Caption         =   "E-Book"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optRealBook3 
            Caption         =   "Real Book"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame frabooktype2 
         Height          =   855
         Left            =   4200
         TabIndex        =   21
         Top             =   1320
         Width           =   1335
         Begin VB.OptionButton optEbook2 
            Caption         =   "E-Book"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optRealBook2 
            Caption         =   "Real Book"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame frabooktype1 
         Height          =   855
         Left            =   4200
         TabIndex        =   18
         Top             =   480
         Width           =   1335
         Begin VB.OptionButton optRealBook1 
            Caption         =   "Real Book"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton optEbook1 
            Caption         =   "E-Book"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.ComboBox cboquantity3 
         Height          =   315
         ItemData        =   "frmShoppingCart.frx":00E4
         Left            =   5760
         List            =   "frmShoppingCart.frx":00F7
         TabIndex        =   15
         Text            =   "1"
         Top             =   2400
         Width           =   735
      End
      Begin VB.ComboBox cboquantity2 
         Height          =   315
         ItemData        =   "frmShoppingCart.frx":010A
         Left            =   5760
         List            =   "frmShoppingCart.frx":011D
         TabIndex        =   14
         Text            =   "1"
         Top             =   1560
         Width           =   735
      End
      Begin VB.ComboBox cboquantity1 
         Height          =   315
         ItemData        =   "frmShoppingCart.frx":0130
         Left            =   5760
         List            =   "frmShoppingCart.frx":0143
         TabIndex        =   8
         Text            =   "1"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblprice3 
         Height          =   375
         Left            =   6840
         TabIndex        =   17
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label lblprice2 
         Height          =   375
         Left            =   6840
         TabIndex        =   16
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblcovertype3 
         Height          =   615
         Left            =   2640
         TabIndex        =   13
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lblcovertype2 
         Height          =   615
         Left            =   2640
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lbltitle3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label lbltitle2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblprice1 
         Height          =   375
         Left            =   6840
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblcovertype1 
         Height          =   615
         Left            =   2640
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lbltitle1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblPriceCaption 
         BackColor       =   &H00FF8080&
         Caption         =   "Price"
         Height          =   255
         Left            =   6840
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblbookquantity 
         BackColor       =   &H00FF8080&
         Caption         =   "Quantity"
         Height          =   255
         Left            =   5640
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblbooktype 
         BackColor       =   &H00FF8080&
         Caption         =   "Book Type"
         Height          =   255
         Left            =   4200
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblcovertypeCaption 
         BackColor       =   &H00FF8080&
         Caption         =   "Harcover/softcover"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblTitleCaption 
         BackColor       =   &H00FF8080&
         Caption         =   "Book Title and Author"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Line lineseperator 
      X1              =   120
      X2              =   10080
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Image imglogobarpart2 
      Height          =   1095
      Left            =   8040
      Picture         =   "frmShoppingCart.frx":0156
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
   Begin VB.Image imglogobar 
      Height          =   1065
      Left            =   0
      Picture         =   "frmShoppingCart.frx":1E2C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8085
   End
End
Attribute VB_Name = "frmShoppingCart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lorraine Li
'Shopping Cart
'Jan.18, 2014
Option Explicit
'This form displays all the items that the user has checked off.
'It also allows the user to access the coupon, displays the subtotal and the total
'Along the sides, there are buttons that allow the user to continue with the purchase or to go back to the main menu.

'allows the user to go back to the menu
Private Sub cmdcontinue_Click()
    Unload frmShoppingCart
    frmCheckout.Show
End Sub
'shows the coupon form
Private Sub cmdcouponredemption_Click()
    frmcoupon.Show
End Sub
'shows the home page
Private Sub cmdmenu_Click()
    Unload frmShoppingCart
    frmhomepage.Show
End Sub

Private Sub Form_Load()
    'changes the background colour
    frmShoppingCart.BackColor = vbYellow
    'initializes numerous variables
    lbltitle1 = strchkoutbooktitle(0)
    lbltitle2 = strchkoutbooktitle(1)
    lbltitle3 = strchkoutbooktitle(2)
    
    lblitem1 = strchkoutitemname(0)
    lblitem2 = strchkoutitemname(1)
    lblitem3 = strchkoutitemname(2)
    
    lblpack1 = strchkoutpackname(0)
    lblpack2 = strchkoutpackname(1)
    lblpack3 = strchkoutpackname(2)
    
    lblcovertype1 = strchkoutbooktype(0)
    lblcovertype2 = strchkoutbooktype(1)
    lblcovertype3 = strchkoutbooktype(2)
  
    lblprice1 = dblchkoutprice(0)
    lblprice2 = dblchkoutprice(1)
    lblprice3 = dblchkoutprice(2)
    lblprice4 = dblchkoutitemprice(0)
    lblprice5 = dblchkoutitemprice(1)
    lblprice6 = dblchkoutitemprice(2)
    lblprice7 = dblchkoutpackprice(0)
    lblprice8 = dblchkoutpackprice(1)
    lblprice9 = dblchkoutpackprice(2)
    
    
    intquantity1 = cboquantity1.Text
    intquantity2 = cboquantity2.Text
    intquantity3 = cboquantity3.Text
    intquantity4 = cboquantity4.Text
    intquantity5 = cboquantity5.Text
    intquantity6 = cboquantity6.Text
    intquantity7 = cboquantity7.Text
    intquantity8 = cboquantity8.Text
    intquantity9 = cboquantity9.Text
    
    dbllineamount1 = 1
    dbllineamount2 = 1
    dbllineamount3 = 1
    dbllineamount4 = 1
    dbllineamount5 = 1
    dbllineamount6 = 1
    dbllineamount7 = 1
    dbllineamount8 = 1
    dbllineamount9 = 1
    'calculates the cost of each line of items
    dbllineamount1 = dblchkoutprice(0) * intquantity3
    dbllineamount2 = dblchkoutprice(1) * intquantity3
    dbllineamount3 = dblchkoutprice(2) * intquantity3
    dbllineamount4 = dblchkoutitemprice(0) * intquantity3
    dbllineamount5 = dblchkoutitemprice(1) * intquantity3
    dbllineamount6 = dblchkoutitemprice(2) * intquantity3
    dbllineamount7 = dblchkoutpackprice(0) * intquantity3
    dbllineamount8 = dblchkoutpackprice(1) * intquantity3
    dbllineamount9 = dblchkoutpackprice(2) * intquantity3
    
    cboquantity1.Enabled = False
    cboquantity2.Enabled = False
    cboquantity3.Enabled = False
    cboquantity4.Enabled = False
    cboquantity5.Enabled = False
    cboquantity6.Enabled = False
    cboquantity7.Enabled = False
    cboquantity8.Enabled = False
    cboquantity9.Enabled = False
    'displays the amount of lines depending on the amount of items chosen
    If dblbuybooks = 1 Then
        cboquantity1.Enabled = True
        frabooktype2.Enabled = False
        frabooktype3.Enabled = False
        lblprice2.Visible = False
        lblprice3.Visible = False
    ElseIf dblbuybooks = 2 Then
        cboquantity1.Enabled = True
        cboquantity2.Enabled = True
        frabooktype3.Enabled = False
        lblprice3.Visible = False
    ElseIf dblbuybooks = 3 Then
        cboquantity1.Enabled = True
        cboquantity2.Enabled = True
        cboquantity3.Enabled = True
    Else
        MsgBox ("You did not buy any books.")
        cboquantity1.Enabled = False
        cboquantity2.Enabled = False
        cboquantity3.Enabled = False
        frabooktype1.Enabled = False
        frabooktype2.Enabled = False
        frabooktype3.Enabled = False
    End If
    
    If dblbuysupplies = 1 Then
        cboquantity4.Enabled = True
        lblprice5.Visible = False
        lblprice6.Visible = False
    ElseIf dblbuysupplies = 2 Then
        cboquantity5.Enabled = True
        cboquantity4.Enabled = True
        lblprice6.Visible = False
    ElseIf dblbuysupplies = 3 Then
        cboquantity6.Enabled = True
        cboquantity5.Enabled = True
        cboquantity4.Enabled = True
    Else
        MsgBox ("You did not buy any supplies.")
    End If
    
    If dblbuysales = 1 Then
        cboquantity7.Enabled = True
        lblprice8.Visible = False
        lblprice9.Visible = False
    ElseIf dblbuysales = 2 Then
        cboquantity7.Enabled = True
        cboquantity8.Enabled = True
        lblprice9.Visible = False
    ElseIf dblbuysales = 3 Then
        cboquantity7.Enabled = True
        cboquantity8.Enabled = True
        cboquantity9.Enabled = True
    Else
        MsgBox ("You did not buy any sale items.")
    End If
    
End Sub
'displays the price after changing the quantity
Private Sub cboquantity1_Click()
    intquantity1 = cboquantity1.Text
    dbllineamount1 = dblchkoutprice(0) * intquantity1
    lblprice1 = "For " & intquantity1 & " copies, price = " & dbllineamount1

End Sub
'displays the price after changing the quantity
Private Sub cboquantity2_Click()
    intquantity2 = cboquantity2.Text
    dbllineamount2 = dblchkoutprice(1) * intquantity2
    lblprice2 = "For " & intquantity2 & " copies, price = " & dbllineamount2
   
End Sub
'displays the price after changing the quantity
Private Sub cboquantity3_Click()
    intquantity3 = cboquantity3.Text
    dbllineamount3 = dblchkoutprice(2) * intquantity3
    lblprice3 = "For " & intquantity3 & " copies, price = " & dbllineamount3
   
End Sub
'displays the price after changing the quantity
Private Sub cboquantity4_Click()
    intquantity4 = cboquantity4.Text
    dbllineamount4 = dblchkoutitemprice(0) * intquantity4
    lblprice4 = "For " & intquantity4 & " copies, price = " & dbllineamount4
End Sub
'displays the price after changing the quantity
Private Sub cboquantity5_Click()
    intquantity5 = cboquantity5.Text
    dbllineamount5 = dblchkoutitemprice(1) * intquantity5
    lblprice5 = "For " & intquantity5 & " copies, price = " & dbllineamount5
End Sub
'displays the price after changing the quantity
Private Sub cboquantity6_Click()
    intquantity6 = cboquantity6.Text
    dbllineamount6 = dblchkoutitemprice(2) * intquantity6
    lblprice6 = "For " & intquantity6 & " copies, price = " & dbllineamount6
End Sub
'displays the price after changing the quantity
Private Sub cboquantity7_Click()
    intquantity7 = cboquantity7.Text
    dbllineamount7 = dblchkoutpackprice(0) * intquantity7
    lblprice7 = "For " & intquantity7 & " copies, price = " & dbllineamount7
End Sub
'displays the price after changing the quantity
Private Sub cboquantity8_Click()
    intquantity8 = cboquantity8.Text
    dbllineamount8 = dblchkoutpackprice(1) * intquantity8
    lblprice8 = "For " & intquantity8 & " copies, price = " & dbllineamount8
End Sub
'displays the price after changing the quantity
Private Sub cboquantity9_Click()
    intquantity9 = cboquantity9.Text
    dbllineamount9 = dblchkoutpackprice(2) * intquantity9
    lblprice9 = "For " & intquantity9 & " copies, price = " & dbllineamount9
End Sub
'calculates the subtotal and possible coupon deduction
Private Sub cmdcalculateprice_Click()
    dblsubtotal = dbllineamount1 + dbllineamount2 + dbllineamount3 + dbllineamount4 + dbllineamount5 + dbllineamount6 + dbllineamount7 + dbllineamount8 + dbllineamount9
    If intcoupon = 1 Then
        lblcouponmessage.Visible = True
        lblcoupon.Visible = True
        lblcouponamount.Visible = True
        dblcoupondeduction = dblsubtotal * 0.1
        dblsubtotal = dblsubtotal - dblcoupondeduction
        dbltax = dblsubtotal * 0.13
        dbltotal = dblsubtotal + dbltax + 5 'adds the shipping fee, the tax, and the subtotal with the coupon deduction, together
        
        lblcouponamount.Caption = dblcoupondeduction
        lblsubtotal.Caption = dblsubtotal
        lbltax.Caption = Round(dbltax, 2)
        lbltotal.Caption = Round(dbltotal, 2)
        
    Else
       
        lblcouponmessage.Visible = False
        lblcoupon.Visible = False
        lblcouponamount.Visible = False
        
        dblsubtotal = dblsubtotal
        dbltax = dblsubtotal * 0.13
        dbltotal = dblsubtotal + dbltax + 5 'adds the shipping fee, the tax, and the subtotal together
        lblcouponamount.Caption = dblcoupondeduction
        lblsubtotal.Caption = dblsubtotal
        lbltax.Caption = Round(dbltax, 2)
        lbltotal.Caption = Round(dbltotal, 2)
    End If
End Sub

