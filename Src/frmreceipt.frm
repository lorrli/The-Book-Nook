VERSION 5.00
Begin VB.Form frmreceipt 
   Caption         =   "Receipt"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   8160
      TabIndex        =   55
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   8160
      TabIndex        =   54
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Frame frapaymentinfo 
      Caption         =   "Payment Details"
      Height          =   1695
      Left            =   120
      TabIndex        =   43
      Top             =   5880
      Width           =   7455
      Begin VB.Label lblcardtype 
         Height          =   375
         Left            =   5400
         TabIndex        =   53
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblcardtypecaption 
         Caption         =   "Card Type:"
         Height          =   255
         Left            =   4320
         TabIndex        =   52
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblcardnum 
         Height          =   255
         Left            =   1800
         TabIndex        =   51
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lblcardnumbercaption 
         Caption         =   "Card Number:"
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
         TabIndex        =   50
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lbltotal 
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
         Left            =   1680
         TabIndex        =   49
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblamountdue 
         Caption         =   "Amount Due: $"
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
         TabIndex        =   48
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lbladdress 
         Height          =   255
         Left            =   1200
         TabIndex        =   47
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label lblAddresscaption 
         Caption         =   "Address:"
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
         TabIndex        =   46
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblname 
         Height          =   255
         Left            =   2040
         TabIndex        =   45
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblcardholdername 
         Caption         =   "Cardholder Name:"
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
         TabIndex        =   44
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame fraSaleorders 
      Caption         =   "Sales Details"
      Height          =   1815
      Left            =   120
      TabIndex        =   30
      Top             =   3960
      Width           =   7455
      Begin VB.Label lblquantity9 
         Height          =   375
         Left            =   3360
         TabIndex        =   42
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblquantity8 
         Height          =   375
         Left            =   3360
         TabIndex        =   41
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblquantity7 
         Height          =   375
         Left            =   3360
         TabIndex        =   40
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblpackcaption 
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
         TabIndex        =   39
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblpackquantity 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   3360
         TabIndex        =   38
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblpackpricecaption 
         Caption         =   "Price"
         Height          =   255
         Left            =   4800
         TabIndex        =   37
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblpack1 
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lblpack2 
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label lblpack3 
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label lblprice7 
         Height          =   255
         Left            =   4800
         TabIndex        =   33
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblprice8 
         Height          =   255
         Left            =   4800
         TabIndex        =   32
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblprice9 
         Height          =   255
         Left            =   4800
         TabIndex        =   31
         Top             =   1440
         Width           =   2175
      End
   End
   Begin VB.Frame frasupplyorders 
      Caption         =   "Supplies Details"
      Height          =   1815
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   7455
      Begin VB.Label lblquantity6 
         Height          =   255
         Left            =   3360
         TabIndex        =   29
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblquantity5 
         Height          =   255
         Left            =   3360
         TabIndex        =   28
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblquantity4 
         Height          =   255
         Left            =   3360
         TabIndex        =   27
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblItemCaption 
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
         TabIndex        =   26
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblsupplyquantity 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   3360
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblsupplyPriceCaption 
         Caption         =   "Price"
         Height          =   255
         Left            =   5160
         TabIndex        =   24
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblitem1 
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label lblitem2 
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblitem3 
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label lblprice4 
         Height          =   255
         Left            =   5040
         TabIndex        =   20
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblprice5 
         Height          =   255
         Left            =   5040
         TabIndex        =   19
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblprice6 
         Height          =   255
         Left            =   5040
         TabIndex        =   18
         Top             =   1440
         Width           =   2295
      End
   End
   Begin VB.Frame frabookorders 
      Caption         =   "Book Details "
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7455
      Begin VB.Label lblquantity3 
         Height          =   255
         Left            =   4200
         TabIndex        =   16
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblquantity2 
         Height          =   255
         Left            =   4200
         TabIndex        =   15
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblquantity1 
         Height          =   255
         Left            =   4200
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblTitleCaption 
         Caption         =   "Book Title and Author"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblcovertypeCaption 
         Caption         =   "Harcover/softcover"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblbookquantity 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   4200
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblPriceCaption 
         Caption         =   "Price"
         Height          =   255
         Left            =   5880
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbltitle1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblcovertype1 
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblprice1 
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lbltitle2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lbltitle3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lblcovertype2 
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblcovertype3 
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblprice2 
         Height          =   375
         Left            =   5760
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblprice3 
         Height          =   375
         Left            =   5880
         TabIndex        =   1
         Top             =   1200
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmreceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lorraine Li
'Receipt Form
'Jan.19, 2014
'This form displays the information collected from the user, except for phone number and ebook/realbook, and displays it as a receipt

Private Sub cmdPrint_Click()
    MsgBox ("Sorry the printer is currently out of ink.")
End Sub

Private Sub cmdsave_Click()
    'brings the user to the thank you form
    MsgBox ("You're information has been saved.")
    Unload frmreceipt
    frmThankYou.Show
End Sub

Private Sub Form_Load()
    'displays the information collected from the past forms
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
    
    
    
    dbllineamount1 = 1
    dbllineamount2 = 1
    dbllineamount3 = 1
    dbllineamount4 = 1
    dbllineamount5 = 1
    dbllineamount6 = 1
    dbllineamount7 = 1
    dbllineamount8 = 1
    dbllineamount9 = 1
    
    dbllineamount1 = dblchkoutprice(0) * intquantity3
    dbllineamount2 = dblchkoutprice(1) * intquantity3
    dbllineamount3 = dblchkoutprice(2) * intquantity3
    dbllineamount4 = dblchkoutitemprice(0) * intquantity3
    dbllineamount5 = dblchkoutitemprice(1) * intquantity3
    dbllineamount6 = dblchkoutitemprice(2) * intquantity3
    dbllineamount7 = dblchkoutpackprice(0) * intquantity3
    dbllineamount8 = dblchkoutpackprice(1) * intquantity3
    dbllineamount9 = dblchkoutpackprice(2) * intquantity3
    
    lblname = strName
    lblAddress = strAddress
    lblcardnum = strCreditcardnum
    lbltotal = dbltotal
    lblcardtype = strCardChosen
End Sub

