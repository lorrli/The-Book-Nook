VERSION 5.00
Begin VB.Form frmcoupon 
   Caption         =   "Redeem Your Coupon"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtcouponcode 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   1050
      Width           =   2895
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdsubmit 
      Caption         =   "Submit"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblfinalmessage 
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
      TabIndex        =   5
      Top             =   1560
      Width           =   6015
   End
   Begin VB.Label lblentercode 
      Caption         =   "Coupon Code:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblinstructions 
      Caption         =   "Please type in the code given from your coupon."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmcoupon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lorraine
'Coupon Form
'Jan.19,2014
Private Sub cmdexit_Click()
    'exits the coupon form
    Unload frmcoupon
End Sub

Private Sub cmdsubmit_Click()
    'declares a variable
    Dim strCode As String
    
    strCode = txtcouponcode.Text
    'validates if the coupon code is correct or not.
    If strCode = "TH4NKS 4 3H0PP1N6 4T THE 3OOK N00K" Then
        lblfinalmessage.Caption = "Congratulations! You get 10% off your final purchase. Press exit to continue with your purchase. Then press calculate price."
        intcoupon = 1
    Else
        lblfinalmessage.Caption = "Invalid coupon code. Please enter a valid code."
        intcoupon = 0
    End If
End Sub
