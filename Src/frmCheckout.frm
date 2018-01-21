VERSION 5.00
Begin VB.Form frmCheckout 
   BackColor       =   &H00404080&
   Caption         =   "Checkout"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next"
      Height          =   375
      Left            =   5520
      TabIndex        =   21
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Frame fraPayment 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Payment"
      Height          =   3135
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   6615
      Begin VB.ComboBox cboyear 
         Height          =   315
         ItemData        =   "frmCheckout.frx":0000
         Left            =   5520
         List            =   "frmCheckout.frx":0002
         TabIndex        =   20
         Text            =   "year"
         Top             =   2520
         Width           =   735
      End
      Begin VB.ComboBox cboexpirymonth 
         Height          =   315
         ItemData        =   "frmCheckout.frx":0004
         Left            =   4560
         List            =   "frmCheckout.frx":0006
         TabIndex        =   19
         Text            =   "month"
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtcardnum 
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   2520
         Width           =   2535
      End
      Begin VB.OptionButton optMastercard 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MasterCard"
         Height          =   255
         Left            =   4440
         TabIndex        =   13
         Top             =   1920
         Width           =   1335
      End
      Begin VB.OptionButton optamericanexpress 
         BackColor       =   &H00FFFFFF&
         Caption         =   "American Express"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   1920
         Width           =   1695
      End
      Begin VB.OptionButton optvisa 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Visa"
         Height          =   255
         Left            =   720
         MaskColor       =   &H00000000&
         TabIndex        =   11
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lblexpirydate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Expiry Date:"
         Height          =   255
         Left            =   4920
         TabIndex        =   18
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblcardnumber 
         BackColor       =   &H00FFFFFF&
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
         Left            =   240
         TabIndex        =   16
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lbltotal 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblamountdue 
         BackColor       =   &H00FFFFFF&
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
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.Image imgmastercard 
         Height          =   975
         Left            =   4200
         Picture         =   "frmCheckout.frx":0008
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1935
      End
      Begin VB.Image imgamericanexpress 
         Height          =   975
         Left            =   2280
         Picture         =   "frmCheckout.frx":2502E
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1815
      End
      Begin VB.Image imgvisa 
         Height          =   975
         Left            =   240
         Picture         =   "frmCheckout.frx":49EEC
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Frame frapersonalinfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Personal Information"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   6615
      Begin VB.TextBox txtaddress 
         Height          =   375
         Left            =   1320
         TabIndex        =   23
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtphonenum3 
         Height          =   285
         Left            =   3360
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtphonenum2 
         Height          =   285
         Left            =   2640
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtphonenum1 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtname 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H00FFFFFF&
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
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblhyphen 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
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
         Left            =   3240
         TabIndex        =   9
         Top             =   840
         Width           =   135
      End
      Begin VB.Label lblRightbracket 
         BackColor       =   &H00FFFFFF&
         Caption         =   ")"
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
         Left            =   2520
         TabIndex        =   8
         Top             =   840
         Width           =   135
      End
      Begin VB.Label lblLeftbracket 
         BackColor       =   &H00FFFFFF&
         Caption         =   "("
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
         Left            =   1680
         TabIndex        =   7
         Top             =   840
         Width           =   135
      End
      Begin VB.Label lblphonenum 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Phone Number:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblcardholdername 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Image imglogobarpart2 
      Height          =   1575
      Left            =   5280
      Picture         =   "frmCheckout.frx":6EF96
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1695
   End
   Begin VB.Image imglogobar 
      Height          =   1545
      Left            =   0
      Picture         =   "frmCheckout.frx":70C6C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5325
   End
End
Attribute VB_Name = "frmCheckout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Lorraine Li
'Jan.19,2014
'Checkout form
Option Explicit
'This form is where the user can pay for all the items from the shopping cart
Private Sub Form_Load()
    Dim intLoop As Integer
    lbltotal.Caption = dbltotal
    'Sets values for expiry year control box
    For intLoop = 0 To 10
        cboyear.AddItem Format(Format((Now), "yy") + intLoop, "00")
    Next
    'Sets values for expiry month control box
    For intLoop = 1 To 12
        cboexpirymonth.AddItem Format(intLoop, "00")
    Next
    optvisa.Value = True
End Sub

Private Sub optamericanexpress_Click()
    strCardChosen = "American Express"
End Sub

Private Sub optMastercard_Click()
    strCardChosen = "MasterCard"
End Sub

Private Sub optvisa_Click()
    strCardChosen = "Visa"
End Sub
Private Sub cmdnext_Click()
    'validates all the textboxes are filled except for the phone number as it is optional
    strName = txtname.Text
    strAddress = txtaddress.Text
    strCreditcardnum = txtcardnum.Text
    'validates that the credit card number contains 16 numbers
    If Len(txtcardnum.Text) <> 16 Then
        MsgBox "Please Enter a valid credit card number. It should have 16 digits."
    ElseIf txtcardnum = "" Then 'makes sure the credit card text box is filled
        MsgBox ("Please fill in the credit card number")
    ElseIf IsNumeric(txtcardnum.Text) = False Then
        MsgBox ("Please enter a valid credit card number without any letters or special characters.")
    ElseIf cboexpirymonth.Text = "month" Or cboyear = "year" Then
        MsgBox ("Please enter the expiry date completly")
    ElseIf optamericanexpress.Value = False And optvisa.Value = False And optMastercard.Value = False Then
        MsgBox ("Please select a payment option.")
    ElseIf txtaddress.Text = "" Then
        MsgBox ("Please enter in your address")
    ElseIf txtname.Text = "" Then
        MsgBox ("Please enter in the cardholder's name.")
    Else
        Unload frmCheckout
        frmreceipt.Show
    End If
End Sub

Private Sub txtcardnum_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9 'allows users to only enter numbers
        Case vbKeyBack 'allows the user to press backspace
        Case Else
            KeyAscii = 0 'ignores the input of the users if they type anything other than numbers
    End Select
End Sub


Private Sub txtname_KeyPress(KeyAscii As Integer)
    If Len(txtname.Text) > 30 Then
        KeyAscii = 0        'Ignores continued input
    Else
    End If
    Select Case KeyAscii
        Case vbKeyBack 'allows the user to press backspace
        Case vbKeySpace 'allows the user to press space
        Case vbKeyA To vbKeyZ 'allows them to in upper case letters
        Case vbKeyA + 32 To vbKeyZ + 32 'allows them to type in lower case letters
        Case Else
        KeyAscii = 0 'ignores the input of the users if they type anything else
    End Select
End Sub

Private Sub txtphonenum1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9 'allows users to only enter numbers
        Case vbKeyBack 'allows the user to press backspace
        Case vbKeySpace 'allows the user to press space
        Case Else
            KeyAscii = 0 'ignores the input of the users if they type anything other than numbers
    End Select
    
    If Len(txtphonenum1.Text) > 2 Then
        KeyAscii = 0        'Ignores continued input
    Else
    End If
End Sub

Private Sub txtphonenum2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9 'allows users to only enter numbers
        Case vbKeyBack 'allows the user to press backspace
        Case Else
            KeyAscii = 0 'ignores the input of the users if they type anything other than numbers
    End Select

    If Len(txtphonenum2.Text) > 2 Then
        KeyAscii = 0        'Ignores continued input
    Else
    End If
End Sub

Private Sub txtphonenum3_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9 'allows users to only enter numbers
        Case vbKeyBack 'allows the user to press backspace
        Case vbKeySpace 'allows the user to press space
        Case Else
            KeyAscii = 0 'ignores the input of the users if they type anything other than numbers
    End Select

    If Len(txtphonenum3.Text) > 3 Then
        KeyAscii = 0        'Ignores continued input
    Else
    End If
End Sub
