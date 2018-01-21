VERSION 5.00
Begin VB.Form frmThankYou 
   Caption         =   "Thank You!"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdenter 
      Caption         =   "ENTER"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblmessage 
      Caption         =   "Thank you for shopping at The Book Nook. Please press the enter key to go back to the home page."
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmThankYou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lorraine Li
'Thank you form
'Jan.19,2014
'This form thanks the user for using my program.
Private Sub cmdenter_Click()
    'brings the user back to the home page
    Unload frmThankYou
    frmhomepage.Show
End Sub
