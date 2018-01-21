VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   1410
   ClientWidth     =   10095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frasplash 
      Height          =   6690
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   9240
      Begin VB.CommandButton cmdenter 
         Caption         =   "Click here for all your book needs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5760
         TabIndex        =   1
         Top             =   5520
         Width           =   1455
      End
      Begin VB.Label lblinstructions 
         BackColor       =   &H00000000&
         Caption         =   "Press any key or press the button to enter."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   120
         TabIndex        =   7
         Top             =   6000
         Width           =   4815
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   7560
         TabIndex        =   6
         Top             =   5520
         Width           =   1590
      End
      Begin VB.Label lblProductName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Planner"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   7440
         TabIndex        =   5
         Top             =   5160
         Width           =   1755
      End
      Begin VB.Label lbladdressStreet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Mainstreet Unionville"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   285
         Left            =   6720
         TabIndex        =   4
         Top             =   1080
         Width           =   2385
      End
      Begin VB.Label lbladdressnumber 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000012&
         Caption         =   "547"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   360
         Left            =   8280
         TabIndex        =   3
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000008&
         Caption         =   "License to Lorraine Li"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   255
         Left            =   7200
         TabIndex        =   2
         Top             =   120
         Width           =   1935
      End
      Begin VB.Image imgLogo 
         Height          =   6705
         Left            =   -480
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9735
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lorraine
'Splash Screen
'Jan.10, 2014
Option Explicit
'This is the slash screen that displays information about my program

Private Sub cmdenter_Click()
    'brings the to the home page
    Unload frmSplash
    frmhomepage.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'allows the user to enter the program through pressing a key
    Unload frmSplash
    frmhomepage.Show
End Sub

Private Sub Form_Load()
    'changes the background colour
    frmSplash.BackColor = vbBlack
    frasplash.BackColor = vbWhite
End Sub

Private Sub imgLogo_Click()
    'allows the user to enter the program through clicking on the image.
    Unload frmSplash
    frmhomepage.Show
End Sub

