VERSION 5.00
Begin VB.Form LoginDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2340
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtServerName 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "chadk2k"
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblServerName 
      Caption         =   "Server to Connect to"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label lblUserName 
      Caption         =   "User Name You Wish to Use"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "LoginDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

