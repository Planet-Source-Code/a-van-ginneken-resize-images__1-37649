VERSION 5.00
Begin VB.Form frmView 
   Caption         =   "View"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFileSize 
      Height          =   285
      Left            =   6120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblFileSize 
      Caption         =   "File size"
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblFileName 
      Caption         =   "File:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   6855
      Left            =   360
      Top             =   480
      Width           =   7935
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
