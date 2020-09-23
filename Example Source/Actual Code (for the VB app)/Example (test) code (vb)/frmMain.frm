VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Example exe for VB"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   3300
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEdit3 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox txtEdit2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox txtEdit1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

