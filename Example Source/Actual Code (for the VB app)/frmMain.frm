VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "This is our example code"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdretrieve 
      Caption         =   "Retrieve code from all textboxes"
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtEdit3 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox txtEdit2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox txtEdit1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const WM_GETTEXT = &HD

' Now, this should pretty familiar to you (other than the second hwnd param).
' If it doesn't then you should download and read my first window tutorial on PSC
' It is named A Supreme Window Tutorial. (Just search for supreme)
' -- Jaime Muscatelli
' webmaster@jaimemuscatelli.zzn.com

Private Sub cmdretrieve_Click()
Dim lMainHwnd As Long
Dim lEdit1 As Long
Dim lEdit2 As Long
Dim lEdit3 As Long

'Yes, these are buffers/pointers. If you don't know what they are, then read my first tutorial
Dim sEdit1 As String * 256
Dim sEdit2 As String * 256
Dim sEdit3 As String * 256

lMainHwnd = FindWindow("ThunderRT6FormDC", "Example exe for VB")
'The Main window.

lEdit1 = FindWindowEx(lMainHwnd, 0&, "ThunderRT6TextBox", vbNullString)
' The first textbox. Nothing New here

lEdit2 = FindWindowEx(lMainHwnd, lEdit1, "ThunderRT6TextBox", vbNullString)
' Now this is where it gets interesting. See the second HWND param? It has the
' first textbox hwnd in it. Just like I said, include the first textbox hwnd
' so it will skip that window and go to the next window! :-)
lEdit3 = FindWindowEx(lMainHwnd, lEdit2, "ThunderRT6TextBox", vbNullString)
' Same process here, except we included the second textbox (Which has the first in it!)
' This is like a spider chart. The second includes the first, so the third will be found!

' THE CODE BELOW SHOWS WHY I DIDN'T WANT TO USE A VB EXAMPLE EXE!
' SEE, THE VB SYSTEM IS BACKWARDS, SO IT GETS THEM FROM BOTTOM UP (UNLIKE ALL OTHER LANGUAGES)

If lEdit3 Then

SendMessageString lEdit1, WM_GETTEXT, 256, sEdit1
SendMessageString lEdit2, WM_GETTEXT, 256, sEdit2
SendMessageString lEdit3, WM_GETTEXT, 256, sEdit3

'SEE WE HAVE TO NOW FLIP THESE BECAUSE VB'S SYSTEM IS CHEEZY.

txtEdit1.Text = sEdit3
txtEdit2.Text = sEdit2
txtEdit3.Text = sEdit1

MsgBox "All Done"

Else
MsgBox "It isn't finding the window, check and see if the program is running..."
End If


End Sub

Private Sub Form_Load()
MsgBox "Make sure my vb Example program is open! "

End Sub
