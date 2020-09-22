VERSION 5.00
Begin VB.Form frmHelp 
   AutoRedraw      =   -1  'True
   Caption         =   "Tint Shop - Help"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8070
   LinkTopic       =   "Form2"
   ScaleHeight     =   287
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   538
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   3975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmHelp.frx":0000
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Text1.Text = "1. Load a picture" + vbCrLf
    Text1.Text = Text1.Text + "2. Choose the number of groups" + vbCrLf
    Text1.Text = Text1.Text + "3. Choose which groups in listbox then click go" + vbCrLf
    Text1.Text = Text1.Text + vbCrLf
    Text1.Text = Text1.Text + "This program breaks the colors of a picture into groups" + vbCrLf
    Text1.Text = Text1.Text + "The groups you choose are painted in color, the others are gray scaled" + vbCrLf
    Text1.Text = Text1.Text + "Isaw a program Tint and just wanted to see if I could build the same effect" + vbCrLf
    
End Sub
