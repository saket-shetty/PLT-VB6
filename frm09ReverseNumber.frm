VERSION 5.00
Begin VB.Form frm09ReverseNumber 
   Caption         =   "Form1"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReverse 
      Caption         =   "Reverse"
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a number"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frm09ReverseNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdreverse_Click()
Dim number As Integer
Dim rev As Integer

number = Val(txtNumber.Text)

While number > 0
    rev = (rev * 10) + (number Mod 10)
    number = number / 10
Wend

MsgBox rev

End Sub
