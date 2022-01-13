VERSION 5.00
Begin VB.Form frm13Factorial 
   Caption         =   "Form2"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11415
   LinkTopic       =   "Form2"
   ScaleHeight     =   4980
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   855
      Left            =   1680
      TabIndex        =   1
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox txtNumber 
      Height          =   1095
      Left            =   1680
      TabIndex        =   0
      Text            =   "Enter a number"
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label lblResult 
      Height          =   735
      Left            =   1680
      TabIndex        =   2
      Top             =   3720
      Width           =   4335
   End
End
Attribute VB_Name = "frm13Factorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
Dim n As Integer
Dim fact As Integer
fact = 1

n = Val(txtNumber.Text)
For index = 1 To n
    fact = fact * index
Next

lblResult.Caption = fact

End Sub
