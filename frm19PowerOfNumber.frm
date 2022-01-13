VERSION 5.00
Begin VB.Form frm19PowerOfNumber 
   Caption         =   "Form2"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9150
   LinkTopic       =   "Form2"
   ScaleHeight     =   3765
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtPower 
      Height          =   405
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtNumber 
      Height          =   405
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblResult 
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Power"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Number"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frm19PowerOfNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
Dim n As Integer
Dim p As Integer

n = Val(txtNumber.Text)
p = Val(txtPower.Text)

Dim res As Integer

res = n ^ p

lblResult.Caption = res

End Sub
