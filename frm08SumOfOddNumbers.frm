VERSION 5.00
Begin VB.Form frm08SumOfOddNumbers 
   Caption         =   "Form2"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11565
   LinkTopic       =   "Form2"
   ScaleHeight     =   6660
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label lblResult 
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Number"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frm08SumOfOddNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
Dim num As Integer
Dim sum As Integer

num = Val(txtNumber.Text)

For index = 1 To num
    If (index Mod 2 <> 0) Then
        sum = sum + index
    End If
Next

lblResult.Caption = "Sum of Odd numbers is: " & sum


End Sub
