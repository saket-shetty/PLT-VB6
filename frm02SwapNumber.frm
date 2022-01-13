VERSION 5.00
Begin VB.Form frm02SwapNumber 
   Caption         =   "Swap Number"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult2 
      Height          =   500
      Left            =   8160
      TabIndex        =   8
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtResult1 
      Height          =   500
      Left            =   8160
      TabIndex        =   7
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   500
      Left            =   960
      TabIndex        =   5
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdSwap 
      Caption         =   "Swap "
      Height          =   500
      Left            =   3240
      TabIndex        =   2
      Top             =   2880
      WhatsThisHelpID =   1455
      Width           =   1455
   End
   Begin VB.TextBox txtSecondNumber 
      Height          =   500
      Left            =   2760
      TabIndex        =   1
      Top             =   1440
      Width           =   2000
   End
   Begin VB.TextBox txtFirstNumber 
      Height          =   500
      Left            =   2760
      TabIndex        =   0
      Top             =   600
      Width           =   2000
   End
   Begin VB.Label Label3 
      Caption         =   "Result"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Second Number"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Enter First Number"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frm02SwapNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSwap_Click()
Dim num1 As Integer
Dim num2 As Integer

num1 = Val(txtFirstNumber.Text)
num2 = Val(txtSecondNumber.Text)

Dim temp As Integer

temp = num1
num1 = num2
num2 = temp

txtResult1.Text = num1
txtResult2.Text = num2

End Sub

Private Sub cmdView_Click()
Dim num1 As Integer
Dim num2 As Integer

num1 = Val(txtFirstNumber.Text)
num2 = Val(txtSecondNumber.Text)

txtResult1.Text = num1
txtResult2.Text = num2

End Sub

Private Sub Form_Load()

End Sub
