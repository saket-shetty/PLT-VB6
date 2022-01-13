VERSION 5.00
Begin VB.Form frm18Sequence 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPattern4 
      Caption         =   "Pattern4"
      Height          =   480
      Left            =   5400
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdPattern3 
      Caption         =   "Pattern3"
      Height          =   480
      Left            =   3720
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdPattern2 
      Caption         =   "Pattern2"
      Height          =   480
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdPattern1 
      Caption         =   "Pattern1"
      Height          =   480
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblOutput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   510
   End
   Begin VB.Label lblEnterNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Number"
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frm18Sequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer

Private Sub cmdPattern1_Click()
Dim prev As Integer
Dim num As Integer
lblOutput = ""
prev = 1
n = Val(txtNumber.Text)
Dim sign As Integer

sign = 1

For x = 0 To n
    num = (prev + (x ^ 2))
    prev = num
    lblOutput = lblOutput & " " & str(num * sign)
    sign = sign * -1
Next
End Sub

Private Sub cmdPattern2_Click()
Dim num1 As Integer
Dim num2 As Integer
Dim num3 As Integer

num1 = 1
num2 = 1

n = Val(txtNumber.Text)

lblOutput = str(num1) & " " & str(num2)

For x = 2 To n
    num3 = num1 + num2
    lblOutput = lblOutput & " " & str(num3)
    num1 = num2
    num2 = num3
Next
End Sub

Private Sub cmdPattern3_Click()
Dim p As Integer
Dim neg As Integer

p = 1
neg = 2

lblOutput = str(p) & " " & str(neg * -1)

n = Val(txtNumber.Text)

For x = 1 To (n / 2)
    p = p + 3
    neg = neg + 4
    lblOutput = lblOutput & " " & str(p) & " " & str(neg * -1)
Next

End Sub

Private Sub cmdPattern4_Click()
Dim num1 As Integer
Dim num2 As Integer
Dim num3 As Integer
Dim num4 As Integer

num1 = 1
num2 = 5
num3 = 8

n = Val(txtNumber.Text)

lblOutput = num1 & " " & num2 & " " & num3

For x = 3 To n
    num4 = num1 + num2 + num3
    lblOutput = lblOutput & " " & num4
    num1 = num2
    num2 = num3
    num3 = num4
Next

End Sub
