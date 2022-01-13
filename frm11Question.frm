VERSION 5.00
Begin VB.Form frm11Question 
   Caption         =   "Form2"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12885
   LinkTopic       =   "Form2"
   ScaleHeight     =   4605
   ScaleWidth      =   12885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPattern6 
      Caption         =   "Pattern6"
      Height          =   480
      Left            =   10440
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdPattern5 
      Caption         =   "Pattern5"
      Height          =   480
      Left            =   8640
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Pattern 4"
      Height          =   495
      Left            =   6600
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdPattern3 
      Caption         =   "Pattern 3"
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdPattern2 
      Caption         =   "Pattern 2"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdPattern1 
      Caption         =   "Pattern 1"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtNumber 
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblResult 
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   2880
      Width           =   10215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a number"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "frm11Question"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPattern1_Click()
lblResult.Caption = ""
Dim num As Long
num = Val(txtNumber.Text)

For i = 1 To num
    If i Mod 2 = 0 Then
        lblResult.Caption = lblResult.Caption & " " & (i * i)
    End If
    
Next
End Sub

Private Sub cmdPattern2_Click()
lblResult.Caption = ""
Dim num As Long
Dim temp As Long

num = Val(txtNumber.Text)

For index = 1 To num
    If (index Mod 2 = 0) Then
        temp = index * -1
        lblResult.Caption = lblResult.Caption & " " & temp
    Else
        lblResult.Caption = lblResult.Caption & " " & index
    End If
Next
End Sub

Private Sub cmdPattern3_Click()
lblResult.Caption = ""
Dim num As Long
Dim temp As Long

num = Val(txtNumber.Text)

For index = 1 To num
    temp = index ^ Val(index)
    
    lblResult.Caption = lblResult.Caption & " " & temp
Next
End Sub

Private Sub cmdPattern5_Click()
num = Val(txtNumber.Text)
lblResult.Caption = ""
Dim xx As Long
For x = 1 To num
    xx = x ^ 2
    lblResult.Caption = lblResult.Caption & " " & (xx)
Next
End Sub

Private Sub cmdPattern6_Click()
num = Val(txtNumber.Text)
lblResult.Caption = ""
Dim xx As Long
For x = 1 To num
    xx = (3 * (x - 1) ^ 2) + (3 + (-1) ^ x) / 2
    lblResult.Caption = lblResult.Caption & " " & (xx)
Next
End Sub

Private Sub Command4_Click()
Dim n1 As Integer
Dim n2 As Integer
Dim n3 As Integer
Dim n4 As Integer
n1 = 1
n2 = 4
n3 = 7
num = Val(txtNumber.Text)
lblResult.Caption = str(n1) & " " & str(n2) & " " & str(n3)

For x = 3 To num
    n4 = n1 + n2 + n3
    lblResult.Caption = lblResult.Caption & str(n4) & " "
    n1 = n2
    n2 = n3
    n3 = n4
Next

End Sub
