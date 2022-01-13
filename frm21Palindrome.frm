VERSION 5.00
Begin VB.Form frm21Palindrome 
   Caption         =   "Form2"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11580
   LinkTopic       =   "Form2"
   ScaleHeight     =   4320
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPalindrome 
      Caption         =   "Check Palindrome"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtString 
      Height          =   735
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblResponse 
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a String"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frm21Palindrome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPalindrome_Click()
Dim str As String
Dim rev As String

str = txtString.Text

For i = 1 To Len(str)
    rev = Mid(str, i, 1) & rev
Next

If rev = str Then
    lblResponse.Caption = "Palindrome"
Else
    lblResponse.Caption = "Not Palindrome"
End If

End Sub

