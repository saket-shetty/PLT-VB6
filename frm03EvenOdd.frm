VERSION 5.00
Begin VB.Form frm03EvenOdd 
   Caption         =   "Form1"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Display the number if it is even or odd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a number"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "frm03EvenOdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCheck_Click()
Dim number As Integer

number = Val(txtNumber.Text)

If (number Mod 2 = 0) Then
    MsgBox number & " is even number"
Else
    MsgBox number & " is odd number"
End If
End Sub

