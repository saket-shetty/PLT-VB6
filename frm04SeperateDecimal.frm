VERSION 5.00
Begin VB.Form frm04SeperateDecimal 
   Caption         =   "Seperate Decimal"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNumber3 
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtNumber2 
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "Split"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtNumber1 
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frm04SeperateDecimal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSplit_Click()
Dim num As Double

num = Val(txtNumber1.Text)

txtNumber2.Text = Int(num)

txtNumber3.Text = num - Int(num)

End Sub

