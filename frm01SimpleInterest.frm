VERSION 5.00
Begin VB.Form frm01SimpleInterest 
   Caption         =   "Simple Interest"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   14355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClick 
      Caption         =   "Calculate Simple Interest"
      Height          =   735
      Left            =   5640
      TabIndex        =   4
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox txtResult 
      Height          =   975
      Left            =   5400
      TabIndex        =   3
      Text            =   "Result"
      Top             =   3720
      Width           =   3615
   End
   Begin VB.TextBox txtRate 
      Height          =   855
      Left            =   9360
      TabIndex        =   2
      Text            =   "Rate"
      Top             =   360
      Width           =   3735
   End
   Begin VB.TextBox txtTime 
      Height          =   855
      Left            =   5400
      TabIndex        =   1
      Text            =   "Time"
      Top             =   360
      Width           =   3615
   End
   Begin VB.TextBox txtPrinciple 
      Height          =   855
      Left            =   1800
      TabIndex        =   0
      Text            =   "Principle"
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frm01SimpleInterest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClick_Click()
Dim p As Integer
Dim time As Integer
Dim rate As Integer
Dim si As Integer

p = Val(txtPrinciple.Text)
time = Val(txtTime.Text)
rate = Val(txtRate.Text)

si = ((p * time * rate) / 100)

txtResult.Text = si
End Sub

