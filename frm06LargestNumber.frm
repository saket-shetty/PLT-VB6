VERSION 5.00
Begin VB.Form frm06LargestNumber 
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtNumber3 
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox txtNumber2 
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtNumber1 
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Number 3:"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Number 2:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Number 1:"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Largest and Second Largest Number"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "frm06LargestNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCheck_Click()
Dim number1 As Integer
Dim number2 As Integer
Dim number3 As Integer

number1 = Val(txtNumber1.Text)
number2 = Val(txtNumber2.Text)
number3 = Val(txtNumber3.Text)



If (number1 > number2 And number1 > number3) Then
    If (number2 > number3) Then
        MsgBox number1 & "is largest and " & number2 & " is second largest"
    Else
        MsgBox number1 & "is largest and " & number3 & " is second largest"
    End If
ElseIf (number2 > number1 And number2 > number3) Then
    If (number1 > number3) Then
        MsgBox number2 & "is largest and " & number1 & " is second largest"
    Else
        MsgBox number2 & "is largest and " & number3 & " is second largest"
    End If
ElseIf (number3 > number1 And number3 > number2) Then
    If (number1 > number2) Then
        MsgBox number3 & "is largest and " & number1 & " is second largest"
    Else
        MsgBox number3 & "is largest and " & number2 & " is second largest"
    End If
End If
        
End Sub

