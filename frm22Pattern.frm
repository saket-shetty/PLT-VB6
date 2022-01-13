VERSION 5.00
Begin VB.Form frm22Pattern 
   Caption         =   "Patterns"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10050
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   5685
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPattern4 
      Caption         =   "Pattern 4"
      Height          =   360
      Left            =   8160
      TabIndex        =   6
      Top             =   240
      Width           =   990
   End
   Begin VB.CommandButton cmdPattern3 
      Caption         =   "Pattern 3"
      Height          =   360
      Left            =   6960
      TabIndex        =   5
      Top             =   240
      Width           =   990
   End
   Begin VB.CommandButton cmdPattern2 
      Caption         =   "Pattern 2"
      Height          =   360
      Left            =   5760
      TabIndex        =   4
      Top             =   240
      Width           =   990
   End
   Begin VB.CommandButton cmdPattern1 
      Caption         =   "Pattern 1"
      Height          =   360
      Left            =   4560
      TabIndex        =   3
      Top             =   240
      Width           =   990
   End
   Begin VB.TextBox txtNumber 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblOutput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output"
      Height          =   4155
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   9135
   End
   Begin VB.Label lblEnterNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Number"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   990
   End
End
Attribute VB_Name = "frm22Pattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Integer
Dim s As String
Dim i As Integer
Dim j As Integer

Private Sub cmdPattern1_Click()

Dim num1 As Integer
Dim num2 As Integer
Dim num3 As Integer
num1 = 0
num2 = 1
n = Val(txtNumber.Text)
s = str(num2) & vbCrLf
For i = 2 To n
    For j = 1 To i
        If num3 < n Then
            num3 = num1 + num2
            s = s & str(num3)
            num1 = num2
            num2 = num3
        End If
    Next
    s = s & vbCrLf
Next
lblOutput.Caption = s


    n = Val(txtNumber.Text)
    s = ""
    For i = 1 To n
        For j = 1 To 5
            s = s + "*"
        Next
        s = s + vbCrLf
    Next
    lblOutput.Caption = s
End Sub

Private Sub cmdPattern2_Click()
    n = Val(txtNumber.Text)
    s = ""
    For i = 1 To n
        For j = 1 To 5
            s = s + str(i)
        Next
        s = s + vbCrLf
    Next
    lblOutput.Caption = s
End Sub

Private Sub cmdPattern3_Click()
    n = Val(txtNumber.Text)
    s = ""
    For i = 1 To n
        For j = 1 To 5
            s = s + str(j)
        Next
        s = s + vbCrLf
    Next
    lblOutput.Caption = s
End Sub

Private Sub cmdPattern4_Click()
    n = Val(txtNumber.Text)
    s = ""
    For i = 1 To n
        For j = 1 To i
            s = s + "*"
        Next
        s = s + vbCrLf
    Next
    lblOutput.Caption = s
End Sub
