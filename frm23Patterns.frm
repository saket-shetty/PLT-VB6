VERSION 5.00
Begin VB.Form frm23Patterns 
   Caption         =   "Patterns"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10545
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
   ScaleHeight     =   5145
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPattern4 
      Caption         =   "Pattern4"
      Height          =   360
      Left            =   8880
      TabIndex        =   5
      Top             =   120
      Width           =   990
   End
   Begin VB.CommandButton cmdPattern3 
      Caption         =   "Pattern3"
      Height          =   360
      Left            =   7560
      TabIndex        =   4
      Top             =   120
      Width           =   990
   End
   Begin VB.CommandButton cmdPattern2 
      Caption         =   "Pattern2"
      Height          =   360
      Left            =   6240
      TabIndex        =   3
      Top             =   120
      Width           =   990
   End
   Begin VB.CommandButton cmdPattern1 
      Caption         =   "Pattern1"
      Height          =   360
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   990
   End
   Begin VB.TextBox txtNumber 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblOutput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output"
      Height          =   3915
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   9615
   End
   Begin VB.Label lblEnterA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a Number"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1125
   End
End
Attribute VB_Name = "frm23Patterns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim j As Integer
Dim n As Integer
Dim s As String

Private Sub cmdPattern1_Click()
    n = Val(txtNumber.Text)
    s = ""
    For i = 1 To n
        For j = 1 To i
            s = s & str(j)
        Next
        s = s & vbCrLf
    Next
    lblOutput.Caption = s
End Sub

Private Sub cmdPattern2_Click()
    n = Val(txtNumber.Text)
    s = ""
    For i = 1 To n
        For j = 1 To i
            s = s & str(i)
        Next
        s = s & vbCrLf
    Next
    lblOutput.Caption = s
End Sub

Private Sub cmdPattern3_Click()
    Dim count As Integer
    count = 1
    n = Val(txtNumber.Text)
    s = ""
    For i = 1 To n
        For j = 1 To i
            s = s & str(count)
            count = count + 1
        Next
        s = s & vbCrLf
    Next
    lblOutput.Caption = s
End Sub

Private Sub cmdPattern4_Click()
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
End Sub
