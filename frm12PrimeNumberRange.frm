VERSION 5.00
Begin VB.Form frm12PrimeNumberRange 
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10185
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
   ScaleHeight     =   4665
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrimeNumber 
      Caption         =   "Prime Number"
      Height          =   360
      Left            =   3000
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtNumber2 
      Height          =   285
      Left            =   5520
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox txtNumber1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblOutput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "output"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label lblEnterNumber2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Number 2"
      Height          =   195
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   1125
   End
   Begin VB.Label lblEnterNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Number 1"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1125
   End
End
Attribute VB_Name = "frm12PrimeNumberRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrimeNumber_Click()
Dim i As Integer
Dim j As Integer
Dim p As Integer

i = Val(txtNumber1.Text)
j = Val(txtNumber2.Text)

For x = i To j
    p = 0
    For y = 2 To x / 2
        If x Mod y = 0 Then
         p = 1
        End If
    Next
    If p = 0 Then
        lblOutput.Caption = lblOutput.Caption & " " & x
    End If
Next
End Sub

