VERSION 5.00
Begin VB.Form frm25LinearSearch 
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9990
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
   ScaleHeight     =   4590
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNumber2 
      Height          =   405
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   360
      Left            =   4560
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdEnterNumber 
      Caption         =   "Enter Number"
      Height          =   360
      Left            =   4560
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtNumber 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblOutput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output"
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Top             =   2280
      Width           =   510
   End
   Begin VB.Label lblSearchNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search Number"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblEnterNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Number"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   990
   End
End
Attribute VB_Name = "frm25LinearSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr(20) As Integer
Dim index As Integer

Private Sub cmdEnterNumber_Click()
Dim n As Integer
n = Val(txtNumber.Text)
arr(index) = n
index = index + 1
txtNumber.Text = ""
End Sub

Private Sub cmdSearch_Click()
Dim search As Integer
Dim exist As Integer
exist = 0
search = Val(txtNumber2.Text)
For s = 0 To index
    If arr(s) = search Then
        exist = 1
        lblOutput.Caption = "Number Exist"
    End If
Next

If exist = 0 Then
    lblOutput.Caption = "Number Doesnot Exist"
End If

End Sub
