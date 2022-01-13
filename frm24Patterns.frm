VERSION 5.00
Begin VB.Form frm24Patterns 
   Caption         =   "Patterns"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10725
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
   ScaleHeight     =   5070
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPattern4 
      Caption         =   "Pattern4"
      Height          =   360
      Left            =   8280
      TabIndex        =   5
      Top             =   360
      Width           =   990
   End
   Begin VB.CommandButton cmdPattern3 
      Caption         =   "Pattern3"
      Height          =   360
      Left            =   7080
      TabIndex        =   4
      Top             =   360
      Width           =   990
   End
   Begin VB.CommandButton cmdPattern2 
      Caption         =   "Pattern2"
      Height          =   360
      Left            =   5880
      TabIndex        =   3
      Top             =   360
      Width           =   990
   End
   Begin VB.CommandButton cmdPattern1 
      Caption         =   "Pattern1"
      Height          =   360
      Left            =   4680
      TabIndex        =   2
      Top             =   360
      Width           =   990
   End
   Begin VB.TextBox txtNumber 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblOutput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "output"
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label lblEnterNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Number"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   990
   End
End
Attribute VB_Name = "frm24Patterns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim j As Integer
Dim n As Integer
Dim s As String
Dim x As Long
Dim num As Long

Private Sub cmdPattern1_Click()
s = ""
num = 1
n = Val(txtNumber.Text)
For i = 1 To n
    For j = 1 To i
        x = num * num
        If x Mod 2 = 0 Then
            x = x * -1
        End If
        s = s & " " & str(x)
        num = num + 1
    Next
    s = s & vbCrLf
Next
lblOutput = s

End Sub

Private Sub cmdPattern2_Click()
Dim fact As Integer
fact = 1
s = ""
n = Val(txtNumber.Text)

Dim cot As Integer
cot = 1

For i = 1 To n
    For j = 1 To i
        fact = fact * cot
        s = s & str(fact)
        cot = cot + 1
    Next
    s = s & vbCrLf
Next
lblOutput = s

End Sub

Private Sub cmdPattern3_Click()
n = Val(txtNumber.Text)
s = ""

For i = 1 To n
 For z = i To n
    s = s & "  "
 Next
 For j = 1 To i
    s = s & "*"
 Next
    s = s & vbCrLf
Next
lblOutput = s
End Sub

Private Sub cmdPattern4_Click()
n = Val(txtNumber.Text)
s = ""

Dim scount As Integer

For i = 1 To n Step 2
    scount = n
    For z = i To n
       s = s & " "
    Next
    For p = 1 To i
        s = s & "*"
    Next
    s = s & vbCrLf
Next

For i = n - 2 To 0 Step -2
    For z = i To n
       s = s & " "
    Next
    For p = 1 To i
        s = s & "*"
    Next
    s = s & vbCrLf
Next

lblOutput = s
End Sub
