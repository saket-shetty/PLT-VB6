VERSION 5.00
Begin VB.Form frm26SortArray 
   Caption         =   "Form1"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9660
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
   ScaleHeight     =   4620
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearchNumber 
      Caption         =   "SearchNumber"
      Height          =   360
      Left            =   5640
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtNumber2 
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CommandButton cmdAddNumber 
      Caption         =   "Add Number"
      Height          =   360
      Left            =   5520
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label lblFinalOutput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FinalOutput"
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   3000
      Width           =   840
   End
   Begin VB.Label lblOutput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output"
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   2040
      Width           =   510
   End
   Begin VB.Label lblSearchNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search Number"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
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
Attribute VB_Name = "frm26SortArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr(30) As Integer
Dim index As Integer

Private Sub cmdAddNumber_Click()
Dim n As Integer
n = Val(txtNumber.Text)
arr(index) = n
txtNumber.Text = ""
index = index + 1
sortArray
printArray
End Sub

Private Sub sortArray()
Dim temp As Integer
For i = 0 To index
    For j = 0 To index
        If arr(i) > arr(j) Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
        End If
    Next
Next
End Sub

Private Sub printArray()
lblOutput.Caption = ""
    For x = 0 To index
        lblOutput.Caption = lblOutput.Caption & " " & arr(x)
    Next
End Sub



Private Sub cmdSearchNumber_Click()
Dim search As Integer
Dim exist As Integer
exist = 0
search = Val(txtNumber2.Text)
For s = 0 To index
    If arr(s) = search Then
        exist = 1
        lblFinalOutput.Caption = "Number Exist"
    End If
Next

If exist = 0 Then
    lblFinalOutput.Caption = "Number Doesnot Exist"
End If
End Sub
