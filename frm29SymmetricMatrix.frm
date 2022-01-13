VERSION 5.00
Begin VB.Form frm29SymmetricMatrix 
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10515
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
   ScaleHeight     =   4695
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   360
      Left            =   3360
      TabIndex        =   9
      Top             =   1320
      Width           =   990
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   360
      Left            =   9000
      TabIndex        =   6
      Top             =   600
      Width           =   990
   End
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtColumn 
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtRow 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label lblOutput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   6000
      TabIndex        =   11
      Top             =   1800
      Width           =   45
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
      Height          =   195
      Left            =   6000
      TabIndex        =   10
      Top             =   1320
      Width           =   450
   End
   Begin VB.Label lblMatrix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Matrix"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   450
   End
   Begin VB.Label lblEnteredMatrix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number"
      Height          =   195
      Left            =   6000
      TabIndex        =   4
      Top             =   240
      Width           =   555
   End
   Begin VB.Label lblColumn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Column"
      Height          =   195
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   525
   End
   Begin VB.Label lblRow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Row"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   315
   End
End
Attribute VB_Name = "frm29SymmetricMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim maxrow As Integer
Dim maxcol As Integer
Dim arr(20, 20) As Integer

Private Sub cmdAdd_Click()
Dim r As Integer
Dim c As Integer

r = Val(txtRow.Text)
c = Val(txtColumn.Text)

If r > maxrow Then
    maxrow = r
End If

If c > maxcol Then
    maxcol = c
End If
arr(r, c) = Val(txtNumber.Text)
printArray
End Sub


Private Sub printArray()
lblEnteredMatrix = ""
For x = 0 To maxrow
    For y = 0 To maxcol
        lblEnteredMatrix = lblEnteredMatrix & " " & arr(x, y)
    Next
    lblEnteredMatrix = lblEnteredMatrix & vbCrLf
Next
End Sub

Private Sub cmdCheck_Click()
Dim check As Integer
check = 0
If maxrow <> maxcol Then
    check = 1
    lblOutput.Caption = "Not a symmetric matrix"
Else
    For x = 0 To maxrow
        For y = 0 To maxcol
            If arr(x, y) <> arr(y, x) Then
             check = 1
            End If
        Next
    Next
End If

If check = 1 Then
    lblOutput.Caption = "Not a symmetric marix"
Else
    lblOutput.Caption = "Symmetric Matrix"
End If

End Sub
