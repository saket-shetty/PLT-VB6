VERSION 5.00
Begin VB.Form frm27MatrixTranspose 
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9885
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
   ScaleHeight     =   4935
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTranspose 
      Caption         =   "Transpose"
      Height          =   360
      Left            =   3600
      TabIndex        =   7
      Top             =   2520
      Width           =   990
   End
   Begin VB.TextBox txtNumber 
      Height          =   1455
      Left            =   5520
      TabIndex        =   6
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdAddNumber 
      Caption         =   "AddNumber"
      Height          =   360
      Left            =   7920
      TabIndex        =   4
      Top             =   1080
      Width           =   990
   End
   Begin VB.TextBox txtColumn 
      Height          =   1455
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtRow 
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblTranspose 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transpose"
      Height          =   195
      Left            =   6120
      TabIndex        =   9
      Top             =   3120
      Width           =   750
   End
   Begin VB.Label lblArray 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Array"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   3120
      Width           =   405
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number"
      Height          =   195
      Left            =   5520
      TabIndex        =   5
      Top             =   240
      Width           =   555
   End
   Begin VB.Label lblColumn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Column"
      Height          =   195
      Left            =   2760
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
Attribute VB_Name = "frm27MatrixTranspose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr(20, 20) As Integer
Dim row As Integer
Dim col As Integer
Dim maxrow As Integer
Dim maxcol As Integer

Private Sub cmdAddNumber_Click()
row = Val(txtRow.Text)
col = Val(txtColumn.Text)

If row > maxrow Then
    maxrow = row
End If

If col > maxcol Then
    maxcol = col
End If

arr(row, col) = Val(txtNumber.Text)
txtRow.Text = ""
txtColumn.Text = ""
txtNumber.Text = ""
printArr
End Sub

Private Sub printArr()
lblArray.Caption = ""
For i = 0 To maxrow
    For j = 0 To maxcol
        lblArray.Caption = lblArray & " " & arr(i, j)
    Next
    lblArray.Caption = lblArray.Caption & vbCrLf
Next
End Sub

Private Sub cmdTranspose_Click()
lblTranspose.Caption = ""
For i = 0 To maxcol
    For j = 0 To maxrow
        lblTranspose.Caption = lblTranspose & " " & arr(j, i)
    Next
    lblTranspose.Caption = lblTranspose.Caption & vbCrLf
Next
End Sub
