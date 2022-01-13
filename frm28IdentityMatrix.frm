VERSION 5.00
Begin VB.Form frm28IdentityMatrix 
   Caption         =   "Form2"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12090
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
   ScaleHeight     =   5025
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   360
      Left            =   10080
      TabIndex        =   6
      Top             =   1200
      Width           =   990
   End
   Begin VB.TextBox txtElement 
      Height          =   1935
      Left            =   6720
      TabIndex        =   5
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click"
      Height          =   360
      Left            =   2520
      TabIndex        =   4
      Top             =   3000
      Width           =   990
   End
   Begin VB.TextBox txtColumns 
      Height          =   1935
      Left            =   3600
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtRows 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label lblColumns 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Columns"
      Height          =   195
      Left            =   3720
      TabIndex        =   2
      Top             =   240
      Width           =   600
   End
   Begin VB.Label lblEnterNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rows"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   390
   End
End
Attribute VB_Name = "frm28IdentityMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr(10, 10) As Integer
Dim i As Integer
Dim j As Integer

Private Sub cmdAdd_Click()
Dim x As Integer
Dim y As Integer
Dim value As Integer

value = Val(txtElement.Text)
i = Val(txtRows.Text)
j = Val(txtColumns.Text)
arr(x, y) = value
End Sub

Private Sub cmdClick_Click()
    For x = 0 To i
        For y = 0 To j
                If arr(x, x) = 1 And arr(y, y) = 1 Then
                    MsgBox "Identity Matrix"
                End If
        Next
    Next
End Sub
