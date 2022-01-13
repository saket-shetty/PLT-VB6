VERSION 5.00
Begin VB.Form frm10NumberToWords 
   Caption         =   "Form2"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15285
   LinkTopic       =   "Form2"
   ScaleHeight     =   5850
   ScaleWidth      =   15285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3720
      TabIndex        =   1
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtNumber 
      Height          =   1095
      Left            =   3600
      TabIndex        =   0
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label lblWord 
      Height          =   855
      Left            =   3720
      TabIndex        =   2
      Top             =   4080
      Width           =   10335
   End
End
Attribute VB_Name = "frm10NumberToWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConvert_Click()
Dim number As Integer
Dim words As String
Dim r As Integer

number = Val(txtNumber.Text)

While number > 0
    r = number Mod 10
    
    Select Case r
        Case 0
            digit = "Zero"
        Case 1
            digit = "One"
        Case 2
            digit = "Two"
        Case 3
            digit = "Three"
        Case 4
            digit = "Four"
        Case 5
            digit = "Five"
        Case 6
            digit = "Six"
        Case 7
            digit = "Seven"
        Case 8
            digit = "Eight"
        Case 9
            digit = "Nine"
    End Select
    
    lblWord.Caption = digit & " " & lblWord.Caption
    number = number / 10
Wend




End Sub
