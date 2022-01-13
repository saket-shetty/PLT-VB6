VERSION 5.00
Begin VB.Form frm20RevString 
   Caption         =   "Form2"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9165
   LinkTopic       =   "Form2"
   ScaleHeight     =   4110
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreverse 
      Caption         =   "Reverse"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtString 
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblResult 
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Enter String"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frm20RevString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdreverse_Click()
Dim str As String
Dim rev As String
    
str = txtString.Text

For i = 1 To Len(str)
    rev = Mid(str, i, 1) & rev
Next

lblResult.Caption = rev
    
End Sub
