VERSION 5.00
Begin VB.Form frm14DecimalToBinary 
   Caption         =   "Form2"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8220
   LinkTopic       =   "Form2"
   ScaleHeight     =   3300
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Decimal To Binary"
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblResult 
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a Number"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frm14DecimalToBinary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConvert_Click()
Dim n As Integer
Dim binary As String
Dim i As Integer
Dim r As Integer

i = 10

n = Val(txtNumber.Text)

While (n > 0)
    r = n Mod 2
    n = n / 2
    binary = binary & (r * i)
    i = i * 10
Wend

lblResult.Caption = binary

End Sub

Private Sub Form_Load()

End Sub
