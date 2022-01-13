VERSION 5.00
Begin VB.Form frm15BinaryToDecimal 
   Caption         =   "Form2"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8700
   LinkTopic       =   "Form2"
   ScaleHeight     =   3180
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtNumber 
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblResult 
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a number"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frm15BinaryToDecimal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdConvert_Click()
Dim n As Integer
Dim dec As Integer
Dim power As Integer

power = 0

n = Val(txtNumber.Text)

While n > 0
    If n Mod 10 = 1 Then
        dec = dec + (2 ^ power)
    End If
    n = n / 10
    power = power + 1
Wend

lblResult.Caption = dec
End Sub
