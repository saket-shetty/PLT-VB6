VERSION 5.00
Begin VB.Form frm16Remainder 
   Caption         =   "Remainder"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8490
   LinkTopic       =   "Form2"
   ScaleHeight     =   4515
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutput 
      Height          =   1215
      Left            =   1920
      TabIndex        =   1
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click"
      Height          =   1095
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frm16Remainder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClick_Click()
    Dim i As Integer
    Dim n As Integer
    i = 1
    Do While True
        n = 7 * i
        i = i + 1
        If (n Mod 2 = 1 And n Mod 3 = 1 And n Mod 4 = 1 And n Mod 5 = 1 And n Mod 6 = 1) Then
            txtOutput.Text = n
            Exit Do
        End If
    Loop
End Sub

