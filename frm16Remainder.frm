VERSION 5.00
Begin VB.Form frm16Remainder 
   Caption         =   "Form2"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8310
   LinkTopic       =   "Form2"
   ScaleHeight     =   4170
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutput 
      Height          =   735
      Left            =   2280
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frm16Remainder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
    Dim i As Integer
    i = 7
    
    Do While True
        If (i Mod 2 = 1 & i Mod 3 = 1 & i Mod 4 = 1 & i Mod 5 = 1 & i Mod 6 = 1) Then
            txtOutput.Text = i
            Exit Do
        End If
        i = i + 7
    Loop

End Sub
