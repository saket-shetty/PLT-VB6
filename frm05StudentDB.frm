VERSION 5.00
Begin VB.Form frm05StudentDB 
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClickLeft 
      Caption         =   "<<"
      Height          =   495
      Left            =   4680
      TabIndex        =   16
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdRightClick 
      Caption         =   ">>"
      Height          =   495
      Left            =   7080
      TabIndex        =   15
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   5760
      TabIndex        =   14
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtTotal 
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtAverage 
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   3375
      Begin VB.TextBox txtSubject3 
         Height          =   400
         Left            =   1500
         TabIndex        =   8
         Top             =   3360
         Width           =   1000
      End
      Begin VB.TextBox txtSubject2 
         Height          =   400
         Left            =   1500
         TabIndex        =   6
         Top             =   2520
         Width           =   1000
      End
      Begin VB.TextBox txtSubject1 
         Height          =   400
         Left            =   1500
         TabIndex        =   4
         Top             =   1560
         Width           =   1000
      End
      Begin VB.TextBox txtName 
         Height          =   400
         Left            =   1500
         TabIndex        =   2
         Top             =   600
         Width           =   1000
      End
      Begin VB.Label Label4 
         Caption         =   "Subject 3"
         Height          =   255
         Left            =   100
         TabIndex        =   7
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Subject 2"
         Height          =   255
         Left            =   100
         TabIndex        =   5
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Subject 1"
         Height          =   375
         Left            =   100
         TabIndex        =   3
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Student Name"
         Height          =   375
         Left            =   100
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Label lblResult 
      Caption         =   "Result:"
      Height          =   495
      Left            =   4560
      TabIndex        =   13
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Total:"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Average"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "frm05StudentDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type student
    stName As String
    subject1 As Integer
    subject2 As Integer
    subject3 As Integer
    total As Integer
    average As Double
    result As String
End Type

Dim std(20) As student
Dim index As Integer
Dim ci As Integer


Private Sub update(index As Integer)
std(index).stName = txtName.Text
With std(index)
    .subject1 = txtSubject1.Text
    .subject2 = txtSubject2.Text
    .subject3 = txtSubject3.Text
    .total = .subject1 + .subject2 + .subject3
    .average = (.total) / 3
txtTotal.Text = .total
txtAverage.Text = .average

    If (.average >= 65) Then
        lblResult.Caption = "First Class"
    ElseIf (.average > 50) Then
        lblResult.Caption = "Second Class"
    ElseIf (.average >= 35) Then
        lblResult.Caption = "Pass Class"
    Else
        lblResult.Caption = "Fail"
    End If
End With
End Sub



Private Sub cmdClickLeft_Click()
    index = index - 1
    getRecord 0
End Sub

Private Sub getRecord(ndex As Integer)
With std(index)
    txtName.Text = .stName
    txtSubject1.Text = .subject1
    txtSubject2.Text = .subject2
    txtSubject3.Text = .subject3
    txtAverage.Text = .average
    txtTotal.Text = .total
End With

End Sub

Private Sub cmdRightClick_Click()
    index = index + 1
    getRecord 0
End Sub

Private Sub cmdSave_Click()
index = index + 1
update (index)
End Sub

