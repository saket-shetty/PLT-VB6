VERSION 5.00
Begin VB.Form frm17ItemDB 
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11040
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
   ScaleHeight     =   4110
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   360
      Left            =   7320
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<<"
      Height          =   360
      Left            =   6120
      TabIndex        =   12
      Top             =   2280
      Width           =   990
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   ">>"
      Height          =   360
      Left            =   8760
      TabIndex        =   11
      Top             =   2280
      Width           =   990
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   360
      Left            =   7320
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame FraItemForm 
      Caption         =   "ItemForm"
      Height          =   3015
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      Begin VB.TextBox txtPrice 
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtQuantity 
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtDescription 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtCode 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   345
      End
      Begin VB.Label lblQuantity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label lblItemDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Description"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblItemCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   750
      End
   End
   Begin VB.Label lblTotalAmount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount :"
      Height          =   195
      Left            =   6120
      TabIndex        =   9
      Top             =   480
      Width           =   1065
   End
End
Attribute VB_Name = "frm17ItemDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Item
    itemCode As String
    itemDescription As String
    quantity As Long
    price As Long
    totalPrice As Long
End Type

Dim itemArr(30) As Item
Dim index As Integer
Dim ci As Integer
Dim prevTotal As Long

Private Sub cmdClear_Click()
    cmdSave.Enabled = True
    txtCode.Text = ""
    txtDescription.Text = ""
    txtQuantity.Text = ""
    txtPrice.Text = ""
End Sub

Private Sub cmdLeft_Click()
    cmdSave.Enabled = False
    If ci > 0 Then
        ci = ci - 1
    End If
    getData ci
End Sub

Private Sub cmdRight_Click()
    cmdSave.Enabled = False
    If ci < index - 1 Then
        ci = ci + 1
    End If
    getData ci
End Sub

Private Sub cmdSave_Click()
    Dim c As String
    Dim d As String
    Dim q As Integer
    Dim p As Integer
    
    c = txtCode.Text
    d = txtDescription.Text
    q = txtQuantity.Text
    p = txtPrice.Text
    
    With itemArr(index)
        .itemCode = c
        .itemDescription = d
        .quantity = q
        .price = p
        .totalPrice = (q * p)
        lblTotalAmount = "Total Amount: " & (prevTotal + .totalPrice)
        prevTotal = prevTotal + .totalPrice
    End With
    index = index + 1
    ci = index
    
    If prevTotal > 10000 Then
        lblTotalAmount = "Total Amount: " & prevTotal - ((10 * prevTotal) / 100)
    End If
    
    txtCode.Text = ""
    txtDescription.Text = ""
    txtQuantity.Text = ""
    txtPrice.Text = ""
End Sub

Private Sub getData(i As Integer)
    With itemArr(i)
        txtCode = .itemCode
        txtDescription = .itemDescription
        txtQuantity = .quantity
        txtPrice = .price
    End With
End Sub


