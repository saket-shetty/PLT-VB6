VERSION 5.00
Begin VB.Form frm07Employee 
   Caption         =   "Employee Data"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11700
   LinkTopic       =   "Form2"
   ScaleHeight     =   6135
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLeftClick 
      Caption         =   "<<"
      Height          =   375
      Left            =   6720
      TabIndex        =   23
      Top             =   3960
      Width           =   800
   End
   Begin VB.Frame Frame2 
      Caption         =   "Controls"
      Height          =   2175
      Left            =   6360
      TabIndex        =   20
      Top             =   3240
      Width           =   4695
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   3720
         TabIndex        =   24
         Top             =   720
         Width           =   800
      End
      Begin VB.CommandButton cmdRightClick 
         Caption         =   ">>"
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Top             =   720
         Width           =   800
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   1440
         TabIndex        =   21
         Top             =   720
         Width           =   800
      End
      Begin VB.Label lblEmpIndex 
         Caption         =   "Emp Index"
         Height          =   495
         Left            =   480
         TabIndex        =   25
         Top             =   1560
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Salary"
      Height          =   2775
      Left            =   6360
      TabIndex        =   13
      Top             =   240
      Width           =   4695
      Begin VB.TextBox txtNetAnnualSalary 
         Height          =   400
         Left            =   1800
         TabIndex        =   19
         Top             =   2040
         Width           =   2500
      End
      Begin VB.TextBox txtAnnualSalary 
         Height          =   400
         Left            =   1800
         TabIndex        =   17
         Top             =   1320
         Width           =   2500
      End
      Begin VB.TextBox txtGrossSalary 
         Height          =   400
         Left            =   1800
         TabIndex        =   15
         Top             =   600
         Width           =   2500
      End
      Begin VB.Label Label9 
         Caption         =   "Annual Net Salary"
         Height          =   495
         Left            =   200
         TabIndex        =   18
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Annual Salary"
         Height          =   495
         Left            =   195
         TabIndex        =   16
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Gross Salary"
         Height          =   375
         Left            =   195
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame Frma 
      Caption         =   "Employee"
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      Begin VB.TextBox txtMonthlyTax 
         Height          =   400
         Left            =   2000
         TabIndex        =   12
         Top             =   4320
         Width           =   2500
      End
      Begin VB.TextBox txtBonusPercentage 
         Height          =   400
         Left            =   2000
         TabIndex        =   10
         Top             =   3600
         Width           =   2500
      End
      Begin VB.TextBox txtSpecialAllowance 
         Height          =   400
         Left            =   2000
         TabIndex        =   8
         Top             =   2760
         Width           =   2500
      End
      Begin VB.TextBox txtBasic 
         Height          =   400
         Left            =   2000
         TabIndex        =   6
         Top             =   2040
         Width           =   2500
      End
      Begin VB.TextBox txtEmpid 
         Height          =   400
         Left            =   2000
         TabIndex        =   4
         Top             =   1320
         Width           =   2500
      End
      Begin VB.TextBox txtName 
         Height          =   400
         Left            =   2000
         TabIndex        =   2
         Top             =   600
         Width           =   2500
      End
      Begin VB.Label Label6 
         Caption         =   "Monthly Tax Saving Investment"
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Bonus Percentage"
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Special allowance"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Basic"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Employee ID"
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm07Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type EmployeeData
    name As String
    id As String
    basic As Long
    specialAllowance As Long
    bonus As Long
    taxSaving As Long
    grossSalary As Long
    annualSalary As Long
    annualNetSalary As Long
End Type


Dim e(20) As EmployeeData
Dim index As Integer
Dim ci As Integer
Dim tax As Long


Private Sub cmdClear_Click()
    With e(ci)
        .name = ""
        .basic = 0
        .bonus = 0
        .id = ""
        .specialAllowance = 0
        .taxSaving = 0
        .annualNetSalary = 0
        .annualSalary = 0
        .grossSalary = 0
    End With
    txtAnnualSalary.Text = ""
    txtBasic.Text = ""
    txtBonusPercentage = ""
    txtEmpid = ""
    txtGrossSalary = ""
    txtMonthlyTax = ""
    txtName = ""
    txtNetAnnualSalary = ""
    txtSpecialAllowance = ""
    cmdSave.Enabled = True
End Sub

Private Sub cmdLeftClick_Click()
cmdSave.Enabled = False
    If ci >= 0 Then
        ci = ci - 1
    End If
    lblEmpIndex.Caption = ci
    getData ci
End Sub

Private Sub cmdRightClick_Click()
cmdSave.Enabled = False
    If ci >= 0 Then
        ci = ci + 1
    End If
    lblEmpIndex.Caption = "The Employee Index is " & ci
    getData ci
End Sub

Private Sub cmdSave_Click()
index = index + 1
ci = index
lblEmpIndex.Caption = "The Employee Index is " & ci
createData index
End Sub

Private Sub createData(index As Integer)
    With e(index)
        .name = txtName.Text
        .basic = txtBasic.Text
        .bonus = Val(txtBonusPercentage.Text)
        .id = Val(txtEmpid.Text)
        .specialAllowance = Val(txtSpecialAllowance.Text)
        .taxSaving = Val(txtMonthlyTax.Text)
        .grossSalary = Val(.basic + .specialAllowance)
        .annualSalary = .grossSalary + ((.bonus * .basic) / 100)
        txtGrossSalary.Text = .grossSalary
        tax = (.annualSalary - .taxSaving)
        If (tax > 100000 & tax < 150000) Then
            .annualNetSalary = .annualSalary - ((20 * .annualSalary) / 100)
        ElseIf (tax > 150000) Then
            .annualNetSalary = .annualSalary - ((30 * .annualSalary) / 100)
        End If
        txtAnnualSalary.Text = .annualSalary
        txtNetAnnualSalary.Text = .annualNetSalary
    End With
End Sub

Private Sub getData(index As Integer)
    With e(index)
        txtName.Text = .name
        txtAnnualSalary.Text = .annualSalary
        txtBasic.Text = .basic
        txtBonusPercentage.Text = .bonus
        txtEmpid.Text = .id
        txtGrossSalary.Text = .grossSalary
        txtMonthlyTax.Text = .taxSaving
        txtNetAnnualSalary = .annualNetSalary
        txtSpecialAllowance = .specialAllowance
    End With
End Sub

Private Sub Form_Load()
    cmdSave.Enabled = False
    cmdLeftClick.Enabled = False
    cmdRightClick.Enabled = False
    cmdClear.Enabled = False
End Sub

Private Sub FormValidation(i As Integer)
    'If Not IsNull(txtName.Text) & Not IsNull(txtEmpid.Text) & Not IsNull(txtBasic) & Not IsNull(txtBonusPercentage) & Not IsNull(txtSpecialAllowance) & Not IsNull(txtMonthlyTax) Then
    If Not IsNull(txtName.Text) Then
        cmdSave.Enabled = True
        cmdClear.Enable = True
    Else
        cmdSave.Enabled = False
        cmdClear.Enabled = False
    End If
End Sub

Private Sub txtBasic_Change()
FormValidation 0
End Sub

Private Sub txtBonusPercentage_Change()
FormValidation 0
End Sub

Private Sub txtEmpid_Change()
FormValidation 0
End Sub

Private Sub txtMonthlyTax_Change()
FormValidation 0
End Sub

Private Sub txtName_Change()
    FormValidation 0
End Sub

Private Sub txtSpecialAllowance_Change()
FormValidation 0
End Sub
