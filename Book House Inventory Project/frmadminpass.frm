VERSION 5.00
Begin VB.Form frmadminpass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrator Password"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   Icon            =   "frmadminpass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtcpass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1545
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1920
      Width           =   2955
   End
   Begin VB.TextBox txtuser 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1545
      MaxLength       =   20
      TabIndex        =   0
      Top             =   960
      Width           =   2955
   End
   Begin VB.TextBox txtpass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1545
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   2955
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   945
      TabIndex        =   3
      Top             =   2450
      Width           =   975
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2625
      TabIndex        =   4
      Top             =   2450
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Confirm Password:"
      Height          =   195
      Left            =   135
      TabIndex        =   8
      Top             =   1973
      Width           =   1305
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   80
      Picture         =   "frmadminpass.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   1020
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   705
      TabIndex        =   6
      Top             =   1500
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "You are using this application for the first time. Type the User Name and Password for the Administrator Account."
      Height          =   585
      Left            =   820
      TabIndex        =   5
      Top             =   120
      Width           =   3735
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmadminpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form appears only when the user uses the application the very first
' time during the entire life cycle of the software
Dim flag As Boolean
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdlogin_Click()
    On Error GoTo err
    If txtuser.Text = "" Then
        MsgBox "User Name is blank", vbInformation
        txtuser.SetFocus
    ElseIf txtpass.Text = "" Then
        MsgBox "Password is blank"
        txtpass.SetFocus
    ElseIf Len(txtpass.Text) < 6 Then
        MsgBox "Password must be at least 6 characters long", vbInformation
        txtpass.SelStart = 0
        txtpass.SelLength = Len(txtpass.Text)
        txtpass.SetFocus
    ElseIf txtpass.Text <> txtcpass.Text Then
        MsgBox "Confirmation Password is not identical", vbInformation
        txtcpass.SelStart = 0
        txtcpass.SelLength = Len(txtcpass.Text)
        txtcpass.SetFocus
    Else
        Data1.conn.Execute "insert into users values('" & txtuser.Text & "','Administrator','" & txtpass.Text & "')"
        ' the current user and the role are stored in global
        ' variables user and role
        flag = True
        user = txtuser.Text
        role = "Administrator"
        Unload Me
        flag = False
        enable = True
        frmmain.Show
    End If
    Exit Sub
err:
    Data1.conn.Cancel
    MsgBox err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If flag = False Then
        If MsgBox("Are you sure?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Cancel = True
    End If
End Sub

Private Sub txtuser_LostFocus()
    txtuser.Text = Trim(txtuser.Text)
End Sub
