VERSION 5.00
Begin VB.Form frmlogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtpass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1080
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2040
      Width           =   2955
   End
   Begin VB.TextBox txtuser 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1560
      Width           =   2955
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   2100
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   1620
      Width           =   840
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   0
      Picture         =   "frmlogin.frx":08CA
      Top             =   0
      Width           =   4830
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form prompts the user to enter a valid user name and password
' to ensure authorized access to the software
Public flag As Boolean
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdlogin_Click()
    On Error GoTo err
    
    If txtuser.Text = "" Then
        MsgBox "User Name is blank", vbInformation
        txtuser.SetFocus
    ElseIf txtpass.Text = "" Then
        MsgBox "Password is blank", vbInformation
        txtpass.SetFocus
    Else
        Dim rs As New ADODB.Recordset
        rs.Open "select * from users where ucase(name)='" & UCase(txtuser.Text) & "'", Data1.conn, adOpenDynamic, adLockOptimistic
        If rs.EOF Then
            MsgBox "User does not exist. Please enter a valid User Name", vbInformation
            txtuser.SelStart = 0
            txtuser.SelLength = Len(txtuser.Text)
            txtuser.SetFocus
        ElseIf txtpass.Text <> rs.Fields("pass") Then
            MsgBox "Password is incorrect. Please enter a valid Password", vbInformation
            txtpass.SelStart = 0
            txtpass.SelLength = Len(txtpass.Text)
            txtpass.SetFocus
        Else
            ' store the current user and its role
            user = txtuser.Text
            role = rs.Fields("role")
            rs.Close
            flag = True
            Unload Me
            flag = False
            With frmmain
                .Show
                If role = "Administrator" Then
                    enable = True
                Else
                    ' if role is other than administrator then restrict access to certain
                    ' functionalites by indicating this in global variable enable
                    enable = False
                End If
                .mnuusers.Enabled = enable
                .mnunew.Enabled = enable
                .mnucredits.Enabled = enable
                For i = 1 To 10
                    .Toolbar1.Buttons(i).Enabled = True
                Next i
                .StatusBar1.Panels(1).Text = "Current User:  " & user & "(" & role & ")"
            End With
        End If
    End If
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If flag = False Then
        If MsgBox("Are you sure?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Cancel = True
        Else: End
        End If
    End If
End Sub

Private Sub txtuser_LostFocus()
    txtuser.Text = Trim(txtuser.Text)
End Sub
