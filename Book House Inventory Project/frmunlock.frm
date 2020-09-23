VERSION 5.00
Begin VB.Form frmunlock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unlock Application"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmunlock.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdunlock 
      Caption         =   "&Unlock"
      Default         =   -1  'True
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtpass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   960
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   120
      Picture         =   "frmunlock.frx":08CA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter Password:"
      Height          =   195
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   1155
   End
End
Attribute VB_Name = "frmunlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this module shows a modal form which can only be
' unloaded through entering the appropriate password
Private Sub cmdunlock_Click()
    Dim rs As New ADODB.Recordset
    rs.Open "select pass from users where ucase(name)='" & UCase(user) & "'", Data1.conn, adOpenDynamic, adLockOptimistic
    If txtpass.Text = "" Then
        txtpass.SetFocus
    ElseIf txtpass.Text <> rs.Fields("pass") Then
        MsgBox "Wrong Password. Unable to unlock", vbInformation
        txtpass.SelStart = 0
        txtpass.SelLength = Len(txtpass.Text)
        txtpass.SetFocus
    Else
        Unload Me
    End If
End Sub
