VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmusers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   Icon            =   "frmusers.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   4455
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   5340
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Default         =   -1  'True
      DownPicture     =   "frmusers.frx":08CA
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   5340
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   -120
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox txtcpass 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1710
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   870
         Width           =   2655
      End
      Begin VB.ComboBox cborole 
         Height          =   315
         ItemData        =   "frmusers.frx":0BD4
         Left            =   1710
         List            =   "frmusers.frx":0BDE
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1245
         Width           =   2655
      End
      Begin VB.TextBox txtpass 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1710
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   495
         Width           =   2655
      End
      Begin VB.TextBox txtuser 
         Height          =   300
         Left            =   1710
         MaxLength       =   20
         TabIndex        =   0
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password:"
         Height          =   195
         Left            =   255
         TabIndex        =   16
         Top             =   930
         Width           =   1305
      End
      Begin VB.Label lblrole 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Role:"
         Height          =   195
         Left            =   1185
         TabIndex        =   15
         Top             =   1305
         Width           =   375
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Index           =   1
         Left            =   825
         TabIndex        =   14
         Top             =   540
         Width           =   735
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "User Name:"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   13
         Top             =   165
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   -360
      TabIndex        =   11
      Top             =   5340
      Width           =   5295
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdnewmod 
         Caption         =   "&Modify"
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdnewmod 
         Caption         =   "&Add New"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   3720
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmusers.frx":0BFB
      Height          =   3375
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   15920516
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "users"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "NAME"
         Caption         =   "User Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "ROLE"
         Caption         =   "Role"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "PASS"
         Caption         =   "Password"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2039.811
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmusers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form displays the currnt user accounts and the corresponding passord
' enable adding, modifying and deleting users
Dim modstate As Boolean
Dim prevusername As String
Private Sub cmdclose_Click()
    Unload Me
End Sub
Private Sub cmdDelete_Click()
    On Error GoTo err
    If MsgBox("Are you sure you want to delete this User?", vbQuestion + vbYesNo, "Confirm deletion") = vbNo Then
        Exit Sub
    End If
    Data1.conn.BeginTrans
    ' convert both the current user and the user entered in the database and then
    ' compare in order to avoid mismatch of case
    With Data1.rsusers
        If UCase(user) = UCase(.Fields("name")) Then
            MsgBox "Current User cannot be deleted", vbInformation
            Data1.conn.RollbackTrans
            Exit Sub
        End If
        .Delete
        If .RecordCount = 0 Then
            .MoveFirst
            enablecontrols False
        Else
            .MoveNext
            If .EOF Then: .MoveLast
        End If
    End With
    Data1.conn.CommitTrans
    Exit Sub
err:
    Data1.conn.RollbackTrans
    Set DataGrid1.DataSource = Nothing
    Data1.rsusers.CancelUpdate
    Set DataGrid1.DataSource = Data1
    MsgBox err.Description, vbCritical, "Error"
End Sub

Private Sub cmdnewmod_Click(Index As Integer)
    makevisible True
    If (Index = 0) Then
        txtuser.Text = ""
        txtpass.Text = ""
        txtcpass.Text = ""
        cborole.ListIndex = 0
        Me.Caption = "New User"
        DataGrid1.Enabled = False
        cborole.Visible = True
        lblrole.Visible = True
    Else
        txtuser.Text = Data1.rsusers.Fields("name")
        txtpass.Text = Data1.rsusers.Fields("pass")
        txtcpass.Text = txtpass.Text
        cborole.Text = Data1.rsusers.Fields("role")
        prevusername = txtuser.Text
        If UCase(txtuser.Text) = UCase(user) Then
            cborole.Visible = False
            lblrole.Visible = False
        Else
            cborole.Visible = True
            lblrole.Visible = True
        End If
        Me.Caption = "Modify User"
    End If
    modstate = Index
    txtuser.SetFocus
End Sub

Private Sub DataGrid1_DblClick()
    cmdnewmod_Click 1
End Sub

Private Sub DataGrid1_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        cmdnewmod_Click 1
    End If
End Sub

Private Sub Form_Load()
    Data1.rsusers.Sort = "name"
    If Data1.rsusers.RecordCount = 0 Then enablecontrols False
End Sub


Sub enablecontrols(val As Boolean)
    cmdnewmod(1).Enabled = val
    cmdDelete.Enabled = val
    DataGrid1.Enabled = val
End Sub
Private Sub cmdcancel_Click()
    makevisible False
    If Data1.rsusers.RecordCount > 0 Then: DataGrid1.Enabled = True
    Me.Caption = "Users"
End Sub

Private Sub cmdsave_Click()
    On Error GoTo err
    Dim bm
    If txtuser.Text = "" Then
        MsgBox "User Name Name is blank", vbInformation
        txtuser.SetFocus
    ElseIf txtpass.Text = "" Then
        MsgBox "Password is blank", vbInformation
        txtpass.SetFocus
    ElseIf Len(txtpass.Text) < 6 Then
        MsgBox "Password be at least 6 characters long", vbInformation
        txtpass.SelStart = 0
        txtpass.SelLength = Len(txtpass.Text)
        txtpass.SetFocus
    ElseIf txtpass.Text <> txtcpass.Text Then
        MsgBox "Confirmation Password is not identical", vbInformation
        txtcpass.SelStart = 0
        txtcpass.SelLength = Len(txtpass.Text)
        txtcpass.SetFocus
    Else
        Data1.conn.BeginTrans
        With Data1.rsusers
            If (modstate = False) Then: .AddNew
            .Fields("name") = txtuser.Text
            .Fields("pass") = txtpass.Text
            .Fields("role") = cborole.Text
            .Update
            If modstate = False Then
                Set DataGrid1.DataSource = Data1
                .Find .Fields(0).Name & "='" & txtuser.Text & "'"
                If MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    txtuser.Text = ""
                    txtpass.Text = ""
                    cborole.ListIndex = 0
                    txtuser.SetFocus
                Else
                    makevisible False
                    DataGrid1.Enabled = True
                    Me.Caption = "Users"
                    If cmdDelete.Enabled = False Then enablecontrols True
                End If
            Else
                ' if the modified user is the current user then reflect
                ' the fact in the statusbar
                If UCase(prevusername) = UCase(user) Then
                    frmmain.StatusBar1.Panels(1).Text = "Current User:" & txtuser.Text & "(" & role & ")"
                    user = txtuser.Text
                End If
                makevisible False
                Me.Caption = "Users"
            End If
        End With
        Data1.conn.CommitTrans
    End If
    Exit Sub
err:
    Data1.conn.RollbackTrans
    Set DataGrid1.DataSource = Nothing
    Data1.rsusers.CancelUpdate
    Set DataGrid1.DataSource = Data1
    If Data1.rsusers.RecordCount > 0 Then
        DataGrid1.Enabled = True
    End If
    MsgBox err.Description, vbCritical, "Error"
End Sub
Sub makevisible(vis As Boolean)
    Frame1.Visible = vis
    Frame2.Visible = Not vis
    cmdsave.Visible = vis
    cmdcancel.Visible = vis
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Data1.rsusers.Close
End Sub


