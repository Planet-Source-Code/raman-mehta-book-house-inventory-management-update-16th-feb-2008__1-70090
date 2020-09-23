VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmsubjects 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subjects"
   ClientHeight    =   5985
   ClientLeft      =   3375
   ClientTop       =   1635
   ClientWidth     =   5415
   Icon            =   "frmsubjects.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmsubjects.frx":08CA
      Height          =   3375
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   5415
      _ExtentX        =   9551
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
         Size            =   9.75
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
      DataMember      =   "subjects"
      Caption         =   "Current Subjects"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "S_CODE"
         Caption         =   "Subject Code"
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
         DataField       =   "S_NAME"
         Caption         =   "Subject Name"
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
            ColumnWidth     =   1980.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3135.118
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Sort Records"
      Height          =   975
      Left            =   0
      TabIndex        =   17
      Top             =   -120
      Width           =   5295
      Begin VB.ComboBox cbosort 
         Height          =   315
         ItemData        =   "frmsubjects.frx":08DE
         Left            =   120
         List            =   "frmsubjects.frx":08E8
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ascending"
         Height          =   495
         Index           =   0
         Left            =   2640
         TabIndex        =   10
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Descending"
         Height          =   495
         Index           =   1
         Left            =   4080
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmsubjects.frx":0908
         Top             =   75
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Sort Records By:"
         Height          =   255
         Left            =   645
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Default         =   -1  'True
      DownPicture     =   "frmsubjects.frx":11D2
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txts_name 
         Height          =   300
         Left            =   1230
         MaxLength       =   40
         TabIndex        =   1
         Top             =   495
         Width           =   3375
      End
      Begin VB.TextBox txts_code 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Subject Name:"
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   15
         Top             =   540
         Width           =   1050
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Subject Code:"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   14
         Top             =   165
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   60
      TabIndex        =   16
      Top             =   5520
      Width           =   5295
      Begin VB.CommandButton cmddelall 
         Caption         =   "Delete A&ll"
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   4320
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdnewmod 
         Caption         =   "&Add"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdnewmod 
         Caption         =   "&Modify"
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmsubjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim order As String * 4
Dim modstate As Boolean
Private Sub cbosort_Click()
    With Data1.rssubjects
        .Sort = .Fields(cbosort.ListIndex).Name & " " & order
    End With
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmddelall_Click()
    On Error GoTo err
    If MsgBox("Are you sure you want to delete all the records?", vbQuestion + vbYesNo, "Confirm deletion") = vbNo Then
        Exit Sub
    End If
    Data1.conn.BeginTrans
    Set DataGrid1.DataSource = Nothing
    Data1.conn.Execute ("delete from subjects")
    Data1.rssubjects.Close
    Data1.rssubjects.Open
    Set DataGrid1.DataSource = Data1
    Data1.conn.CommitTrans
    enablecontrols False
    If titleaddmodformisopen Then
        Set frmtitle_addmod.cbosub.RowSource = Data1
        frmtitle_addmod.cbosub.Text = ""
    End If
    Exit Sub
err:
    Data1.conn.RollbackTrans
    Set DataGrid1.DataSource = Nothing
    Data1.conn.Cancel
    Set DataGrid1.DataSource = Data1
    MsgBox err.Description, vbCritical, "Error"
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo err
    If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Confirm deletion") = vbNo Then
        Exit Sub
    End If
    Data1.conn.BeginTrans
    With Data1.rssubjects
        .Delete
        If .RecordCount = 0 Then
            .MoveFirst
            enablecontrols False
        Else
            .MoveNext
            If .EOF Then: .MoveLast
        End If
        
    End With
    ' if titleaddform is open then reflect the changes in the combo box displayed in title
    ' addform for selection of subject category
    If titleaddmodformisopen Then
        Set frmtitle_addmod.cbosub.RowSource = Data1
        frmtitle_addmod.cbosub.Text = ""
    End If
    Data1.conn.CommitTrans
    Exit Sub
err:
    Data1.conn.RollbackTrans
    Set DataGrid1.DataSource = Nothing
    Data1.rssubjects.CancelUpdate
    Set DataGrid1.DataSource = Data1
    MsgBox err.Description, vbCritical, "Error"
End Sub

Private Sub cmdnewmod_Click(Index As Integer)
    makevisible True
    If (Index = 0) Then
        txts_code.Text = makecode()
        txts_name.Text = ""
        Me.Caption = "New Subject"
        DataGrid1.Enabled = False
    Else
        txts_code.Text = Data1.rssubjects.Fields("s_code")
        txts_name.Text = Data1.rssubjects.Fields("s_name")
        Me.Caption = "Modify Subject"
    End If
    modstate = Index
    txts_name.SetFocus
End Sub

Private Sub DataGrid1_DblClick()
    If enable = True Then cmdnewmod_Click 1
End Sub

Private Sub DataGrid1_Keydown(KeyCode As Integer, Shift As Integer)
    If enable = True Then
        If KeyCode = vbKeyReturn Then
            KeyCode = 0
            cmdnewmod_Click 1
        End If
    End If
End Sub

Private Sub Form_Load()
    cmdnewmod(0).Enabled = enable
    cmdnewmod(1).Enabled = enable
    cmdDelete.Enabled = enable
    cmddelall.Enabled = enable
    cbosort.ListIndex = 0
    If Data1.rssubjects.RecordCount = 0 Then enablecontrols False
End Sub


Sub enablecontrols(val As Boolean)
    cmdnewmod(1).Enabled = val
    cmdDelete.Enabled = val
    cmddelall.Enabled = val
    DataGrid1.Enabled = val
    cbosort.Enabled = val
    Option1(0).Enabled = val
    Option1(1).Enabled = val
End Sub
Private Sub cmdcancel_Click()
    makevisible False
    If Data1.rssubjects.RecordCount > 0 Then
        DataGrid1.Enabled = True
        enablecontrols True
    End If
    Me.Caption = "Subjects"
End Sub

Private Sub cmdsave_Click()
    On Error GoTo err
    Dim bm
    If txts_name.Text = "" Then
        MsgBox "Subject Name is blank", vbInformation
        txts_name.SetFocus
        Exit Sub
    End If
    Data1.conn.BeginTrans
    With Data1.rssubjects
        If (modstate = False) Then: .AddNew
        .Fields("s_code") = txts_code.Text
        .Fields("s_name") = txts_name.Text
        .Update
        ' if titles form is open then reflect the changes in the form
        If titlesformisopen Then
            With Data1.rstitles
                If .RecordCount > 0 Then bm = .Bookmark
                Set frmtitles.DataGrid1.DataSource = Nothing
                .Close
                .Open
                Set frmtitles.DataGrid1.DataSource = Data1
                If .RecordCount > 0 Then .Bookmark = bm
            End With
        End If
        ' if titleaddform is open then reflect the changes in the combo box displayed in title
        ' addform for selection of subject category
        If titleaddmodformisopen Then
            Set frmtitle_addmod.cbosub.RowSource = Data1
            frmtitle_addmod.cbosub.Text = ""
        End If
        If modstate = False Then
            Set DataGrid1.DataSource = Data1
            .Find .Fields(0).Name & "='" & txts_code.Text & "'"
            If MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                txts_name.Text = ""
                txts_name.SetFocus
                txts_code = makecode()
            Else
                makevisible False
                DataGrid1.Enabled = True
                Me.Caption = "Subjects"
                If cmdDelete.Enabled = False Then enablecontrols True
            End If
        Else
            makevisible False
            Me.Caption = "Subjects"
        End If
    End With
    Data1.conn.CommitTrans
    Exit Sub
err:
    Data1.conn.RollbackTrans
    Set DataGrid1.DataSource = Nothing
    Data1.rssubjects.CancelUpdate
    Set DataGrid1.DataSource = Data1
    If Data1.rssubjects.RecordCount > 0 Then
        DataGrid1.Enabled = True
    End If
    MsgBox err.Description, vbCritical, "Error"
End Sub
Private Function makecode() As String
    Dim code As String
    Dim n As Byte
    Dim rs As New ADODB.Recordset
    rs.Open "select max(mid(s_code,2))as maxcode from subjects", Data1.conn
    If (IsNull(rs("maxcode"))) Then
        code = "S0000001"
    Else
        n = 7 - Len(rs("maxcode") + 1)
        For i = 1 To n
            code = code & "0"
        Next i
        code = "s" & code & rs("maxcode") + 1
    End If
    makecode = code
    rs.Close
End Function

Sub makevisible(vis As Boolean)
    Frame1.Visible = vis
    Frame2.Visible = Not vis
    cmdsave.Visible = vis
    cmdcancel.Visible = vis
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Data1.rssubjects.Close
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0: order = "ASC"
        Case 1: order = "DESC"
    End Select
    cbosort_Click
End Sub
