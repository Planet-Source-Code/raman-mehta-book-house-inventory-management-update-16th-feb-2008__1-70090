VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmdemand 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demand"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   Icon            =   "frmdemand.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   5415
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Default         =   -1  'True
      DownPicture     =   "frmdemand.frx":08CA
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   5520
      Width           =   5295
      Begin VB.CommandButton cmdmod 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   600
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox txtt_code 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox txtdemand 
         Height          =   300
         Left            =   1230
         MaxLength       =   5
         TabIndex        =   1
         Top             =   495
         Width           =   2055
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Title Code:"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   14
         Top             =   165
         Width           =   765
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Demand:"
         Height          =   195
         Index           =   1
         Left            =   435
         TabIndex        =   13
         Top             =   540
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Sort Records"
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   -120
      Width           =   5295
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Descending"
         Height          =   495
         Index           =   1
         Left            =   4080
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ascending"
         Height          =   495
         Index           =   0
         Left            =   2640
         TabIndex        =   7
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.ComboBox cbosort 
         Height          =   315
         ItemData        =   "frmdemand.frx":0BD4
         Left            =   120
         List            =   "frmdemand.frx":0BE4
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Sort Records By:"
         Height          =   255
         Left            =   645
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmdemand.frx":0C17
         Top             =   75
         Width           =   480
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmdemand.frx":14E1
      Height          =   3375
      Left            =   0
      TabIndex        =   9
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
      DataMember      =   "demand"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "T_CODE"
         Caption         =   "Title Code"
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
         DataField       =   "T_NAME"
         Caption         =   "Title Name"
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
         DataField       =   "STOCK"
         Caption         =   "Current Stock"
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
      BeginProperty Column03 
         DataField       =   "DEMAND"
         Caption         =   "Demand"
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
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmdemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form displays the demand for various titles and
' also enables to change the demand
Dim order As String * 4
Private Sub cbosort_Click()
    With Data1.rsdemand
        .Sort = .Fields(cbosort.ListIndex).Name & " " & order
    End With
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub
Private Sub cmdmod_Click()
    ' show the controls to change the field when the modify button is clicked
    makevisible True
    txtt_code.Text = Data1.rsdemand.Fields("t_code")
    txtdemand.Text = Data1.rsdemand.Fields("demand")
    Me.Caption = "Modify Record"
    txtdemand.SetFocus
End Sub

Private Sub DataGrid1_DblClick()
    If enable = True Then cmdmod_Click
End Sub

Private Sub DataGrid1_Keydown(KeyCode As Integer, Shift As Integer)
    If enable = True Then
        If KeyCode = vbKeyReturn Then
            KeyCode = 0
            cmdmod_Click
        End If
    End If
End Sub

Private Sub Form_Load()
    cmdmod.Enabled = enable
    cbosort.ListIndex = 0
    If Data1.rsdemand.RecordCount = 0 Then enablecontrols False
End Sub


Sub enablecontrols(val As Boolean)
    cmdmod.Enabled = val
    DataGrid1.Enabled = val
    cbosort.Enabled = val
    Option1(0).Enabled = val
    Option1(1).Enabled = val
End Sub
Private Sub cmdcancel_Click()
    makevisible False
    DataGrid1.Enabled = True
    Me.Caption = "Demand"
End Sub

Private Sub cmdsave_Click()
    On Error GoTo err
    Dim bm
    If txtdemand.Text = "" Then
        MsgBox "Demand is blank", vbInformation
        txtdemand.SetFocus
        Exit Sub
    End If
    Data1.conn.BeginTrans
    With Data1.rsdemand
        .Fields("t_code") = txtt_code.Text
        .Fields("demand") = txtdemand.Text
        .Update
        ' if titles form is open then reflect the changes made in the form
        If titlesformisopen Then
            With Data1.rstitles
                Set frmtitles.DataGrid1.DataSource = Nothing
                .Close
                .Open
                Set frmtitles.DataGrid1.DataSource = Data1
                .Find "t_code='" & txtt_code.Text & "'"
            End With
        End If
        makevisible False
        Me.Caption = "Demand"
    End With
    Data1.conn.CommitTrans
    Exit Sub
err:
    Data1.conn.RollbackTrans
    Set DataGrid1.DataSource = Nothing
    Data1.rsdemand.CancelUpdate
    Set DataGrid1.DataSource = Data1
    MsgBox err.Description, vbCritical, "Error"
End Sub
Sub makevisible(vis As Boolean)
    Frame1.Visible = vis
    Frame2.Visible = Not vis
    cmdsave.Visible = vis
    cmdcancel.Visible = vis
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Data1.rsdemand.Close
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0: order = "ASC"
        Case 1: order = "DESC"
    End Select
    cbosort_Click
End Sub

Private Sub txtdemand_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub
