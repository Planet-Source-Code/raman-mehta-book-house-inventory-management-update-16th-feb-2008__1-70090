VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcrreceipts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Receipts"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   Icon            =   "frmcrreceipts.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3900
      TabIndex        =   15
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      DownPicture     =   "frmcrreceipts.frx":08CA
      Height          =   375
      Left            =   2580
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   23
      Top             =   4080
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox txtcredit 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtcname 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtamount 
         Height          =   300
         Left            =   2880
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtdigits 
         Height          =   300
         Left            =   4755
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1800
         Width           =   540
      End
      Begin VB.TextBox txtbalance 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtt_code 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   360
         Width           =   2535
      End
      Begin MSDataListLib.DataCombo cboccode 
         Height          =   315
         Left            =   4560
         TabIndex        =   2
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   ""
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSComCtl2.DTPicker dtp1 
         Height          =   300
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   64946179
         CurrentDate     =   39335
      End
      Begin VB.TextBox txtccode 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Customer Code:"
         Height          =   195
         Left            =   4560
         TabIndex        =   31
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Credit Balance:"
         Height          =   195
         Left            =   4560
         TabIndex        =   29
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Receipt Date:"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   29.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   4605
         TabIndex        =   27
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Amount Recieved"
         Height          =   195
         Left            =   2880
         TabIndex        =   26
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Balance:"
         Height          =   195
         Left            =   5760
         TabIndex        =   25
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Code:"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   120
         Width           =   1305
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1080
      TabIndex        =   22
      Top             =   6840
      Width           =   5295
      Begin VB.CommandButton cmddelall 
         Caption         =   "Delete A&ll"
         Height          =   375
         Left            =   3240
         TabIndex        =   12
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   4320
         TabIndex        =   13
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdnewmod 
         Caption         =   "&Add New"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdnewmod 
         Caption         =   "&Modify"
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   10
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Sort Records"
      Height          =   1095
      Left            =   0
      TabIndex        =   20
      Top             =   -120
      Width           =   7455
      Begin VB.ComboBox cbosort 
         Height          =   315
         ItemData        =   "frmcrreceipts.frx":0BD4
         Left            =   120
         List            =   "frmcrreceipts.frx":0BEA
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   600
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ascending"
         Height          =   495
         Index           =   0
         Left            =   3720
         TabIndex        =   17
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Descending"
         Height          =   495
         Index           =   1
         Left            =   5160
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmcrreceipts.frx":0C4E
         Top             =   75
         Width           =   480
      End
      Begin VB.Label Label8 
         Caption         =   "Sort Records By:"
         Height          =   255
         Left            =   645
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmcrreceipts.frx":1518
      Height          =   2895
      Left            =   0
      TabIndex        =   19
      Top             =   1080
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5106
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
      DataMember      =   "c_receipts"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "R_CODE"
         Caption         =   "Transasction Code"
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
         DataField       =   "C_CODE"
         Caption         =   "Customer Code"
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
         DataField       =   "C_NAME"
         Caption         =   "Customer Name"
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
         DataField       =   "REC_DATE"
         Caption         =   "Receipt Date"
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
      BeginProperty Column04 
         DataField       =   "CREDIT"
         Caption         =   "Credit Balance"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "AMOUNT"
         Caption         =   "Amount Received"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1409.953
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmcrreceipts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form is responsible for managing credit receipts
Dim order As String * 4
Dim modstate As Boolean
Dim rs As ADODB.Recordset
Dim prevamount As Single
Private Sub cboccode_Change()
    ' whenever user selects an option from the combobox
    ' the customer name and his credit balance are automatically displayed
    If cboccode.Text <> "" Then
        rs.MoveFirst
        rs.Find "c_code='" & cboccode.Text & "'"
        txtcname.Text = rs.Fields("c_name")
        txtcredit.Text = FormatNumber(rs.Fields("credit"), 2)
    Else
        txtcname.Text = ""
        txtcredit.Text = ""
    End If
End Sub
Private Sub cbosort_Click()
    With Data1.rsc_receipts
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
    Data1.conn.Execute ("delete from c_receipts")
    ' the close and open refreshes the recordset
    Data1.rsc_receipts.Close
    Data1.rsc_receipts.Open
    Set DataGrid1.DataSource = Data1
    Data1.conn.CommitTrans
    enablecontrols False
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
    Dim bm
    Dim flag As Boolean
    Data1.conn.BeginTrans
    Set DataGrid1.DataSource = Nothing
    Data1.conn.Execute "delete from c_receipts where r_code='" & Data1.rsc_receipts.Fields(0) & "'"
    With Data1.rsc_receipts
        If .RecordCount = 1 Then
            ' if before deletion there was only one record then disable the controls on the form
            ' to restrict the user interaction to avoid errors and set the flag to true
            flag = True
            enablecontrols False
        Else
            .MoveNext
            If .EOF Then: .MoveLast
            bm = .Bookmark - 1
        End If
        .Close
        .Open
        If Not flag Then
            .Bookmark = bm
        End If
    End With
    Set DataGrid1.DataSource = Data1
    Data1.conn.CommitTrans
    Exit Sub
err:
    Data1.conn.RollbackTrans
    Data1.conn.Cancel
    Set DataGrid1.DataSource = Data1
    MsgBox err.Description, vbCritical, "Error"
End Sub

Private Sub cmdnewmod_Click(Index As Integer)
    makevisible True
    cboccode.Visible = True
    txtccode.Visible = False
    If (Index = 0) Then
        ' if add button was clicked then generate the code automatically
        txtt_code.Text = makecode()
        cboccode.Text = ""
        cboccode_Change ' call the change event to cleare the fields
        dtp1.Value = Date
        txtamount.Text = "": txtdigits.Text = ""
        txtbalance.Text = ""
        ' there is no need to store prevvious credit balance of the new tranasaction
        ' it is only required for modification
        prevamount = 0
        Me.Caption = "New Receipt"
        DataGrid1.Enabled = False
        cboccode.SetFocus
    Else
        With Data1.rsc_receipts
            ' retrieve the particualars of the record in the various controls
            ' to display to the users for modification
            txtt_code.Text = .Fields("r_code")
            cboccode.Visible = False
            txtccode.Visible = True
            txtccode.Text = .Fields("c_code")
            txtcname.Text = .Fields("c_name")
            txtcredit.Text = FormatNumber(.Fields("credit"), 2, vbTrue)
            dtp1.Value = .Fields("rec_date")
            Dim pos As Byte
            pos = InStr(.Fields("amount"), ".")
            If pos = 0 Then
                txtamount.Text = .Fields("amount")
                txtdigits.Text = "00"
            Else
                txtamount.Text = Left(.Fields("amount"), pos - 1)
                txtdigits.Text = Mid(.Fields("amount"), pos + 1)
                For i = 1 To 2 - Len(txtdigits.Text)
                    txtdigits = txtdigits & "0"
                Next i
            End If
        End With
        prevamount = CSng(txtamount.Text & "." & txtdigits.Text)
        Me.Caption = "Modify Record"
        txtamount.SetFocus
    End If
    modstate = Index
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
    If Data1.rsc_receipts.RecordCount > 0 Then
        DataGrid1.Enabled = True
        enablecontrols True
    End If
    Me.Caption = "Credit Receipts"
End Sub

Private Sub cmdsave_Click()
    On Error GoTo err
    Dim bm
    If modstate = False Then
        If cboccode.Text = "" Then
            MsgBox "Customer Code is blank", vbInformation
            cboccode.SetFocus
            Exit Sub
        End If
    End If
    If (txtamount & txtdigits.Text) = "" Then
        MsgBox "Amount is blank", vbInformation
        txtamount.SetFocus
        Exit Sub
    End If
    If CSng(txtbalance.Text) < 0 Then
        MsgBox "Amount entered is greater than Credit Balance", vbInformation
        txtamount.SetFocus
        Exit Sub
    End If
    Data1.conn.BeginTrans
    cboccode.Enabled = True
    With Data1.rsc_receipts
        If (modstate = False) Then
            .AddNew
            .Fields("c_code") = cboccode.Text
        End If
        .Fields("r_code") = txtt_code.Text
        .Fields("rec_date") = dtp1.Value
        .Fields("credit") = CSng(txtcredit.Text)
        .Fields("amount") = txtamount.Text & "." & txtdigits.Text
        .Update
        If Data1.rscustomers.State = adStateClosed Then Data1.rscustomers.Open
        With Data1.rscustomers
            ' remove the filter if any
            .Filter = adFilterNone
            .MoveFirst
            .Find "c_code='" & Data1.rsc_receipts.Fields("c_code") & "'"
            .Fields("credit") = .Fields("credit") + prevamount - CSng(txtamount.Text & "." & txtdigits.Text)
            .Update
        End With
        Set DataGrid1.DataSource = Nothing
        .Close
        .Open
        Set DataGrid1.DataSource = Data1
         .Find .Fields(0).Name & "='" & txtt_code.Text & "'"
        ' if customers form is open then reflect the changes in the form
        If customersformisopen Then
            With Data1.rscustomers
                Set frmcustomers.DataGrid1.DataSource = Nothing
                .Close
                .Open
                Set frmcustomers.DataGrid1.DataSource = Data1
                .Find "c_code='" & Data1.rsc_receipts.Fields("c_code") & "'"
                ' deal with the filter and remove filter buttons on suppliers form
                frmcustomers.cmdremfilter_Click
            End With
        Else
            Data1.rscustomers.Close
        End If
        rs.Requery
        If modstate = False Then
            If MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                txtt_code.Text = makecode()
                cboccode.Text = ""
                dtp1.Value = Date
                txtamount.Text = "": txtdigits.Text = ""
                txtbalance.Text = ""
                txtamount.SetFocus
            Else
                makevisible False
                Me.Caption = "Credit Receipts"
                DataGrid1.Enabled = True
                If cmdDelete.Enabled = False Then enablecontrols True
            End If
        Else
            makevisible False
            Me.Caption = "Credit Receipts"
        End If
    End With
    Data1.conn.CommitTrans
    Exit Sub
err:
    Data1.conn.RollbackTrans
    Set DataGrid1.DataSource = Nothing
    Data1.rsc_receipts.CancelUpdate
    Set DataGrid1.DataSource = Data1
    If Data1.rsc_receipts.RecordCount > 0 Then
        DataGrid1.Enabled = True
    End If
    MsgBox err.Description, vbCritical, "Error"
End Sub
Private Function makecode() As String
    Dim code As String
    Dim n As Byte
    Dim rs As New ADODB.Recordset
    rs.Open "select max(mid(r_code,2))as maxcode from c_Receipts", Data1.conn
    If (IsNull(rs("maxcode"))) Then
        code = "R0000001"
    Else
        n = 7 - Len(rs("maxcode") + 1)
        For i = 1 To n
            code = code & "0"
        Next i
        code = "R" & code & rs("maxcode") + 1
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
    Data1.rsc_receipts.Close
    rs.Close
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0: order = "ASC"
        Case 1: order = "DESC"
    End Select
    cbosort_Click
End Sub

Private Sub Form_Load()
    cmdnewmod(0).Enabled = enable
    cmdnewmod(1).Enabled = enable
    cmdDelete.Enabled = enable
    cmddelall.Enabled = enable
    Set rs = New ADODB.Recordset
    rs.Open "select c_code,c_name,credit from customers where credit > 0 order by c_code", Data1.conn, adOpenDynamic, adLockOptimistic
    Set cboccode.RowSource = rs
    cboccode.ListField = "c_code"
    cbosort.ListIndex = 0
    If Data1.rsc_receipts.RecordCount = 0 Then enablecontrols False
End Sub
Private Sub txtamount_Change()
    If txtcredit.Text <> "" Then
        If (txtamount.Text & txtdigits.Text) = "" Then
            txtbalance.Text = ""
        Else
            txtbalance.Text = FormatNumber(CSng(txtcredit.Text) - CSng(txtamount.Text & "." & txtdigits.Text), 2, vbTrue)
        End If
    End If
End Sub
Private Sub txtamount_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub

Private Sub txtcredit_Change()
    txtamount_Change
End Sub

Private Sub txtdigits_Change()
    txtamount_Change
End Sub

Private Sub txtdigits_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub

