VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcrpayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Payments"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   Icon            =   "frmcrpayments.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Sort Records"
      Height          =   1095
      Left            =   0
      TabIndex        =   29
      Top             =   -120
      Width           =   7455
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
      Begin VB.ComboBox cbosort 
         Height          =   315
         ItemData        =   "frmcrpayments.frx":08CA
         Left            =   120
         List            =   "frmcrpayments.frx":08E0
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Sort Records By:"
         Height          =   255
         Left            =   645
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmcrpayments.frx":093F
         Top             =   75
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      DownPicture     =   "frmcrpayments.frx":1209
      Height          =   375
      Left            =   2580
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
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
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1080
      TabIndex        =   21
      Top             =   6840
      Width           =   5295
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2160
         TabIndex        =   11
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
      Begin VB.CommandButton cmdnewmod 
         Caption         =   "&Add New"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   9
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
      Begin VB.CommandButton cmddelall 
         Caption         =   "Delete A&ll"
         Height          =   375
         Left            =   3240
         TabIndex        =   12
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   7695
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
      Begin VB.TextBox txtdigits 
         Height          =   300
         Left            =   4755
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1800
         Width           =   540
      End
      Begin VB.TextBox txtamount 
         Height          =   300
         Left            =   2880
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtsname 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3015
      End
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
      Begin MSDataListLib.DataCombo cboscode 
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
         Format          =   3801091
         CurrentDate     =   39335
      End
      Begin VB.TextBox txtscode 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Code:"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Balance:"
         Height          =   195
         Left            =   5760
         TabIndex        =   28
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Amount Paid:"
         Height          =   195
         Left            =   2880
         TabIndex        =   27
         Top             =   1560
         Width           =   945
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
         TabIndex        =   26
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Payment Date:"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   1560
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Credit Balance:"
         Height          =   195
         Left            =   4560
         TabIndex        =   24
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Code:"
         Height          =   195
         Left            =   4560
         TabIndex        =   22
         Top             =   120
         Width           =   1035
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmcrpayments.frx":1513
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
      DataMember      =   "c_payments"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "P_CODE"
         Caption         =   "Transaction Code"
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
         DataField       =   "S_CODE"
         Caption         =   "Supplier Code"
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
         DataField       =   "S_NAME"
         Caption         =   "Supplier Name"
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
         DataField       =   "PAY_DATE"
         Caption         =   "Payment Date"
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
         Caption         =   "Amount Paid"
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
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1124.787
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
            ColumnWidth     =   1289.764
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmcrpayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form is responsible for managing credit payments
Dim order As String * 4
Dim modstate As Boolean
Dim rs As ADODB.Recordset
Dim prevamount As Single
Private Sub cboscode_Change()
    ' whenever user selects an option from the combobox
    ' the supplier name and his credit balance are automatically displayed
    If cboscode.Text <> "" Then
        rs.MoveFirst
        rs.Find "s_code='" & cboscode.Text & "'"
        txtsname.Text = rs.Fields("s_name")
        txtcredit.Text = FormatNumber(rs.Fields("credit"), 2)
    Else
        txtsname.Text = ""
        txtcredit.Text = ""
    End If
End Sub
Private Sub cbosort_Click()
    With Data1.rsc_payments
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
    Data1.conn.Execute ("delete from c_payments")
    ' the close and open refreshes the recordset
    Data1.rsc_payments.Close
    Data1.rsc_payments.Open
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
    Data1.conn.Execute "delete from c_payments where p_code='" & Data1.rsc_payments.Fields(0) & "'"
    With Data1.rsc_payments
        If .RecordCount = 1 Then
            ' if before deletion there was only one record then disable the controls on the form
            ' to restrict the user interaction to avoid errors and set the flag to true
            flag = True
            enablecontrols False
        Else
            .MoveNext
            If .EOF Then: .MoveLast
            ' store the bookmark of the current record in variable bm
            ' to retrieve it after refreshing the recordset
            bm = .Bookmark - 1
        End If
        .Close
        .Open
        If Not flag Then
            'if there was more than one record prior to deletion then retrieve the record
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
    cboscode.Visible = True
    txtscode.Visible = False
    If (Index = 0) Then
        ' if add button was clicked then generate the code automatically
        txtt_code.Text = makecode()
        cboscode.Text = ""
        cboscode_Change ' call the change event to cleare the fields
        dtp1.Value = Date
        txtamount.Text = "": txtdigits.Text = ""
        txtbalance.Text = ""
        ' there is no need to store prevvious credit balance of the new tranasaction
        ' it is only required for modification
        prevamount = 0
        Me.Caption = "New Payment"
        DataGrid1.Enabled = False
        cboscode.SetFocus
    Else
        With Data1.rsc_payments
            ' retrieve the particualars of the record in the various controls
            ' to display to the users for modification
            txtt_code.Text = .Fields("p_code")
            cboscode.Visible = False
            txtscode.Visible = True
            txtscode.Text = .Fields("s_code")
            txtsname.Text = .Fields("s_name")
            txtcredit.Text = FormatNumber(.Fields("credit"), 2, vbTrue)
            dtp1.Value = .Fields("pay_date")
            Dim pos As Byte
            ' if there is a fractional part in the amount then display it in two
            ' digits by attaching 0's
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
        ' prevbalance is required for modification to the credit balance of the supplier
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
    If Data1.rsc_payments.RecordCount > 0 Then
        DataGrid1.Enabled = True
        enablecontrols True
    End If
    Me.Caption = "Credit Payments"
End Sub

Private Sub cmdsave_Click()
    On Error GoTo err
    Dim bm
    ' validation begins here
    If modstate = False Then
        If cboscode.Text = "" Then
            MsgBox "Supplier Code is blank", vbInformation
            cboscode.SetFocus
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
    cboscode.Enabled = True
    With Data1.rsc_payments
        ' add button was clicked then make space for adding new record
        If (modstate = False) Then
            .AddNew
            .Fields("s_code") = cboscode.Text
        End If
        .Fields("p_code") = txtt_code.Text
        .Fields("pay_date") = dtp1.Value
        .Fields("credit") = CSng(txtcredit.Text)
        .Fields("amount") = txtamount.Text & "." & txtdigits.Text
        .Update
        ' open the recordset for changing the credit balance of the supplier
        If Data1.rssuppliers.State = adStateClosed Then Data1.rssuppliers.Open
        With Data1.rssuppliers
            ' remove the filter if any
            .Filter = adFilterNone
            .MoveFirst
            .Find "s_code='" & Data1.rsc_payments.Fields("s_code") & "'"
            .Fields("credit") = .Fields("credit") + prevamount - CSng(txtamount.Text & "." & txtdigits.Text)
            .Update
        End With
        Set DataGrid1.DataSource = Nothing
        .Close
        .Open
        Set DataGrid1.DataSource = Data1
        .Find .Fields(0).Name & "='" & txtt_code.Text & "'"
        ' if suppliers form is open then reflect the changes in the form
        If suppliersformisopen Then
            With Data1.rssuppliers
                Set frmsuppliers.DataGrid1.DataSource = Nothing
                .Close
                .Open
                Set frmsuppliers.DataGrid1.DataSource = Data1
                .Find "s_code='" & Data1.rsc_payments.Fields("s_code") & "'"
                ' deal with the filter and remove filter buttons on suppliers form
                frmsuppliers.cmdremfilter_Click
            End With
        Else
            Data1.rssuppliers.Close
        End If
        rs.Requery
        If modstate = False Then
            If MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                txtt_code.Text = makecode()
                cboscode.Text = ""
                dtp1.Value = Date
                txtamount.Text = "": txtdigits.Text = ""
                txtbalance.Text = ""
                txtamount.SetFocus
            Else
                ' show the controls when the operation is complete
                makevisible False
                DataGrid1.Enabled = True
                Me.Caption = "Credit Payments"
                ' if there was no record prior to addition then enable the controls
                If cmdDelete.Enabled = False Then enablecontrols True
            End If
        Else
            makevisible False
            Me.Caption = "Credit Payments"
        End If
    End With
    Data1.conn.CommitTrans
    Exit Sub
err:
    Data1.conn.RollbackTrans
    Set DataGrid1.DataSource = Nothing
    Data1.rsc_payments.CancelUpdate
    Set DataGrid1.DataSource = Data1
    If Data1.rsc_payments.RecordCount > 0 Then
        DataGrid1.Enabled = True
    End If
    MsgBox err.Description, vbCritical, "Error"
End Sub
Private Function makecode() As String
    Dim code As String
    Dim n As Byte
    Dim rs As New ADODB.Recordset
    rs.Open "select max(mid(p_code,2))as maxcode from c_payments", Data1.conn
    If (IsNull(rs("maxcode"))) Then
        code = "P0000001"
    Else
        ' make the code eight digits long by attaching 0's
        n = 7 - Len(rs("maxcode") + 1)
        For i = 1 To n
            code = code & "0"
        Next i
        code = "P" & code & rs("maxcode") + 1
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
    Data1.rsc_payments.Close
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
    rs.Open "select s_code,s_name,credit from suppliers where credit > 0 order by s_code", Data1.conn, adOpenDynamic, adLockOptimistic
    Set cboscode.RowSource = rs
    cboscode.ListField = "s_code"
    cbosort.ListIndex = 0
    ' if there is no record then disable the controls on the form
    If Data1.rsc_payments.RecordCount = 0 Then enablecontrols False
End Sub
Private Sub txtamount_Change()
    ' change the balance according to amount entered
    If txtcredit.Text <> "" Then
        If (txtamount.Text & txtdigits.Text) = "" Then
            txtbalance.Text = ""
        Else
            txtbalance.Text = FormatNumber(CSng(txtcredit.Text) - CSng(txtamount.Text & "." & txtdigits.Text), 2, vbTrue)
        End If
    End If
End Sub
Private Sub txtamount_KeyPress(KeyAscii As Integer)
    ' make sure user only enters numbers
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
