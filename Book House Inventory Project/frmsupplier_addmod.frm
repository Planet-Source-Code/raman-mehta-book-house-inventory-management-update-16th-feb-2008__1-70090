VERSION 5.00
Begin VB.Form frmsupplier_addmod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Supplier"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   Icon            =   "frmsupplier_addmod.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5730
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtemail_addr 
      Height          =   300
      Left            =   1680
      MaxLength       =   60
      TabIndex        =   5
      Top             =   2520
      Width           =   3975
   End
   Begin VB.TextBox txtfax_no 
      Height          =   300
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   4
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox txts_code 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   3975
   End
   Begin VB.TextBox txts_name 
      Height          =   300
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox txts_addr 
      Height          =   300
      Left            =   1680
      MaxLength       =   80
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txts_cont_no 
      Height          =   300
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(Fields marked * are mandatory)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1680
      TabIndex        =   15
      Top             =   2880
      Width           =   2385
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Email Address(s):"
      Height          =   195
      Index           =   5
      Left            =   375
      TabIndex        =   14
      Top             =   2580
      Width           =   1200
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fax Number(s):"
      Height          =   195
      Index           =   4
      Left            =   510
      TabIndex        =   13
      Top             =   2100
      Width           =   1065
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Supplie Code:"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   12
      Top             =   180
      Width           =   990
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "*Name:"
      Height          =   195
      Index           =   1
      Left            =   1050
      TabIndex        =   11
      Top             =   660
      Width           =   525
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "*Address:"
      Height          =   195
      Index           =   2
      Left            =   900
      TabIndex        =   10
      Top             =   1140
      Width           =   675
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "*Contact Number(s):"
      Height          =   195
      Index           =   3
      Left            =   150
      TabIndex        =   9
      Top             =   1620
      Width           =   1425
   End
End
Attribute VB_Name = "frmsupplier_addmod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form handles the actual addition or modifications of the supplier records
Public modstate As Boolean
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdsave_Click()
    On Error GoTo err
    Dim bm
    If validate() = False Then Exit Sub
    Data1.conn.BeginTrans
    With Data1.rssuppliers
        If (.State = adStateClosed) Then: .Open
        If (modstate = False) Then
            .AddNew
            .Fields("credit") = 0
        End If
        .Fields("s_code") = txts_code.Text
        .Fields("s_name") = txts_name.Text
        .Fields("s_addr") = txts_addr.Text
        .Fields("s_cont_no") = txts_cont_no.Text
        If txtfax_no.Text = "" Then
            .Fields("fax_no") = Null
        Else
            .Fields("fax_no") = txtfax_no.Text
        End If
        If txtemail_addr.Text = "" Then
            .Fields("email_addr") = Null
        Else
            .Fields("email_addr") = txtemail_addr.Text
        End If
        .Update
        ' if purchases form is open then reflect the changes in the form for customers
        If purchasesformisopen Then
            With Data1.rspurchases
                If .RecordCount > 0 Then bm = .Bookmark
                Set frmpurchases.DataGrid1.DataSource = Nothing
                .Close
                .Open
                Set frmpurchases.DataGrid1.DataSource = Data1
                If .RecordCount > 0 Then .Bookmark = bm
            End With
            With Data1.rst_purchases
                If .RecordCount > 0 Then bm = .Bookmark
                Set frmpurchases.DataGrid2.DataSource = Nothing
                .Close
                .Open
                Set frmpurchases.DataGrid2.DataSource = Data1
                If .RecordCount > 0 Then .Bookmark = bm
            End With
        End If
        If modstate = False Then
            ' if the add supplier form was called from suppliers form then
            ' deal with the records displayed on the form
            If suppliersformisopen Then
                Set frmsuppliers.DataGrid1.DataSource = Data1
                If (.Filter <> adFilterNone) Then
                    frmsuppliers.cmdremfilter_Click
                End If
                ' goto the newly added record
                .Find .Fields(0).Name & "='" & txts_code.Text & "'"
            End If
            If MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                cmdreset_Click
                txts_code = makecode()
            Else
                If Not suppliersformisopen Then
                    .Close
                Else
                    ' enable the controls after addition if they were disabled
                    If frmsuppliers.cmdDelete.Enabled = False Then frmsuppliers.enablecontrols True
                End If
                Unload Me
            End If
        Else
            Unload Me
        End If
    End With
    Data1.conn.CommitTrans
    Exit Sub
err:
    Data1.conn.RollbackTrans
    If suppliersformisopen Then
        Set frmsuppliers.DataGrid1.DataSource = Nothing
        Data1.rssuppliers.CancelUpdate
        Set frmsuppliers.DataGrid1.DataSource = Data1
    Else
        Data1.rssuppliers.CancelUpdate
    End If
    MsgBox err.Description, vbCritical, "Error"
End Sub
Private Function makecode() As String
    Dim code As String
    Dim n As Byte
    Dim rs As New ADODB.Recordset
    rs.Open "select max(mid(s_code,2))as maxcode from suppliers", Data1.conn
    If (IsNull(rs("maxcode"))) Then
        code = "S0000001"
    Else
        n = 7 - Len(rs("maxcode") + 1)
        For i = 1 To n
            code = code & "0"
        Next i
        code = "S" & code & rs("maxcode") + 1
    End If
    makecode = code
    rs.Close
End Function

Private Sub Form_Load()
    If (modstate = True) Then
        With Data1.rssuppliers
            txts_code.Text = .Fields("s_code")
            txts_name.Text = .Fields("s_name")
            txts_addr.Text = .Fields("s_addr")
            txts_cont_no.Text = .Fields("s_cont_no")
            If .Fields("fax_no") <> "" Then txtfax_no.Text = .Fields("fax_no")
            If .Fields("email_addr") <> "" Then txtemail_addr.Text = .Fields("email_addr")
        End With
    Else
        txts_code.Text = makecode()
    End If
End Sub

Private Sub cmdreset_Click()
    txts_name.Text = ""
    txts_addr.Text = ""
    txts_cont_no.Text = ""
    txtfax_no.Text = ""
    txtemail_addr.Text = ""
    txts_name.SetFocus
End Sub

Private Function validate() As Boolean
    If txts_name.Text = "" Then
        MsgBox "Supplier Name is blank", vbInformation
        txts_name.SetFocus
        validate = False
        Exit Function
    End If
    If txts_addr.Text = "" Then
        MsgBox "Supplier Address is blank", vbInformation
        txts_addr.SetFocus
        validate = False
        Exit Function
    End If
    If txts_cont_no.Text = "" Then
        MsgBox "Contact Number is blank", vbInformation
        txts_cont_no.SetFocus
        validate = False
        Exit Function
    End If
    validate = True
End Function

Private Sub txts_addr_LostFocus()
    txts_addr.Text = Trim(txts_addr.Text)
End Sub
Private Sub txts_cont_no_LostFocus()
    txts_cont_no.Text = Trim(txts_cont_no.Text)
End Sub

Private Sub txts_name_LostFocus()
    txts_name.Text = Trim(txts_name.Text)
End Sub
