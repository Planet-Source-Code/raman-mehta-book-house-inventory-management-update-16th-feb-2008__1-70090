VERSION 5.00
Begin VB.Form frmcustomer_addmod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Customer"
   ClientHeight    =   4065
   ClientLeft      =   2775
   ClientTop       =   1635
   ClientWidth     =   5730
   Icon            =   "frmcustomer_add.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5730
   StartUpPosition =   1  'CenterOwner
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
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
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
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtc_cont_no 
      Height          =   300
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox txtc_addr 
      Height          =   300
      Left            =   1680
      MaxLength       =   80
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txtc_name 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   300
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox txtc_code 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
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
      Left            =   360
      TabIndex        =   14
      Top             =   2573
      Width           =   1200
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fax Number(s):"
      Height          =   195
      Index           =   4
      Left            =   495
      TabIndex        =   13
      Top             =   2093
      Width           =   1065
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "*Contact Number(s):"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1613
      Width           =   1425
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "*Address:"
      Height          =   195
      Index           =   2
      Left            =   870
      TabIndex        =   11
      Top             =   1133
      Width           =   675
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "*Name:"
      Height          =   195
      Index           =   1
      Left            =   1020
      TabIndex        =   10
      Top             =   653
      Width           =   525
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Customer Code:"
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   9
      Top             =   173
      Width           =   1125
   End
End
Attribute VB_Name = "frmcustomer_addmod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form handles the actual addition or modifications of the customer records
Public modstate As Boolean
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdsave_Click()
    On Error GoTo err
    If validate() = False Then Exit Sub
    Data1.conn.BeginTrans
    With Data1.rscustomers
        If (.State = adStateClosed) Then: .Open
        If (modstate = False) Then
            .AddNew
            .Fields("credit") = 0
        End If
        .Fields("c_code") = txtc_code.Text
        .Fields("c_name") = txtc_name.Text
        .Fields("c_addr") = txtc_addr.Text
        .Fields("c_cont_no") = txtc_cont_no.Text
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
        ' if sales form is open then reflect the changes in the form for customers
        If salesformisopen Then
            With Data1.rssales
                If .RecordCount > 0 Then bm = .Bookmark
                Set frmsales.DataGrid1.DataSource = Nothing
                .Close
                .Open
                Set frmsales.DataGrid1.DataSource = Data1
                If .RecordCount > 0 Then .Bookmark = bm
            End With
            With Data1.rst_sales
                If .RecordCount > 0 Then bm = .Bookmark
                Set frmsales.DataGrid2.DataSource = Nothing
                .Close
                .Open
                Set frmsales.DataGrid2.DataSource = Data1
                If .RecordCount > 0 Then .Bookmark = bm
            End With
        End If
        If modstate = False Then
            ' if the add customer form was called from customers form then
            ' deal with the records displayed on the form
            If customersformisopen Then
                Set frmcustomers.DataGrid1.DataSource = Data1
                ' remove the filter before addition
                If (.Filter <> adFilterNone) Then
                    frmcustomers.cmdremfilter_Click
                End If
                ' goto the newly added record
                .Find .Fields(0).Name & "='" & txtc_code.Text & "'"
            End If
            If MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                cmdreset_Click
                txtc_code.Text = makecode()
            Else
                If Not customersformisopen Then
                    .Close
                Else
                    ' enable the controls after addition if they were disabled
                    If frmcustomers.cmdDelete.Enabled = False Then frmcustomers.enablecontrols True
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
    If customersformisopen Then
        Set frmcustomers.DataGrid1.DataSource = Nothing
        Data1.rscustomers.CancelUpdate
        Set frmcustomers.DataGrid1.DataSource = Data1
    Else
        Data1.rscustomers.CancelUpdate
    End If
    MsgBox err.Description, vbCritical, "Error"
End Sub
Private Function makecode() As String
    Dim code As String
    Dim n As Byte
    Dim rs As New ADODB.Recordset
    rs.Open "select max(mid(c_code,2))as maxcode from customers", Data1.conn
    If (IsNull(rs("maxcode"))) Then
        code = "C0000001"
    Else
        n = 7 - Len(rs("maxcode") + 1)
        For i = 1 To n
            code = code & "0"
        Next i
        code = "C" & code & rs("maxcode") + 1
    End If
    makecode = code
    rs.Close
End Function

Private Sub cmdreset_Click()
    txtc_name.Text = ""
    txtc_addr.Text = ""
    txtc_cont_no.Text = ""
    txtfax_no.Text = ""
    txtemail_addr.Text = ""
    txtc_name.SetFocus
End Sub

Private Sub Form_Load()
    If (modstate = True) Then
        With Data1.rscustomers
            txtc_code.Text = .Fields("c_code")
            txtc_name.Text = .Fields("c_name")
            txtc_addr.Text = .Fields("c_addr")
            txtc_cont_no.Text = .Fields("c_cont_no")
            If .Fields("fax_no") <> "" Then txtfax_no.Text = .Fields("fax_no")
            If .Fields("email_addr") <> "" Then txtemail_addr.Text = .Fields("email_addr")
        End With
    Else
        txtc_code.Text = makecode()
    End If
End Sub
' validation begins here
Private Function validate() As Boolean
    If txtc_name.Text = "" Then
        MsgBox "Customer Name is blank", vbInformation
        txtc_name.SetFocus
        validate = False
        Exit Function
    End If
    If txtc_addr.Text = "" Then
        MsgBox "Customer Address is blank", vbInformation
        txtc_addr.SetFocus
        validate = False
        Exit Function
    End If
    If txtc_cont_no.Text = "" Then
        MsgBox "Contact Number is blank", vbInformation
        txtc_cont_no.SetFocus
        validate = False
        Exit Function
    End If
    validate = True
End Function
Private Sub txtc_addr_LostFocus()
    txtc_addr.Text = Trim(txtc_addr.Text)
End Sub

Private Sub txtc_cont_no_LostFocus()
    txtc_cont_no.Text = Trim(txtc_cont_no.Text)
End Sub

Private Sub txtc_name_LostFocus()
    txtc_name.Text = Trim(txtc_name.Text)
End Sub
