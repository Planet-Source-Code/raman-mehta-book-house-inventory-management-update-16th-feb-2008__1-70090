VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmtitle_addmod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Title"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   Icon            =   "frmtitle_addmod.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6195
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtdigits 
      Height          =   300
      Left            =   4920
      MaxLength       =   2
      TabIndex        =   10
      Top             =   3480
      Width           =   735
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   290
      Left            =   5401
      TabIndex        =   8
      Top             =   3000
      Width           =   255
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txted_no"
      BuddyDispid     =   196616
      OrigLeft        =   5640
      OrigTop         =   3000
      OrigRight       =   5895
      OrigBottom      =   3300
      Max             =   99
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdsub 
      Caption         =   "..."
      Height          =   300
      Left            =   5760
      TabIndex        =   3
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtauthors 
      Height          =   300
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2040
      Width           =   3975
   End
   Begin MSDataListLib.DataCombo cbosub 
      Bindings        =   "frmtitle_addmod.frx":08CA
      DataField       =   "S_CODE"
      DataMember      =   "titles"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "S_NAME"
      BoundColumn     =   "S_CODE"
      Text            =   "DataCombo1"
      Object.DataMember      =   "subjects"
   End
   Begin VB.TextBox txtprice 
      Height          =   300
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   9
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox txted_no 
      Height          =   300
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   7
      Top             =   3000
      Width           =   3720
   End
   Begin VB.TextBox txtisbn 
      Height          =   300
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   6
      Top             =   2520
      Width           =   3975
   End
   Begin VB.TextBox txtt_desc 
      Height          =   300
      Left            =   1680
      MaxLength       =   60
      TabIndex        =   4
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox txtt_name 
      Height          =   300
      Left            =   1680
      MaxLength       =   60
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox txtt_code 
      BackColor       =   &H00E0E0E0&
      DataSource      =   "Data1"
      Height          =   300
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label2 
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
      Left            =   4740
      TabIndex        =   23
      Top             =   3165
      Width           =   135
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
      TabIndex        =   22
      Top             =   3840
      Width           =   2385
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "*Author(s):"
      Height          =   195
      Index           =   7
      Left            =   750
      TabIndex        =   21
      Top             =   2100
      Width           =   735
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "*Price:"
      Height          =   195
      Index           =   6
      Left            =   1020
      TabIndex        =   20
      Top             =   3540
      Width           =   465
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "*Edition Number:"
      Height          =   195
      Index           =   5
      Left            =   300
      TabIndex        =   19
      Top             =   3060
      Width           =   1185
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "*ISBN:"
      Height          =   195
      Index           =   4
      Left            =   1005
      TabIndex        =   18
      Top             =   2580
      Width           =   480
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   195
      Index           =   3
      Left            =   645
      TabIndex        =   17
      Top             =   1605
      Width           =   840
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "*Subject:"
      Height          =   195
      Index           =   2
      Left            =   840
      TabIndex        =   16
      Top             =   1110
      Width           =   645
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "*Title Name:"
      Height          =   195
      Index           =   1
      Left            =   615
      TabIndex        =   15
      Top             =   600
      Width           =   870
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Title Code:"
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   14
      Top             =   165
      Width           =   765
   End
End
Attribute VB_Name = "frmtitle_addmod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form is responsible for the actual additon or deletion
' of the title records
Public modstate As Boolean

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdsave_Click()
    On Error GoTo err
    Dim bm
    If validate() = False Then Exit Sub
    Dim subcode As String
    Data1.conn.BeginTrans
    subcode = cbosub.BoundText
    With Data1.rstitles
        If (.State = adStateClosed) Then: .Open
        If (modstate = False) Then: .AddNew
        .Fields("t_code") = txtt_code.Text
        .Fields("t_name") = txtt_name.Text
        .Fields("s_code") = subcode
        If txtt_desc.Text = "" Then
            .Fields("t_desc") = Null
        Else
            .Fields("t_desc") = txtt_desc.Text
        End If
        .Fields("authors") = txtauthors.Text
        .Fields("isbn") = txtisbn.Text
        .Fields("ed_no") = txted_no.Text
        ' concatenate the fractional part to the price amount
        .Fields("price") = txtprice.Text & "." & txtdigits.Text
        .Update
        ' if purchasesform is open then reflect the changes in the form
        ' regarding the title name, price etc.
        If purchasesformisopen Then
            With Data1.rspurchases
                If .RecordCount > 0 Then bm = .Bookmark
                Set frmpurchases.DataGrid1.DataSource = Nothing
                .Close
                .Open
                Set frmpurchases.DataGrid1.DataSource = Data1
                If .RecordCount > 0 Then .Bookmark = bm
            End With
        End If
        If salesformisopen Then
            With Data1.rssales
                If .RecordCount > 0 Then bm = .Bookmark
                Set frmsales.DataGrid1.DataSource = Nothing
                .Close
                .Open
                Set frmsales.DataGrid1.DataSource = Data1
                If .RecordCount > 0 Then .Bookmark = bm
            End With
        End If
        If modstate = False Then
            If titlesformisopen Then
                .Close
                .Open
                Set frmtitles.DataGrid1.DataSource = Data1
                If (.Filter <> adFilterNone) Then
                    frmtitles.cmdremfilter_Click
                End If
                .Find .Fields(0).Name & "='" & txtt_code.Text & "'"
            End If
            If MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                cmdreset_Click
                txtt_code.Text = makecode()
            Else
                If Not titlesformisopen Then
                    .Close
                Else
                    If frmtitles.cmdDelete.Enabled = False Then frmtitles.enablecontrols True
                End If
                Unload Me
            End If
        Else
            ' close and open of the recordset is necessary for refreshing the
            ' recordset as it is based on a query which will be invoked against the
            ' changed records only if it is reopened
            Set frmtitles.DataGrid1.DataSource = Nothing
            .Close
            .Open
            .Find .Fields(0).Name & "='" & txtt_code.Text & "'"
            Set frmtitles.DataGrid1.DataSource = Data1
            Unload Me
        End If
    End With
    Data1.conn.CommitTrans
    Exit Sub
err:
    Data1.conn.RollbackTrans
    If titlesformisopen Then
        Set frmtitles.DataGrid1.DataSource = Nothing
        Data1.rstitles.CancelUpdate
        Set frmtitles.DataGrid1.DataSource = Data1
    Else
        Data1.rstitles.CancelUpdate
    End If
    MsgBox err.Description, vbCritical, "Error"
End Sub
Private Function makecode() As String
    Dim code As String
    Dim n As Byte
    Dim rs As New ADODB.Recordset
    rs.Open "select max(mid(t_code,2))as maxcode from titles", Data1.conn
    ' if this is the first record to be entered then nothing is found in the titles
    ' table so generate the first code manually
    If (IsNull(rs("maxcode"))) Then
        code = "T0000001"
    Else
        n = 7 - Len(rs("maxcode") + 1)
        For i = 1 To n
            code = code & "0"
        Next i
        code = "T" & code & rs("maxcode") + 1
    End If
    makecode = code
    rs.Close
End Function


Private Sub cmdsub_Click()
    frmsubjects.Show vbModal
End Sub


Private Sub Form_Load()
    If (modstate = True) Then
        Me.Caption = "Modify Title"
        With Data1.rstitles
            txtt_code.Text = .Fields("t_code")
            txtt_name.Text = .Fields("t_name")
            If .Fields("t_desc") <> "" Then txtt_desc.Text = .Fields("t_desc")
            txtauthors.Text = .Fields("authors")
            txtisbn.Text = .Fields("isbn")
            txted_no.Text = .Fields("ed_no")
            Dim pos As Byte
            pos = InStr(.Fields("price"), ".")
            ' if the price contains the fractional part then show it in
            ' two digits by appnding 0's.
            ' the fractional part is determined by searching the position
            ' of a . in the amount
            If pos = 0 Then
                txtprice.Text = .Fields("price")
                txtdigits.Text = "00"
            Else
                txtprice.Text = Left(.Fields("price"), pos - 1)
                txtdigits.Text = Mid(.Fields("price"), pos + 1)
                For i = 1 To 2 - Len(txtdigits.Text)
                    txtdigits = txtdigits & "0"
                Next i
            End If
        End With
    Else
        txtt_code.Text = makecode()
        cbosub.Text = ""
    End If
    titleaddmodformisopen = True
End Sub

Private Sub cmdreset_Click()
    txtt_name.Text = ""
    cbosub.Text = ""
    txtt_desc.Text = ""
    txtauthors.Text = ""
    txtisbn.Text = ""
    txted_no.Text = ""
    txtprice.Text = ""
    txtdigits.Text = ""
    txtt_name.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    titleaddmodformisopen = False
End Sub



Private Sub txtauthors_LostFocus()
    txtauthors.Text = Trim(txtauthors.Text)
End Sub


Private Sub txtdigits_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub

Private Sub txted_no_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub
Private Function validate() As Boolean
    If txtt_name.Text = "" Then
        MsgBox "Title Name is blank", vbInformation
        txtt_name.SetFocus
        validate = False
        Exit Function
    End If
    If cbosub.Text = "" Then
        MsgBox "Subject Name is blank", vbInformation
        cbosub.SetFocus
        validate = False
        Exit Function
    End If
    If txtauthors.Text = "" Then
        MsgBox "Authors' Name is blank", vbInformation
        txtauthors.SetFocus
        validate = False
        Exit Function
    End If
    If txtisbn.Text = "" Then
        MsgBox "ISBN is blank", vbInformation
        txtisbn.SetFocus
        validate = False
        Exit Function
    End If
    If txted_no.Text = "" Then
        MsgBox "Edition Number is blank", vbInformation
        txted_no.SetFocus
        validate = False
        Exit Function
    End If
    If txtprice.Text = "" Then
        MsgBox "Price is blank", vbInformation
        txtprice.SetFocus
        validate = False
        Exit Function
    End If
    validate = True
End Function
Private Sub txtisbn_LostFocus()
    txtisbn.Text = Trim(txtisbn.Text)
End Sub

Private Sub txtprice_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub

Private Sub txtt_name_LostFocus()
    ' delete the leading and trailing spaces in the title name
    txtt_name.Text = Trim(txtt_name.Text)
End Sub
