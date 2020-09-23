VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmcustomers 
   Caption         =   "Customers"
   ClientHeight    =   7650
   ClientLeft      =   420
   ClientTop       =   2370
   ClientWidth     =   11880
   Icon            =   "frmcustomers.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame framescroll 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10320
      TabIndex        =   32
      Top             =   6720
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   400
      Left            =   4320
      SmallChange     =   200
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1335
      LargeChange     =   400
      Left            =   6600
      Max             =   1000
      SmallChange     =   200
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmcustomers.frx":08CA
      Height          =   6015
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10610
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   15920516
      ColumnHeaders   =   -1  'True
      ForeColor       =   0
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
      DataMember      =   "customers"
      Caption         =   "Current Customers"
      ColumnCount     =   7
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "C_ADDR"
         Caption         =   "Address"
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
         DataField       =   "C_CONT_NO"
         Caption         =   "Contact Number(s)"
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
         DataField       =   "FAX_NO"
         Caption         =   "Fax Number(s)"
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
      BeginProperty Column05 
         DataField       =   "EMAIL_ADDR"
         Caption         =   "Email Address(s)"
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
      BeginProperty Column06 
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1785.26
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1470.047
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Operations"
      Height          =   855
      Left            =   1080
      TabIndex        =   24
      Top             =   6120
      Width           =   6615
      Begin VB.CommandButton cmddelall 
         Caption         =   "Delete A&ll"
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   5520
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   4440
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdnewmod 
         Caption         =   "&Modiy"
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdnewmod 
         Caption         =   "&Add"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filter Records"
      Height          =   2775
      Left            =   8760
      TabIndex        =   20
      Top             =   2160
      Width           =   3015
      Begin VB.CommandButton cmdremfilter 
         Caption         =   "&Remove Filter"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdfilter 
         Caption         =   "&Filter"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   1215
      End
      Begin VB.ComboBox cbofield 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmcustomers.frx":08DE
         Left            =   240
         List            =   "frmcustomers.frx":08F7
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1770
         Width           =   2535
      End
      Begin VB.TextBox txtvalue 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   1170
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter a text which you want to filter and select a field where to locate it."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   700
         TabIndex        =   23
         Top             =   240
         Width           =   2235
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   120
         Picture         =   "frmcustomers.frx":095C
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label8 
         Caption         =   "Look in:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1530
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Look for:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   930
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sort Records"
      Height          =   1695
      Left            =   8760
      TabIndex        =   18
      Top             =   120
      Width           =   3015
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Descending"
         Height          =   495
         Index           =   1
         Left            =   1680
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ascending"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.ComboBox cbosort 
         Height          =   315
         ItemData        =   "frmcustomers.frx":1626
         Left            =   240
         List            =   "frmcustomers.frx":163F
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Sort By:"
         Height          =   255
         Left            =   645
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmcustomers.frx":16A4
         Top             =   200
         Width           =   480
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   8760
      TabIndex        =   27
      Top             =   5640
      Width           =   3015
      Begin VB.CommandButton cmdnavigate 
         Height          =   265
         Index           =   3
         Left            =   1860
         Picture         =   "frmcustomers.frx":1F6E
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   400
      End
      Begin VB.CommandButton cmdnavigate 
         Height          =   265
         Index           =   2
         Left            =   1440
         Picture         =   "frmcustomers.frx":22F8
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         Width           =   400
      End
      Begin VB.CommandButton cmdnavigate 
         Height          =   265
         Index           =   1
         Left            =   1020
         Picture         =   "frmcustomers.frx":2682
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Width           =   400
      End
      Begin VB.CommandButton cmdnavigate 
         Height          =   265
         Index           =   0
         Left            =   600
         Picture         =   "frmcustomers.frx":2A0C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Width           =   400
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Record"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lblcurrec 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   720
         TabIndex        =   30
         Top             =   360
         Width           =   90
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "of"
         Height          =   195
         Left            =   1560
         TabIndex        =   29
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lblttlrec 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   2160
         TabIndex        =   28
         Top             =   360
         Width           =   90
      End
   End
End
Attribute VB_Name = "frmcustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form displays the customer records and enables addition,
' deletion etc. of records
Dim order As String * 4
Dim vsval As Integer
Dim hsval As Integer
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
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
    ' call the database directly to delete all customers
    Data1.conn.Execute ("delete from customers")
    Data1.rscustomers.Close
    Data1.rscustomers.Open
    Set DataGrid1.DataSource = Data1
    Data1.conn.CommitTrans
    enablecontrols False
    cmdremfilter.Enabled = False
    Exit Sub
err:
    Data1.conn.RollbackTrans
    Set DataGrid1.DataSource = Nothing
    Data1.conn.Cancel
    Set DataGrid1.DataSource = Data1
    MsgBox err.Description, vbCritical, "Error"
End Sub

Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If customersformisopen Then: updatelabels
End Sub

Private Sub cbosort_Click()
    With Data1.rscustomers
        .Sort = .Fields(cbosort.ItemData(cbosort.ListIndex)).Name & " " & order
    End With
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo err
    If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Confirm deletion") = vbNo Then
        Exit Sub
    End If
    Data1.conn.BeginTrans
    With Data1.rscustomers
        .Delete
        If .RecordCount = 0 Then
            .MoveFirst
            enablecontrols False
        Else
            .MoveNext
            If .EOF Then
                .MoveLast
            End If
        End If
    End With
    Data1.conn.CommitTrans
    Exit Sub
err:
    Data1.conn.RollbackTrans
    Set DataGrid1.DataSource = Nothing
    Data1.rscustomers.CancelUpdate
    Set DataGrid1.DataSource = Data1
    MsgBox err.Description, vbCritical, "Error"
End Sub

Private Sub cmdfilter_Click()
    With Data1.rscustomers
        .Filter = .Fields(cbofield.ItemData(cbofield.ListIndex)).Name & "='" & txtvalue.Text & "'"
        If .RecordCount = 0 Then enablecontrols False
    End With
    cmdremfilter.Enabled = True
End Sub

Private Sub cmdnavigate_Click(Index As Integer)
    With Data1.rscustomers
        Select Case Index
        Case 0:
            .MoveFirst
        Case 1:
            .MovePrevious
            If .BOF Then
                .MoveFirst
            End If
        Case 2:
            .MoveNext
            If .EOF Then
                .MoveLast
            End If
        Case 3:
            .MoveLast
        End Select
    End With
End Sub

Public Sub cmdnewmod_Click(Index As Integer)
    With frmcustomer_addmod
        ' modstate variable indicates whether the customer add form would be called for addition or modification
        .modstate = Index
        If Index = 1 Then: .Caption = "Modify Customer Record"
        .Show vbModal
    End With
End Sub
Public Sub cmdremfilter_Click()
    cmdremfilter.Enabled = False
    Data1.rscustomers.Filter = adFilterNone
    ' if filter results in no records then disable the controls
    If Data1.rscustomers.RecordCount = 0 Then
        enablecontrols False
    ElseIf cmdDelete.Enabled = False Then
        enablecontrols True
    End If
    cmdremfilter.Enabled = False
End Sub

Private Sub cmdsearch_Click()
    With frmsearch
        Set .srsrs = Data1.rscustomers
        ' the srsarr variable of search form is used to fill the combo box dynamically
        ' with appropriate fields
        .srsarr = Array("Customer Code", "Customer Name", "Address", "Contact Number", "Fax Number", "Email Address", "Credit Amount")
        ' the srsdata variable stores indexes for fields in the current recordset
        ' in order to facilitate search
        .srsdata = Array(0, 1, 2, 3, 4, 5, 6)
        .Show vbModal
    End With
End Sub

Private Sub DataGrid1_DblClick()
    If enable = True Then cmdnewmod_Click 1
End Sub
Private Sub DataGrid1_Keydown(KeyCode As Integer, Shift As Integer)
    If enable = True Then
        If KeyCode = vbKeyReturn Then cmdnewmod_Click 1
    End If
End Sub

Private Sub Form_Load()
    cmdnewmod(0).Enabled = enable
    cmdnewmod(1).Enabled = enable
    cmdDelete.Enabled = enable
    cmddelall.Enabled = enable
    cmdfilter.Enabled = enable
    cbosort.ListIndex = 0
    cbofield.ListIndex = 0
    If Data1.rscustomers.RecordCount = 0 Then enablecontrols False
    updatelabels
    vsval = 0
    hsval = 0
    Set rs = Data1.rscustomers
    customersformisopen = True
End Sub
' this module displays scroll bars on the form when the form
' is resized below its maximum length
Private Sub Form_Resize()
    'the hdisp and vdisp variables calculate the difference between the
    ' current dimensions of the form and the dimenstions of the maximized form
    ' in order to calculate whether scroll bars should be displayed or not
    Dim hdisp As Integer
    Dim vdisp As Integer
    Dim h As Integer
    Dim w As Integer
    ' 12060 is the width of the maximized form
    hdisp = Me.Width - 12060
    ' 8310 is the height of the maximized form plus 375 for the height of the status bar
    vdisp = Me.Height - 8310 + 375 + 600
    If hdisp >= 0 And vdisp >= 0 Then
        'if form is resized above the maximized dimensions then
        ' hide both the scrollbars
        VScroll1.Visible = False
        HScroll1.Visible = False
        VScroll1.Value = 0
        HScroll1.Value = 0
        ' framescroll must be visible along with both the scroll bars
        framescroll.Visible = False
        ' postion the controls to their original locations
        positioncontrols
    ElseIf hdisp >= 0 And vdisp < 0 Then
        ' if height of the resized form falls below the maximized height then
        ' display only the vertical scroll bar
        VScroll1.Visible = True
        HScroll1.Visible = False
        HScroll1.Value = 0
        ' make horizontal space because horizontal scroll bar will not
        ' be visible
        h = 0
        ' change the maximum value of the scroll bar so that scrolling is made
        ' only to the appropriate extent
        HScroll1.Max = -hdisp + VScroll1.Width
        framescroll.Visible = False
    ElseIf hdisp < 0 And vdisp >= 0 Then
        HScroll1.Visible = True
        VScroll1.Visible = False
        VScroll1.Value = 0
        w = 0
        VScroll1.Max = -vdisp + HScroll1.Height
        framescroll.Visible = False
    Else
        HScroll1.Max = -hdisp + VScroll1.Width
        VScroll1.Max = -vdisp + HScroll1.Height
        VScroll1.Visible = True
        HScroll1.Visible = True
        h = HScroll1.Height
        w = VScroll1.Width
        framescroll.Visible = True
    End If
    ' position the scroll bars according to the size of the form
    With VScroll1
        If .Visible Then
            .Top = Me.ScaleTop
            .Left = Me.ScaleWidth - .Width
            .Height = Abs(Me.ScaleHeight - h)
        End If
    End With
    With HScroll1
        If .Visible Then
            .Top = Me.ScaleHeight - .Height
            .Left = Me.ScaleLeft
            .Width = Abs(Me.ScaleWidth - w)
        End If
    End With
    With framescroll
        .Left = VScroll1.Left
        .Top = HScroll1.Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    customersformisopen = False
    If Data1.rscustomers.Filter <> adFilterNone Then Data1.rscustomers.Filter = adFilterNone
    Data1.rscustomers.Close
End Sub


Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0: order = "ASC"
        Case 1: order = "DESC"
    End Select
    cbosort_Click
End Sub
Sub enablecontrols(val As Boolean)
    cmdnewmod(1).Enabled = val
    cmdDelete.Enabled = val
    cmddelall.Enabled = val
    cmdsearch.Enabled = val
    cbosort.Enabled = val
    Option1(0).Enabled = val
    Option1(1).Enabled = val
    cmdfilter.Enabled = val
    DataGrid1.Enabled = val
    cmdnavigate(0).Enabled = val
    cmdnavigate(1).Enabled = val
    cmdnavigate(2).Enabled = val
    cmdnavigate(3).Enabled = val
 End Sub

Private Sub VScroll1_Change()
    Dim inc As Integer
    ' inc stores the amount of scrolling
    inc = VScroll1.Value - vsval
    vsval = VScroll1.Value
    ' the controls are positioned according to the amount of scrolling
    DataGrid1.Top = DataGrid1.Top - inc
    Frame1.Top = Frame1.Top - inc
    Frame2.Top = Frame2.Top - inc
    Frame3.Top = Frame3.Top - inc
    Frame4.Top = Frame4.Top - inc
End Sub
Private Sub vScroll1_Scroll()
    VScroll1_Change
End Sub
Private Sub HScroll1_Change()
    Dim inc As Integer
    inc = HScroll1.Value - hsval
    hsval = HScroll1.Value
    DataGrid1.Left = DataGrid1.Left - inc
    Frame1.Left = Frame1.Left - inc
    Frame2.Left = Frame2.Left - inc
    Frame3.Left = Frame3.Left - inc
    Frame4.Left = Frame4.Left - inc
End Sub
Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub
Private Sub positioncontrols()
    ' position the controls to their original locations
    DataGrid1.Left = 120
    DataGrid1.Top = 120
    Frame1.Left = 8760
    Frame1.Top = 120
    Frame2.Left = 8760
    Frame2.Top = 2280
    Frame3.Left = 1080
    Frame3.Top = 6120
    Frame4.Left = 8760
    Frame4.Top = 5520
End Sub
Private Sub updatelabels()
    ' show the current record no and total no of records whenever
    ' recordset's move event is called
    With Data1.rscustomers
        If .RecordCount = 0 Then
            lblcurrec.Caption = "0"
        Else
            lblcurrec.Caption = .AbsolutePosition
        End If
        lblttlrec.Caption = .RecordCount
    End With
End Sub
