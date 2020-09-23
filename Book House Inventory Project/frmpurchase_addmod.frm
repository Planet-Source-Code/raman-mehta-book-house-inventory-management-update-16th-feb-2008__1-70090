VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpurchase_addmod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Purchase"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   Icon            =   "frmpurchase_addmod.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10335
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdreset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   6360
      TabIndex        =   28
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox txtnettotal 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox txtcredittotal 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox txtdistotal 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox txtgtotal 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear &All"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8880
      TabIndex        =   21
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdremove 
      Caption         =   "Remove &from List"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8880
      TabIndex        =   20
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add to &List"
      Height          =   495
      Left            =   8880
      TabIndex        =   19
      Top             =   4200
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid list1 
      Height          =   2175
      Left            =   120
      TabIndex        =   29
      Top             =   4080
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      BackColorBkg    =   13479683
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   27
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   26
      Top             =   7080
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Height          =   1400
      Left            =   120
      TabIndex        =   38
      Top             =   2340
      Width           =   10100
      Begin VB.TextBox txts_name 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   6525
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtaddr 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txts_credit 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   840
         Width           =   1380
      End
      Begin VB.CommandButton cmdsupplier 
         Caption         =   "..."
         Height          =   300
         Left            =   2970
         TabIndex        =   16
         Top             =   345
         Width           =   375
      End
      Begin VB.TextBox txts_code 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   345
         Width           =   1455
      End
      Begin VB.Label lbladdr 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address:"
         Height          =   195
         Left            =   640
         TabIndex        =   42
         Top             =   885
         Width           =   615
      End
      Begin VB.Label lbls_credit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Credit Balance:"
         Height          =   195
         Left            =   7275
         TabIndex        =   41
         Top             =   885
         Width           =   1080
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Supplier Code:"
         Height          =   195
         Index           =   3
         Left            =   220
         TabIndex        =   40
         Top             =   390
         Width           =   1035
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Supplier Name:"
         Height          =   195
         Index           =   4
         Left            =   5295
         TabIndex        =   39
         Top             =   405
         Width           =   1080
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   33
      Top             =   720
      Width           =   10100
      Begin VB.TextBox txtcreditdigits 
         Height          =   300
         Left            =   7440
         MaxLength       =   2
         TabIndex        =   13
         Top             =   1060
         Width           =   540
      End
      Begin VB.TextBox txtdisdigits 
         Height          =   300
         Left            =   4850
         MaxLength       =   2
         TabIndex        =   11
         Top             =   1060
         Width           =   540
      End
      Begin VB.TextBox txtnetamount 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1060
         Width           =   1380
      End
      Begin VB.TextBox txtdis 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   300
         Left            =   3525
         MaxLength       =   5
         TabIndex        =   10
         Top             =   1060
         Width           =   1140
      End
      Begin VB.TextBox txtcredit 
         Height          =   300
         Left            =   5880
         MaxLength       =   6
         TabIndex        =   12
         Top             =   1060
         Width           =   1380
      End
      Begin VB.TextBox txttotal 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1060
         Width           =   1380
      End
      Begin VB.TextBox txtqty 
         Height          =   300
         Left            =   220
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1060
         Width           =   765
      End
      Begin VB.CommandButton cmdtitle 
         Caption         =   "..."
         Height          =   300
         Left            =   1680
         TabIndex        =   3
         Top             =   420
         Width           =   375
      End
      Begin VB.TextBox txtstock 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   420
         Width           =   900
      End
      Begin VB.TextBox txtt_code 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   220
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox txtt_name 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   420
         Width           =   3135
      End
      Begin VB.TextBox txtprice 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   420
         Width           =   1305
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   990
         TabIndex        =   8
         Top             =   1065
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtqty"
         BuddyDispid     =   196635
         OrigLeft        =   2640
         OrigTop         =   915
         OrigRight       =   2895
         OrigBottom      =   1200
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label7 
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
         Left            =   7290
         TabIndex        =   53
         Top             =   720
         Width           =   135
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
         Left            =   4700
         TabIndex        =   52
         Top             =   720
         Width           =   135
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Discount:"
         Height          =   195
         Index           =   9
         Left            =   3525
         TabIndex        =   47
         Top             =   840
         Width           =   675
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Credit Amount:"
         Height          =   195
         Index           =   10
         Left            =   5880
         TabIndex        =   46
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Net Amount:"
         Height          =   195
         Index           =   11
         Left            =   8520
         TabIndex        =   45
         Top             =   840
         Width           =   885
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Index           =   8
         Left            =   1740
         TabIndex        =   44
         Top             =   840
         Width           =   405
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Quantity:"
         Height          =   195
         Index           =   7
         Left            =   220
         TabIndex        =   43
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblstock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Stock:"
         Height          =   195
         Left            =   8520
         TabIndex        =   37
         Top             =   195
         Width           =   465
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Title Code:"
         Height          =   195
         Index           =   1
         Left            =   220
         TabIndex        =   36
         Top             =   200
         Width           =   765
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Title Name:"
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   35
         Top             =   200
         Width           =   810
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   195
         Index           =   5
         Left            =   6465
         TabIndex        =   34
         Top             =   195
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Height          =   685
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   10100
      Begin MSComCtl2.DTPicker dtp1 
         Height          =   300
         Left            =   7770
         TabIndex        =   1
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3801091
         CurrentDate     =   39335
      End
      Begin VB.TextBox txto_id 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date:"
         Height          =   195
         Left            =   7200
         TabIndex        =   32
         Top             =   270
         Width           =   390
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bill No:"
         Height          =   195
         Index           =   0
         Left            =   220
         TabIndex        =   31
         Top             =   270
         Width           =   495
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Net Amount:"
      Height          =   195
      Left            =   8040
      TabIndex        =   51
      Top             =   6360
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total Credit:"
      Height          =   195
      Left            =   5520
      TabIndex        =   50
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Discount:"
      Height          =   195
      Left            =   3000
      TabIndex        =   49
      Top             =   6360
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Grand Total:"
      Height          =   195
      Left            =   480
      TabIndex        =   48
      Top             =   6360
      Width           =   885
   End
End
Attribute VB_Name = "frmpurchase_addmod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form is used to make new purchase transactions as well as modifictions
' to the existing records
Public modstate As Boolean
Dim grandtotal As Single, distotal As Single, credittotal As Single
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdadd_Click()
    Dim currow
    ' if validation fails then exit the module
    If validate() = False Then Exit Sub
    ' make sure the user does not enter the same book title in the list
    For i = 1 To list1.Rows - 1
        If txtt_code.Text = list1.TextMatrix(i, 1) Then
            MsgBox "Title already exists in the list", vbInformation
            Exit Sub
        End If
    Next i
    ' increase the rows on addition of new records
    list1.Rows = list1.Rows + 1
    currow = list1.Rows - 1
    ' when at least one book is added then disable the following controls
    If cmdremove.Enabled = False Then
        cmdsupplier.Enabled = False
        dtp1.Enabled = False
        cmdremove.Enabled = True
        cmdclear.Enabled = True
        cmdsave.Enabled = True
    End If
    list1.TextMatrix(currow, 1) = txtt_code.Text
    list1.TextMatrix(currow, 2) = FormatNumber(txtprice.Text, 2, vbTrue)
    list1.TextMatrix(currow, 3) = FormatNumber(txtqty.Text, 2, vbTrue)
    list1.TextMatrix(currow, 4) = FormatNumber(txttotal.Text, 2, vbTrue)
    If txtdis.Text & txtdisdigits.Text = "" Then
        list1.TextMatrix(currow, 5) = "0.00"
    Else
        ' concatenate the fractional part
        list1.TextMatrix(currow, 5) = FormatNumber(CSng(txtdis.Text & "." & txtdisdigits.Text), 2, vbTrue)
    End If
    If txtcredit.Text & txtcreditdigits.Text = "" Then
        list1.TextMatrix(currow, 6) = "0.00"
    Else
        list1.TextMatrix(currow, 6) = FormatNumber(CSng(txtcredit.Text & "." & txtcreditdigits.Text), 2, vbTrue)
    End If
    list1.TextMatrix(currow, 7) = FormatNumber(txtnetamount.Text, 2, vbTrue)
    ' calculate grandtotal, discount total and credit total
    grandtotal = grandtotal + CSng(list1.TextMatrix(currow, 4))
    distotal = distotal + CSng(list1.TextMatrix(currow, 5))
    credittotal = credittotal + CSng(list1.TextMatrix(currow, 6))
    txtgtotal.Text = FormatNumber(grandtotal, 2, vbTrue)
    txtdistotal.Text = FormatNumber(distotal, 2, vbTrue)
    txtcredittotal.Text = FormatNumber(credittotal, 2, vbTrue)
    txtnettotal.Text = FormatNumber(CSng(txtgtotal.Text) - CSng(txtdistotal.Text) - CSng(txtcredittotal.Text), 2, vbTrue)
End Sub
Private Sub cmdclear_Click()
    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
        ' reduce the list to only header row
        list1.Rows = 1
        grandtotal = 0: distotal = 0: credittotal = 0
        txtgtotal.Text = ""
        txtdistotal.Text = ""
        txtcredittotal.Text = ""
        txtnettotal.Text = ""
        cmdsupplier.Enabled = True
        dtp1.Enabled = True
        cmdremove.Enabled = False
        cmdclear.Enabled = False
        cmdsave.Enabled = False
    End If
End Sub

Private Sub cmdremove_Click()
    Dim start As Byte, finish As Byte
    With list1
    ' muliple rows are selected for deletion then calculate the starting and ending positions
    ' for deletion
        If .Row < .RowSel Then
            start = .Row: finish = .RowSel
        Else
            start = .RowSel: finish = .Row
        End If
        For i = start To finish
            If .Rows = 2 Then
                ' if prior to deletion there was only one row other than the header row then
                ' reduce to only header row. simple deletion induces error
                .Rows = 1
                grandtotal = 0: distotal = 0: credittotal = 0
                txtgtotal.Text = ""
                txtdistotal.Text = ""
                txtcredittotal.Text = ""
                txtnettotal.Text = ""
                cmdsupplier.Enabled = True
                dtp1.Enabled = True
                cmdremove.Enabled = False
                cmdclear.Enabled = False
                cmdsave.Enabled = False
            Else
                grandtotal = grandtotal - CSng(list1.TextMatrix(.Row, 4))
                distotal = distotal - CSng(list1.TextMatrix(.Row, 5))
                credittotal = credittotal - CSng(list1.TextMatrix(.Row, 6))
                txtgtotal.Text = FormatNumber(CSng(txtgtotal.Text) - CSng(list1.TextMatrix(.Row, 4)), 2, vbTrue)
                txtdistotal.Text = FormatNumber(CSng(txtdistotal.Text) - CSng(list1.TextMatrix(.Row, 5)), 2, vbTrue)
                txtcredittotal.Text = FormatNumber(CSng(txtcredittotal.Text) - CSng(list1.TextMatrix(.Row, 6)), 2, vbTrue)
                txtnettotal.Text = FormatNumber(CSng(txtnettotal.Text) - CSng(list1.TextMatrix(.Row, 7)), 2, vbTrue)
                .RemoveItem (.Row)
            End If
        Next i
    End With
    End Sub

Private Sub cmdsave_Click()
    On Error GoTo err:
    Dim rs As New ADODB.Recordset
    Data1.conn.BeginTrans
    With Data1.rst_purchases
        If (.State = adStateClosed) Then: .Open
        If modstate = False Then
            .AddNew
        Else
            ' if the form was called for modification then delete the previous values from
            ' the purchasedetail records and supplier records
            rs.Open "select * from purchasedetails where o_id='" & txto_id.Text & "'", Data1.conn, adOpenForwardOnly, adLockOptimistic
            While Not rs.EOF
                With Data1.rstitles
                    ' remove the filter if any
                    .Filter = adFilterNone
                    .MoveFirst
                    .Find "t_code='" & rs.Fields("t_code") & "'"
                    .Fields("stock") = .Fields("stock") - rs.Fields("qty")
                End With
                With Data1.rssuppliers
                    ' remove the filter if any
                    .Filter = adFilterNone
                    .MoveFirst
                    .Find "s_code='" & txts_code.Text & "'"
                    .Fields("credit") = .Fields("credit") - rs.Fields("credit")
                End With
                rs.MoveNext
            Wend
            rs.Close
            Set frmpurchases.DataGrid1.DataSource = Nothing
            Data1.conn.Execute "delete from purchasedetails where o_id='" & Data1.rst_purchases.Fields(0) & "'"
            Data1.rspurchases.Close
            Data1.rspurchases.Open
            Set frmpurchases.DataGrid1.DataSource = Data1
        End If
        'update the purchasedetails table and suppliers table with new information
        ' but first remove the filters if any
        Data1.rstitles.Filter = adFilterNone
        Data1.rssuppliers.Filter = adFilterNone
        For i = 1 To list1.Rows - 1
            With Data1.rstitles
                .MoveFirst
                .Find "t_code='" & list1.TextMatrix(i, 1) & "'"
                .Fields("stock") = .Fields("stock") + val(list1.TextMatrix(i, 3))
                .Update
            End With
        Next i
        With Data1.rssuppliers
            .MoveFirst
            .Find "s_code='" & txts_code.Text & "'"
            .Fields("credit") = .Fields("credit") + credittotal
            .Update
        End With
        .Fields("o_id") = txto_id.Text
        .Fields("s_code") = txts_code.Text
        .Fields("p_date") = dtp1.Value
        .Fields("g_total") = grandtotal
        .Fields("d_total") = distotal
        .Fields("c_total") = credittotal
        .Update
    End With
    With Data1.rspurchases
        If (.State = adStateClosed) Then: .Open
        ' remove the filter if any
        .Filter = adFilterNone
        For i = 1 To list1.Rows - 1
              .AddNew
              .Fields("o_id") = txto_id.Text
              .Fields("t_code") = list1.TextMatrix(i, 1)
              .Fields("price") = CSng(list1.TextMatrix(i, 2))
              .Fields("qty") = list1.TextMatrix(i, 3)
              .Fields("total") = CSng(list1.TextMatrix(i, 4))
              .Fields("dis") = CSng(list1.TextMatrix(i, 5))
              .Fields("credit") = CSng(list1.TextMatrix(i, 6))
              .Update
        Next i
     End With
     If modstate = False Then
        ' if purchasesformisopen then reflect the changes in that form
        If purchasesformisopen Then
            With Data1.rst_purchases
                Set frmpurchases.DataGrid2.DataSource = Nothing
                .Close
                .Open
                Set frmpurchases.DataGrid2.DataSource = Data1
                If (.Filter <> adFilterNone) Then
                    frmpurchases.cmdremfilter2_Click
                End If
                .Find "o_id='" & txto_id.Text & "'"
            End With
            With Data1.rspurchases
                Set frmpurchases.DataGrid1.DataSource = Nothing
                .Close
                .Open
                Set frmpurchases.DataGrid1.DataSource = Data1
                If (.Filter <> adFilterNone) Then
                    frmpurchases.cmdremfilter_Click
                End If
                .Find "o_id='" & txto_id.Text & "'"
            End With
        End If
        If MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            cmdreset_Click
            txto_id.Text = makecode()
        Else
            ' close the recordsets as there is no further need for them
            If Not purchasesformisopen Then
                Data1.rst_purchases.Close
                Data1.rspurchases.Close
            Else
                If frmpurchases.cmdDelete.Enabled = False Then
                    frmpurchases.enablecontrols True
                    frmpurchases.enablecontrols2 True
                End If
            End If
            Unload Me
        End If
    Else
        With Data1.rst_purchases
            Set frmpurchases.DataGrid2.DataSource = Nothing
            .Close
            .Open
            Set frmpurchases.DataGrid2.DataSource = Data1
            frmpurchases.cmdremfilter2_Click
            .Find "o_id='" & txto_id.Text & "'"
        End With
        With Data1.rspurchases
            Set frmpurchases.DataGrid1.DataSource = Nothing
            .Close
            .Open
            Set frmpurchases.DataGrid1.DataSource = Data1
            frmpurchases.cmdremfilter_Click
            .Find "o_id='" & txto_id.Text & "'"
        End With
        Unload Me
    End If
    ' deal with the filter and remove filter buttons on the suppliers and titles form
    If suppliersformisopen Then frmsuppliers.cmdremfilter_Click
    If titlesformisopen Then frmtitles.cmdremfilter_Click
    Data1.conn.CommitTrans
Exit Sub
err:
    Data1.conn.RollbackTrans
    ' cancel the updates made in case of errors
    If purchasesformisopen Then
        Set frmpurchases.DataGrid1.DataSource = Nothing
        Set frmpurchases.DataGrid2.DataSource = Nothing
        Data1.rspurchases.CancelUpdate
        Data1.rst_purchases.CancelUpdate
        Set frmpurchases.DataGrid1.DataSource = Data1
        Set frmpurchases.DataGrid2.DataSource = Data1
    Else
        Data1.rspurchases.CancelUpdate
        Data1.rst_purchases.CancelUpdate
    End If
    If suppliersformisopen Then
        Set frmsuppliers.DataGrid1.DataSource = Nothing
        Data1.rssuppliers.CancelUpdate
        Set frmsuppliers.DataGrid1.DataSource = Data1
    Else
        Data1.rssuppliers.CancelUpdate
    End If
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
    rs.Open "select max(mid(o_id,2))as maxcode from purchases", Data1.conn
    If (IsNull(rs("maxcode"))) Then
        code = "O0000001"
    Else
        n = 7 - Len(rs("maxcode") + 1)
        For i = 1 To n
            code = code & "0"
        Next i
        code = "O" & code & rs("maxcode") + 1
    End If
    makecode = code
    rs.Close
End Function
Private Sub cmdsupplier_Click()
    ' call the select form in order to select supplier records for
    ' the purpose of purchase
    With frmselect
        Set .srsrs = Data1.conn.Execute("SELECT s_code,s_name,s_addr,s_cont_no,fax_no,email_addr,credit FROM suppliers")
        .srsarr = Array("Supplier Code", "Supplier Name", "Address", "Contact Number", "Fax Number", "Email Address", "Credit Amount")
        .srsdata = Array(0, 1, 2, 3, 4, 5, 6)
        .Show vbModal
        If (.recordselected = True) Then
            txts_code.Text = .srsrs.Fields(0)
            txts_name.Text = .srsrs.Fields(1)
            txtaddr.Text = .srsrs.Fields(2)
            txts_credit.Text = FormatNumber(.srsrs.Fields(6), 2, vbTrue)
        End If
        .srsrs.Close
    End With
End Sub

Private Sub cmdtitle_Click()
    With frmselect
        Set .srsrs = Data1.conn.Execute("SELECT t_code,t_name,s_name,t_desc,authors,isbn,ed_no,price,stock,demand FROM titles,subjects WHERE titles.s_code=subjects.s_code")
        .srsarr = Array("Title Code", "Title Name", "Subject", "Description", "Authors", "ISBN", "Edition Number", "Price", "Stock", "Demand")
        .srsdata = Array(0, 1, 3, 4, 5, 6, 7, 8, 9, 10)
        .Show vbModal
        If (.recordselected = True) Then
            txtt_code.Text = .srsrs.Fields(0)
            txtt_name.Text = .srsrs.Fields(1)
            txtprice.Text = FormatNumber(.srsrs.Fields(7), 2, vbTrue)
            txtstock.Text = .srsrs.Fields(8)
            txtqty.Text = 1
            txtdis.Text = "": txtdisdigits.Text = ""
            txtcredit.Text = "": txtcreditdigits.Text = ""
        End If
        .srsrs.Close
    End With
End Sub
Private Sub Form_Load()
    If (Data1.rstitles.State = adStateClosed) Then
            Data1.rstitles.Open
    End If
    If Data1.rssuppliers.State = adStateClosed Then
        Data1.rssuppliers.Open
    End If
    If (modstate = True) Then
        ' if form was called for modification then
        ' store various variables for the purpose of modification
        ' and fill the grid and other controls with current values
        With Data1.rst_purchases
            txto_id.Text = .Fields("o_id")
            dtp1.Value = .Fields("p_date")
            txts_code.Text = .Fields("s_code")
            txts_name.Text = .Fields("s_name")
            txtaddr.Text = .Fields("s_addr")
            txts_credit.Text = FormatNumber(.Fields("credit"), 2, vbTrue)
            cmdsupplier.Enabled = False
            dtp1.Enabled = False
            txtgtotal.Text = FormatNumber(.Fields("g_total"), 2, vbTrue)
            txtdistotal.Text = FormatNumber(.Fields("d_total"), 2, vbTrue)
            txtcredittotal.Text = FormatNumber(.Fields("c_total"), 2, vbTrue)
            txtnettotal.Text = FormatNumber(.Fields("total amount paid"), 2, vbTrue)
            grandtotal = CSng(txtgtotal.Text)
            distotal = CSng(txtdistotal.Text)
            credittotal = CSng(txtcredittotal.Text)
            cmdremove.Enabled = True
            cmdclear.Enabled = True
            Dim i As Integer
            i = 1
            With Data1.rspurchases
                ' remove the filter if any
                .Filter = adFilterNone
                .MoveFirst
                Do While True
                    .Find "o_id='" & txto_id.Text & "'"
                    If .EOF Then Exit Do
                    list1.Rows = list1.Rows + 1
                    list1.TextMatrix(i, 1) = .Fields("t_code")
                    list1.TextMatrix(i, 2) = FormatNumber(.Fields("price"), 2, vbTrue)
                    list1.TextMatrix(i, 3) = .Fields("qty")
                    list1.TextMatrix(i, 4) = FormatNumber(.Fields("total"), 2, vbTrue)
                    list1.TextMatrix(i, 5) = FormatNumber(.Fields("dis"), 2, vbTrue)
                    list1.TextMatrix(i, 6) = FormatNumber(.Fields("credit"), 2, vbTrue)
                    list1.TextMatrix(i, 7) = FormatNumber(.Fields("net amount"), 2, vbTrue)
                    .MoveNext
                    i = i + 1
                Loop
                ' go to the first record after traversing all the records
                ' to avoid errors
                .MoveFirst
            End With
        End With
        cmdsave.Enabled = True
    Else
        grandtotal = 0: distotal = 0: credittotal = 0
        txto_id.Text = makecode()
        dtp1.Value = Date
    End If
    list1.TextMatrix(0, 1) = "Title Code"
    list1.TextMatrix(0, 2) = "Price"
    list1.TextMatrix(0, 3) = "Quantity"
    list1.TextMatrix(0, 4) = "Total"
    list1.TextMatrix(0, 5) = "Discount"
    list1.TextMatrix(0, 6) = "Credit Amount"
    list1.TextMatrix(0, 7) = "Net Amount"
End Sub

Private Sub cmdreset_Click()
    txtt_code.Text = ""
    txtt_name.Text = ""
    txtprice.Text = ""
    txtstock.Text = ""
    txtqty.Text = ""
    txts_code.Text = ""
    txts_name.Text = ""
    txtaddr.Text = ""
    txts_credit.Text = ""
    txttotal.Text = ""
    txtdis.Text = ""
    txtdisdigits.Text = ""
    txtcredit.Text = ""
    txtcreditdigits.Text = ""
    txtnetamount.Text = ""
    txtgtotal.Text = ""
    txtdistotal.Text = ""
    txtcredittotal.Text = ""
    txtnettotal.Text = ""
    list1.Rows = 1
    cmdsupplier.Enabled = True
    dtp1.Enabled = True
    cmdremove.Enabled = False
    cmdclear.Enabled = False
    cmdsave.Enabled = False
    grandtotal = 0: distotal = 0: credittotal = 0
    dtp1.SetFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If titlesformisopen = False Then Data1.rstitles.Close
    If suppliersformisopen = False Then Data1.rssuppliers.Close
End Sub

Private Sub txtcredit_Change()
    txtprice_Change
End Sub

Private Sub txtcredit_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub


Private Sub txtcreditdigits_Change()
    txtprice_Change
End Sub

Private Sub txtcreditdigits_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub

Private Sub txtdis_Change()
    txtprice_Change
End Sub

Private Sub txtdis_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub

Private Sub txtdisdigits_Change()
    txtprice_Change
End Sub

Private Sub txtdisdigits_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub

Private Sub txtprice_Change()
    If Not (txtprice.Text = "" Or txtqty.Text = "") Then
        txttotal.Text = FormatNumber(CSng(txtprice.Text) * val(txtqty.Text), 2, vbTrue)
        If (txtdis.Text & txtdisdigits.Text) = "" And (txtcredit.Text & txtcreditdigits.Text) = "" Then
            txtnetamount.Text = FormatNumber(CSng(txttotal.Text), 2, vbTrue)
        ElseIf (txtdis.Text & txtdisdigits.Text) = "" Then
            txtnetamount.Text = FormatNumber(CSng(txttotal.Text) - CSng(txtcredit.Text & "." & txtcreditdigits.Text), 2, vbTrue)
        ElseIf (txtcredit.Text & txtcreditdigits.Text) = "" Then
            txtnetamount.Text = FormatNumber(CSng(txttotal.Text) - CSng(txtdis.Text & "." & txtdisdigits.Text), 2, vbTrue)
        Else
            txtnetamount.Text = FormatNumber(CSng(txttotal.Text) - CSng(txtdis.Text & "." & txtdisdigits.Text) - CSng(txtcredit.Text & "." & txtcreditdigits.Text), 2, vbTrue)
        End If
    End If
End Sub

Private Sub txtqty_Change()
    txtprice_Change
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub
Private Function validate() As Boolean
    If txtt_code.Text = "" Then
        MsgBox "Title Code is blank", vbInformation
        txtt_code.SetFocus
        validate = False
        Exit Function
    End If
    If txtqty.Text = "" Then
        MsgBox "Quantity is blank", vbInformation
        txtqty.SetFocus
        validate = False
        Exit Function
    End If
    If CSng(txtnetamount.Text) < 0 Then
        MsgBox "Net Amount can't be negative", vbInformation
        validate = False
        Exit Function
    End If
    If txts_code.Text = "" Then
        MsgBox "Supplier Code is blank", vbInformation
        txts_code.SetFocus
        validate = False
        Exit Function
    End If
    
    validate = True
End Function
