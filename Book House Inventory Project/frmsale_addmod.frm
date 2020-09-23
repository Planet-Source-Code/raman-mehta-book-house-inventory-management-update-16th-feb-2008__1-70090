VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsale_addmod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Sale"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11040
   Icon            =   "frmsale_addmod.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11040
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   685
      Left            =   120
      TabIndex        =   50
      Top             =   0
      Width           =   10800
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
      Begin MSComCtl2.DTPicker dtp1 
         Height          =   300
         Left            =   8490
         TabIndex        =   1
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3735555
         CurrentDate     =   39335
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bill No.:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   52
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date:"
         Height          =   195
         Left            =   7920
         TabIndex        =   51
         Top             =   270
         Width           =   390
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   38
      Top             =   720
      Width           =   10800
      Begin VB.TextBox txtdigits 
         Height          =   300
         Left            =   1710
         MaxLength       =   2
         TabIndex        =   10
         Top             =   1065
         Width           =   540
      End
      Begin VB.TextBox txts_price 
         Height          =   300
         Left            =   220
         MaxLength       =   6
         TabIndex        =   9
         Top             =   1065
         Width           =   1305
      End
      Begin VB.TextBox txtprice 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   420
         Width           =   1305
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
      Begin VB.TextBox txtstock 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   7880
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   420
         Width           =   900
      End
      Begin VB.CommandButton cmdtitle 
         Caption         =   "..."
         Height          =   300
         Left            =   1680
         TabIndex        =   3
         Top             =   420
         Width           =   375
      End
      Begin VB.TextBox txtqty 
         Height          =   300
         Left            =   9225
         MaxLength       =   2
         TabIndex        =   7
         Top             =   420
         Width           =   510
      End
      Begin VB.TextBox txttotal 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1060
         Width           =   1635
      End
      Begin VB.TextBox txtcredit 
         Height          =   300
         Left            =   6730
         MaxLength       =   6
         TabIndex        =   14
         Top             =   1060
         Width           =   1380
      End
      Begin VB.TextBox txtdis 
         Height          =   300
         Left            =   4550
         MaxLength       =   5
         TabIndex        =   12
         Top             =   1060
         Width           =   1140
      End
      Begin VB.TextBox txtnetamount 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   9180
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1060
         Width           =   1380
      End
      Begin VB.TextBox txtdisdigits 
         Height          =   300
         Left            =   5870
         MaxLength       =   2
         TabIndex        =   13
         Top             =   1060
         Width           =   540
      End
      Begin VB.TextBox txtcreditdigits 
         Height          =   300
         Left            =   8290
         MaxLength       =   2
         TabIndex        =   15
         Top             =   1060
         Width           =   540
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   9735
         TabIndex        =   8
         Top             =   420
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtqty"
         BuddyDispid     =   196621
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
      Begin VB.Label Label8 
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
         Left            =   1560
         TabIndex        =   58
         Top             =   720
         Width           =   135
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Selling Price:"
         Height          =   195
         Index           =   6
         Left            =   220
         TabIndex        =   57
         Top             =   840
         Width           =   915
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cost Price:"
         Height          =   195
         Index           =   5
         Left            =   6120
         TabIndex        =   49
         Top             =   195
         Width           =   765
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Title Name:"
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   48
         Top             =   200
         Width           =   810
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Title Code:"
         Height          =   195
         Index           =   1
         Left            =   220
         TabIndex        =   47
         Top             =   200
         Width           =   765
      End
      Begin VB.Label lblstock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Stock:"
         Height          =   195
         Left            =   7880
         TabIndex        =   46
         Top             =   195
         Width           =   465
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Quantity:"
         Height          =   195
         Index           =   7
         Left            =   9225
         TabIndex        =   45
         Top             =   195
         Width           =   630
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Index           =   8
         Left            =   2580
         TabIndex        =   44
         Top             =   840
         Width           =   405
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Net Amount:"
         Height          =   195
         Index           =   11
         Left            =   9180
         TabIndex        =   43
         Top             =   840
         Width           =   885
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Credit Amount:"
         Height          =   195
         Index           =   10
         Left            =   6730
         TabIndex        =   42
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Discount:"
         Height          =   195
         Index           =   9
         Left            =   4550
         TabIndex        =   41
         Top             =   840
         Width           =   675
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
         Left            =   5720
         TabIndex        =   40
         Top             =   720
         Width           =   135
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
         Left            =   8140
         TabIndex        =   39
         Top             =   720
         Width           =   135
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1400
      Left            =   120
      TabIndex        =   33
      Top             =   2340
      Width           =   10800
      Begin VB.TextBox txtc_name 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   7245
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtc_code 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   345
         Width           =   1455
      End
      Begin VB.CommandButton cmdcustomer 
         Caption         =   "..."
         Height          =   300
         Left            =   2970
         TabIndex        =   18
         Top             =   345
         Width           =   375
      End
      Begin VB.TextBox txtc_credit 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   840
         Width           =   1380
      End
      Begin VB.TextBox txtaddr 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Customer Name:"
         Height          =   195
         Index           =   4
         Left            =   5925
         TabIndex        =   37
         Top             =   405
         Width           =   1170
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Customer Code:"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   36
         Top             =   390
         Width           =   1125
      End
      Begin VB.Label lbls_credit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Credit Balance:"
         Height          =   195
         Left            =   7995
         TabIndex        =   35
         Top             =   885
         Width           =   1080
      End
      Begin VB.Label lbladdr 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address:"
         Height          =   195
         Left            =   640
         TabIndex        =   34
         Top             =   885
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   29
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4920
      TabIndex        =   30
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add to &List"
      Height          =   495
      Left            =   9710
      TabIndex        =   22
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdremove 
      Caption         =   "Remove &from List"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9710
      TabIndex        =   23
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear &All"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9710
      TabIndex        =   24
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox txtgtotal 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox txtdistotal 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox txtcredittotal 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox txtnettotal 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "&Reset"
      Height          =   495
      Left            =   6600
      TabIndex        =   31
      Top             =   7080
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid list1 
      Height          =   2175
      Left            =   120
      TabIndex        =   32
      Top             =   4080
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      BackColorBkg    =   13479683
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Grand Total:"
      Height          =   195
      Left            =   480
      TabIndex        =   56
      Top             =   6360
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Discount:"
      Height          =   195
      Left            =   3000
      TabIndex        =   55
      Top             =   6360
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total Credit:"
      Height          =   195
      Left            =   5520
      TabIndex        =   54
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Net Amount:"
      Height          =   195
      Left            =   8040
      TabIndex        =   53
      Top             =   6360
      Width           =   885
   End
End
Attribute VB_Name = "frmsale_addmod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' this form is used to make new sale transactions as well as modifictions
' to the existing records
' it provides functionality similar to the purchaseaddmod form
Public modstate As Boolean
Dim grandtotal As Single, distotal As Single, credittotal As Single
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdadd_Click()
    Dim currow
    If validate() = False Then Exit Sub
    For i = 1 To list1.Rows - 1 Step 1
        If txtt_code.Text = list1.TextMatrix(i, 1) Then
            MsgBox "Title already exists in the list", vbInformation
            Exit Sub
        End If
    Next i
    If CSng(txts_price.Text & "." & txtdigits.Text) < CSng(txtprice.Text) Then
        If MsgBox("Selling Price is less than Cost Price. Proceed any way?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    list1.Rows = list1.Rows + 1
    currow = list1.Rows - 1
    If cmdremove.Enabled = False Then
        cmdcustomer.Enabled = False
        dtp1.Enabled = False
        cmdremove.Enabled = True
        cmdclear.Enabled = True
        cmdsave.Enabled = True
    End If
    list1.TextMatrix(currow, 1) = txtt_code.Text
    list1.TextMatrix(currow, 2) = FormatNumber(txtprice.Text, 2, vbTrue)
    list1.TextMatrix(currow, 3) = FormatNumber(CSng(txts_price.Text & "." & txtdigits.Text), 2, vbTrue)
    list1.TextMatrix(currow, 4) = txtqty.Text
    list1.TextMatrix(currow, 5) = FormatNumber(txttotal.Text, 2, vbTrue)
    If txtdis.Text & txtdisdigits.Text = "" Then
        list1.TextMatrix(currow, 6) = "0.00"
    Else
        list1.TextMatrix(currow, 6) = FormatNumber(CSng(txtdis.Text & "." & txtdisdigits.Text), 2, vbTrue)
    End If
    If txtcredit.Text & txtcreditdigits.Text = "" Then
        list1.TextMatrix(currow, 7) = "0.00"
    Else
        list1.TextMatrix(currow, 7) = FormatNumber(CSng(txtcredit.Text & "." & txtcreditdigits.Text), 2, vbTrue)
    End If
    list1.TextMatrix(currow, 8) = FormatNumber(txtnetamount.Text, 2, vbTrue)
    grandtotal = grandtotal + CSng(list1.TextMatrix(currow, 5))
    distotal = distotal + CSng(list1.TextMatrix(currow, 6))
    credittotal = credittotal + CSng(list1.TextMatrix(currow, 7))
    txtgtotal.Text = FormatNumber(grandtotal, 2, vbTrue)
    txtdistotal.Text = FormatNumber(distotal, 2, vbTrue)
    txtcredittotal.Text = FormatNumber(credittotal, 2, vbTrue)
    txtnettotal.Text = FormatNumber(CSng(txtgtotal.Text) - CSng(txtdistotal.Text) - CSng(txtcredittotal.Text), 2, vbTrue)
End Sub
Private Sub cmdclear_Click()
    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
        list1.Rows = 1
        grandtotal = 0: distotal = 0: credittotal = 0
        txtgtotal.Text = ""
        txtdistotal.Text = ""
        txtcredittotal.Text = ""
        txtnettotal.Text = ""
        cmdcustomer.Enabled = True
        dtp1.Enabled = True
        cmdremove.Enabled = False
        cmdclear.Enabled = False
        cmdsave.Enabled = False
    End If
End Sub

Private Sub cmdremove_Click()
    Dim start As Byte, finish As Byte
    With list1
        If .Row < .RowSel Then
            start = .Row: finish = .RowSel
        Else
            start = .RowSel: finish = .Row
        End If
        For i = start To finish
            If .Rows = 2 Then
                .Rows = 1
                grandtotal = 0: distotal = 0: credittotal = 0
                txtgtotal.Text = ""
                txtdistotal.Text = ""
                txtcredittotal.Text = ""
                txtnettotal.Text = ""
                cmdcustomer.Enabled = True
                dtp1.Enabled = True
                cmdremove.Enabled = False
                cmdclear.Enabled = False
                cmdsave.Enabled = False
            Else
                grandtotal = grandtotal - CSng(list1.TextMatrix(.Row, 5))
                distotal = distotal - CSng(list1.TextMatrix(.Row, 6))
                credittotal = credittotal - CSng(list1.TextMatrix(.Row, 7))
                txtgtotal.Text = FormatNumber(CSng(txtgtotal.Text) - CSng(list1.TextMatrix(.Row, 5)), 2, vbTrue)
                txtdistotal.Text = FormatNumber(CSng(txtdistotal.Text) - CSng(list1.TextMatrix(.Row, 6)), 2, vbTrue)
                txtcredittotal.Text = FormatNumber(CSng(txtcredittotal.Text) - CSng(list1.TextMatrix(.Row, 7)), 2, vbTrue)
                txtnettotal.Text = FormatNumber(CSng(txtnettotal.Text) - CSng(list1.TextMatrix(.Row, 8)), 2, vbTrue)
                .RemoveItem (.Row)
            End If
        Next i
    End With
    End Sub

Private Sub cmdsave_Click()
    On Error GoTo err:
    Dim rs As New ADODB.Recordset
    Data1.conn.BeginTrans
    With Data1.rst_sales
        If (.State = adStateClosed) Then: .Open
        If modstate = False Then
            .AddNew
        Else
            rs.Open "select * from saledetails where o_id='" & txto_id.Text & "'", Data1.conn, adOpenForwardOnly, adLockOptimistic
            While Not rs.EOF
                With Data1.rstitles
                    ' remove the filter if any
                    .Filter = adFilterNone
                    .MoveFirst
                    .Find "t_code='" & rs.Fields("t_code") & "'"
                    .Fields("stock") = .Fields("stock") + rs.Fields("qty")
                End With
                With Data1.rscustomers
                    ' remove the filter if any
                    .Filter = adFilterNone
                    .MoveFirst
                    .Find "c_code='" & txtc_code.Text & "'"
                    .Fields("credit") = .Fields("credit") - rs.Fields("credit")
                End With
                rs.MoveNext
            Wend
            rs.Close
            Set frmsales.DataGrid1.DataSource = Nothing
            Data1.conn.Execute "delete from saledetails where o_id='" & Data1.rst_sales.Fields(0) & "'"
            Data1.rssales.Close
            Data1.rssales.Open
            Set frmsales.DataGrid1.DataSource = Data1
        End If
        'update the saledetails table and cusotmers table with new information
        ' but first remove the filters if any
        Data1.rstitles.Filter = adFilterNone
        Data1.rscustomers.Filter = adFilterNone
        For i = 1 To list1.Rows - 1
            With Data1.rstitles
                .MoveFirst
                .Find "t_code='" & list1.TextMatrix(i, 1) & "'"
                .Fields("stock") = .Fields("stock") - val(list1.TextMatrix(i, 4))
                .Update
            End With
        Next i
        With Data1.rscustomers
            .MoveFirst
            .Find "c_code='" & txtc_code.Text & "'"
            .Fields("credit") = .Fields("credit") + credittotal
            .Update
        End With
        .Fields("o_id") = txto_id.Text
        .Fields("c_code") = txtc_code.Text
        .Fields("s_date") = dtp1.Value
        .Fields("g_total") = grandtotal
        .Fields("d_total") = distotal
        .Fields("c_total") = credittotal
        .Update
    End With
    With Data1.rssales
        If (.State = adStateClosed) Then: .Open
        ' remove the filter if any
        .Filter = adFilterNone
        For i = 1 To list1.Rows - 1
              .AddNew
              .Fields("o_id") = txto_id.Text
              .Fields("t_code") = list1.TextMatrix(i, 1)
              .Fields("price") = CSng(list1.TextMatrix(i, 2))
              .Fields("s_price") = CSng(list1.TextMatrix(i, 3))
              .Fields("qty") = list1.TextMatrix(i, 4)
              .Fields("total") = CSng(list1.TextMatrix(i, 5))
              .Fields("dis") = CSng(list1.TextMatrix(i, 6))
              .Fields("credit") = CSng(list1.TextMatrix(i, 7))
              .Update
        Next i
     End With
     If modstate = False Then
        If salesformisopen Then
            With Data1.rst_sales
                Set frmsales.DataGrid2.DataSource = Nothing
                .Close
                .Open
                Set frmsales.DataGrid2.DataSource = Data1
                If (.Filter <> adFilterNone) Then
                    frmsales.cmdremfilter2_Click
                End If
                .Find "o_id='" & txto_id.Text & "'"
            End With
            With Data1.rssales
                Set frmsales.DataGrid1.DataSource = Nothing
                .Close
                .Open
                Set frmsales.DataGrid1.DataSource = Data1
                If (.Filter <> adFilterNone) Then
                    frmsales.cmdremfilter_Click
                End If
                .Find "o_id='" & txto_id.Text & "'"
            End With
        End If
        ' print the Sale Invoice
        printrep
        If MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            cmdreset_Click
            txto_id.Text = makecode()
        Else
            If Not salesformisopen Then
                Data1.rst_sales.Close
                Data1.rssales.Close
            Else
                If frmsales.cmdDelete.Enabled = False Then
                    frmsales.enablecontrols True
                    frmsales.enablecontrols2 True
                End If
            End If
            Unload Me
        End If
    Else
        With Data1.rst_sales
            Set frmsales.DataGrid2.DataSource = Nothing
            .Close
            .Open
            Set frmsales.DataGrid2.DataSource = Data1
            frmsales.cmdremfilter2_Click
            .Find "o_id='" & txto_id.Text & "'"
        End With
        With Data1.rssales
            bm = .Bookmark
            .Close
            .Open
            Set frmsales.DataGrid1.DataSource = Data1
            frmsales.cmdremfilter_Click
            .Find "o_id='" & txto_id.Text & "'"
        End With
        Unload Me
    End If
    ' deal with the filter and remove filter buttons on the customers and titles form
    If customersformisopen Then frmcustomers.cmdremfilter_Click
    If titlesformisopen Then frmtitles.cmdremfilter_Click
    Data1.conn.CommitTrans
Exit Sub
err:
    Data1.conn.RollbackTrans
    If salesformisopen Then
        Set frmsales.DataGrid1.DataSource = Nothing
        Set frmsales.DataGrid2.DataSource = Nothing
        Data1.rssales.CancelUpdate
        Data1.rst_sales.CancelUpdate
        Set frmsales.DataGrid1.DataSource = Data1
        Set frmsales.DataGrid2.DataSource = Data1
    Else
        Data1.rssales.CancelUpdate
        Data1.rst_sales.CancelUpdate
    End If
    If customersformisopen Then
        Set frmcustomers.DataGrid1.DataSource = Nothing
        Data1.rscustomers.CancelUpdate
        Set frmcustomers.DataGrid1.DataSource = Data1
    Else
        Data1.rscustomers.CancelUpdate
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
    rs.Open "select max(mid(o_id,2))as maxcode from sales", Data1.conn
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
Private Sub cmdcustomer_Click()
    With frmselect
        Set .srsrs = Data1.conn.Execute("SELECT c_code,c_name,c_addr,c_cont_no,fax_no,email_addr,credit FROM customers")
        .srsarr = Array("Customer Code", "Customer Name", "Address", "Contact Number", "Fax Number", "Email Address", "Credit Amount")
        .srsdata = Array(0, 1, 2, 3, 4, 5, 6)
        .Show vbModal
        If (.recordselected = True) Then
            txtc_code.Text = .srsrs.Fields(0)
            txtc_name.Text = .srsrs.Fields(1)
            txtaddr.Text = .srsrs.Fields(2)
            txtc_credit.Text = FormatNumber(.srsrs.Fields(6), 2, vbTrue)
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
            txts_price.Text = "": txtdigits.Text = ""
            txttotal.Text = ""
            txtdis.Text = "": txtdisdigits.Text = ""
            txtcredit.Text = "": txtcreditdigits.Text = ""
            txtnetamount.Text = ""
        End If
        .srsrs.Close
    End With
End Sub
Private Sub printrep()
    On Error GoTo err:
    Data1.rsrptsales.Filter = "o_id='" & txto_id.Text & "'"
    Load rptsales
    Set rptsales.DataSource = Data1
    MsgBox "The Invoice would print now", vbInformation
    rptsales.PrintReport
    Unload rptsales
    Exit Sub
err:
    MsgBox "Print Error"
    Unload rptsales
End Sub

Private Sub Form_Load()
    If (Data1.rstitles.State = adStateClosed) Then
            Data1.rstitles.Open
    End If
    If Data1.rscustomers.State = adStateClosed Then
        Data1.rscustomers.Open
    End If
    If (modstate = True) Then
        With Data1.rst_sales
            txto_id.Text = .Fields("o_id")
            dtp1.Value = .Fields("s_date")
            txtc_code.Text = .Fields("c_code")
            txtc_name.Text = .Fields("c_name")
            txtaddr.Text = .Fields("c_addr")
            txtc_credit.Text = FormatNumber(.Fields("credit"), 2, vbTrue)
            cmdcustomer.Enabled = False
            dtp1.Enabled = False
            txtgtotal.Text = FormatNumber(.Fields("g_total"), 2, vbTrue)
            txtdistotal.Text = FormatNumber(.Fields("d_total"), 2, vbTrue)
            txtcredittotal.Text = FormatNumber(.Fields("c_total"), 2, vbTrue)
            txtnettotal.Text = FormatNumber(.Fields("total amount received"), 2, vbTrue)
            grandtotal = CSng(txtgtotal.Text)
            distotal = CSng(txtdistotal.Text)
            credittotal = CSng(txtcredittotal.Text)
            cmdremove.Enabled = True
            cmdclear.Enabled = True
            Dim i As Integer
            i = 1
            With Data1.rssales
                ' remove the filter if any
                .Filter = adFilterNone
                .MoveFirst
                Do While True
                    .Find "o_id='" & txto_id.Text & "'"
                    If .EOF Then Exit Do
                    list1.Rows = list1.Rows + 1
                    list1.TextMatrix(i, 1) = .Fields("t_code")
                    list1.TextMatrix(i, 2) = FormatNumber(.Fields("price"), 2, vbTrue)
                    list1.TextMatrix(i, 3) = FormatNumber(.Fields("s_price"), 2, vbTrue)
                    list1.TextMatrix(i, 4) = .Fields("qty")
                    list1.TextMatrix(i, 5) = FormatNumber(.Fields("total"), 2, vbTrue)
                    list1.TextMatrix(i, 6) = FormatNumber(.Fields("dis"), 2, vbTrue)
                    list1.TextMatrix(i, 7) = FormatNumber(.Fields("credit"), 2, vbTrue)
                    list1.TextMatrix(i, 8) = FormatNumber(.Fields("net amount"), 2, vbTrue)
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
        cmdsave.Caption = "&Save and Print"
        dtp1.Value = Date
    End If
    list1.TextMatrix(0, 1) = "Title Code"
    list1.TextMatrix(0, 2) = "Price"
    list1.TextMatrix(0, 3) = "Selling Price"
    list1.TextMatrix(0, 4) = "Quantity"
    list1.TextMatrix(0, 5) = "Total"
    list1.TextMatrix(0, 6) = "Discount"
    list1.TextMatrix(0, 7) = "Credit Amount"
    list1.TextMatrix(0, 8) = "Net Amount"
End Sub

Private Sub cmdreset_Click()
    txtt_code.Text = ""
    txtt_name.Text = ""
    txtprice.Text = ""
    txtstock.Text = ""
    txts_price.Text = ""
    txtdigits.Text = ""
    txtqty.Text = ""
    txtc_code.Text = ""
    txtc_name.Text = ""
    txtaddr.Text = ""
    txtc_credit.Text = ""
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
    cmdcustomer.Enabled = True
    dtp1.Enabled = True
    cmdremove.Enabled = False
    cmdclear.Enabled = False
    cmdsave.Enabled = False
    grandtotal = 0: distotal = 0: credittotal = 0
    dtp1.SetFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If titlesformisopen = False Then Data1.rstitles.Close
    If customersformisopen = False Then Data1.rscustomers.Close
End Sub

Private Sub txtcredit_Change()
    txts_price_Change
End Sub

Private Sub txtcredit_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub


Private Sub txtcreditdigits_Change()
    txts_price_Change
End Sub

Private Sub txtcreditdigits_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub

Private Sub txtdigits_Change()
    txts_price_Change
End Sub

Private Sub txtdigits_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub

Private Sub txtdis_Change()
    txts_price_Change
End Sub

Private Sub txtdis_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub


Private Sub txtdisdigits_Change()
    txts_price_Change
End Sub

Private Sub txtdisdigits_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub

Private Sub txts_price_Change()
    If Not ((txts_price.Text & txtdigits.Text) = "" Or txtqty.Text = "") Then
        txttotal.Text = FormatNumber(CSng(txts_price.Text & "." & txtdigits.Text) * val(txtqty.Text), 2, vbTrue)
        If (txtdis.Text & txtdisdigits.Text) = "" And (txtcredit.Text & txtcreditdigits.Text) = "" Then
            txtnetamount.Text = FormatNumber(CSng(txttotal.Text), 2, vbTrue)
        ElseIf (txtdis.Text & txtdisdigits.Text) = "" Then
            txtnetamount.Text = FormatNumber(CSng(txttotal.Text) - CSng(txtcredit.Text & "." & txtcreditdigits.Text), 2, vbTrue)
        ElseIf (txtcredit.Text & txtcreditdigits.Text) = "" Then
            txtnetamount.Text = FormatNumber(CSng(txttotal.Text) - CSng(txtdis.Text & "." & txtdisdigits.Text), 2, vbTrue)
        Else
            txtnetamount.Text = FormatNumber(CSng(txttotal.Text) - CSng(txtdis.Text & "." & txtdisdigits.Text) - CSng(txtcredit.Text & "." & txtcreditdigits.Text), 2, vbTrue)
        End If
    Else
        ' clear the text boxes if user does not fill any selling price
        txttotal.Text = ""
        txtnetamount.Text = ""
    End If
End Sub
Private Sub txts_price_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack Or KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub

Private Sub txtqty_Change()
    txts_price_Change
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
    If val(txtqty.Text) > val(txtstock.Text) Then
        MsgBox "Quantity can't be greater than Stock", vbInformation
        txtqty.SetFocus
        validate = False
        Exit Function
    End If
    If (txts_price.Text & txtdigits.Text) = "" Then
        MsgBox "Selling Price is blank", vbInformation
        txts_price.SetFocus
        validate = False
        Exit Function
    End If
    If CSng(txtnetamount.Text) < 0 Then
        MsgBox "Net Amount can't be negative", vbInformation
        validate = False
        Exit Function
    End If
    If txtc_code.Text = "" Then
        MsgBox "Customer Code is blank", vbInformation
        txtc_code.SetFocus
        validate = False
        Exit Function
    End If
    validate = True
End Function


