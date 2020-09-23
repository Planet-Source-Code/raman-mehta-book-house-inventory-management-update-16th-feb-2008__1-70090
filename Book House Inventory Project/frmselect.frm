VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmselect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Record"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9255
   Icon            =   "frmselect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   9255
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdselect 
      Caption         =   "&Select"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   5040
      Width           =   975
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6000
      TabIndex        =   21
      Top             =   4920
      Width           =   3015
      Begin VB.CommandButton cmdnavigate 
         Height          =   265
         Index           =   3
         Left            =   1860
         Picture         =   "frmselect.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   400
      End
      Begin VB.CommandButton cmdnavigate 
         Height          =   265
         Index           =   2
         Left            =   1440
         Picture         =   "frmselect.frx":0C54
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   400
      End
      Begin VB.CommandButton cmdnavigate 
         Height          =   265
         Index           =   1
         Left            =   1020
         Picture         =   "frmselect.frx":0FDE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   400
      End
      Begin VB.CommandButton cmdnavigate 
         Height          =   265
         Index           =   0
         Left            =   600
         Picture         =   "frmselect.frx":1368
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   400
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Record"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lblcurrec 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   720
         TabIndex        =   24
         Top             =   360
         Width           =   90
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "of"
         Height          =   195
         Left            =   1560
         TabIndex        =   23
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lblttlrec 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   2160
         TabIndex        =   22
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Searc&h"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   5040
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sort Records"
      Height          =   2055
      Left            =   360
      TabIndex        =   19
      Top             =   5520
      Width           =   3015
      Begin VB.ComboBox cbosort 
         Height          =   315
         ItemData        =   "frmselect.frx":16F2
         Left            =   240
         List            =   "frmselect.frx":16F4
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ascending"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Descending"
         Height          =   495
         Index           =   1
         Left            =   1680
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmselect.frx":16F6
         Top             =   200
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Sort Records By:"
         Height          =   255
         Left            =   650
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filter Records"
      Height          =   2055
      Left            =   4320
      TabIndex        =   15
      Top             =   5520
      Width           =   4815
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
         TabIndex        =   11
         Top             =   930
         Width           =   2535
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
         ItemData        =   "frmselect.frx":1FC0
         Left            =   240
         List            =   "frmselect.frx":1FC2
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1530
         Width           =   2535
      End
      Begin VB.CommandButton cmdfilter 
         Caption         =   "&Filter"
         Height          =   495
         Left            =   3000
         TabIndex        =   13
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdremfilter 
         Caption         =   "&Remove Filter"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3000
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
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
         TabIndex        =   18
         Top             =   690
         Width           =   975
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
         TabIndex        =   17
         Top             =   1290
         Width           =   975
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   120
         Picture         =   "frmselect.frx":1FC4
         Top             =   240
         Width           =   480
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
         Left            =   705
         TabIndex        =   16
         Top             =   240
         Width           =   3435
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   15920516
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form displays different records dynamically based on the parameters
' passed from the calling modules
' used for selection of records while making sale or purchase transaction
' for example selecting the book title which is to be sold
' provides only read-only functionality
' records cannot be selected
Dim order As String * 4
Public WithEvents srsrs As ADODB.Recordset
Attribute srsrs.VB_VarHelpID = -1
Public srsarr As Variant
Public srsdata As Variant
Public recordselected As Boolean

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub srsrs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    updatelabels
End Sub
Private Sub cbosort_Click()
    With srsrs
        .Sort = .Fields(cbosort.ItemData(cbosort.ListIndex)).Name & " " & order
    End With
End Sub

Private Sub cmdfilter_Click()
    With srsrs
        .Filter = .Fields(cbofield.ItemData(cbofield.ListIndex)).Name & "='" & txtvalue.Text & "'"
        If .RecordCount = 0 Then enablecontrols False
    End With
    cmdremfilter.Enabled = True
End Sub
Private Sub cmdnavigate_Click(Index As Integer)
    With srsrs
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

Private Sub cmdremfilter_Click()
    enablecontrols True
    cmdremfilter.Enabled = False
    srsrs.Filter = adFilterNone
End Sub

Private Sub cmdsearch_Click()
    With frmsearch
        Set .srsrs = srsrs
        .srsarr = srsarr
        .srsdata = srsdata
        .Show vbModal
    End With
End Sub

Private Sub cmdselect_Click()
    recordselected = True
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    cmdselect_Click
End Sub

Private Sub DataGrid1_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdselect_Click
End Sub

Private Sub Form_Load()
    Set DataGrid1.DataSource = srsrs
    For i = LBound(srsarr) To UBound(srsarr)
        DataGrid1.Columns(i).Caption = srsarr(i)
        cbosort.AddItem (srsarr(i))
        cbosort.ItemData(i) = srsdata(i)
        cbofield.List(i) = srsarr(i)
        cbofield.ItemData(i) = srsdata(i)
    Next i
    cbosort.ListIndex = 0
    cbofield.ListIndex = 0
    updatelabels
    If srsrs.RecordCount = 0 Then enablecontrols False
    recordselected = False
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0: order = "ASC"
        Case 1: order = "DESC"
    End Select
    cbosort_Click
End Sub
Sub enablecontrols(val As Boolean)
    cmdsearch.Enabled = val
    cbosort.Enabled = val
    cmdselect.Enabled = val
    Option1(0).Enabled = val
    Option1(1).Enabled = val
    cmdfilter.Enabled = val
    DataGrid1.Enabled = val
    cmdnavigate(0).Enabled = val
    cmdnavigate(1).Enabled = val
    cmdnavigate(2).Enabled = val
    cmdnavigate(3).Enabled = val
End Sub


Private Sub updatelabels()
    With srsrs
        If .RecordCount = 0 Then
            lblcurrec.Caption = "0"
        Else
            lblcurrec.Caption = .AbsolutePosition
        End If
        lblttlrec.Caption = .RecordCount
    End With
End Sub

