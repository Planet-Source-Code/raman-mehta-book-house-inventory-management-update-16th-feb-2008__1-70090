VERSION 5.00
Begin VB.Form frmsearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4155
   Icon            =   "frmsearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4155
   StartUpPosition =   1  'CenterOwner
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
      Left            =   120
      TabIndex        =   0
      Top             =   855
      Width           =   3855
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.ComboBox cbosearch 
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
      ItemData        =   "frmsearch.frx":08CA
      Left            =   120
      List            =   "frmsearch.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1575
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   120
      Picture         =   "frmsearch.frx":08CE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   675
   End
   Begin VB.Label Label3 
      Caption         =   "Search for:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   615
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a text which you want to search and select a field where to locate it."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   840
      TabIndex        =   5
      Top             =   45
      Width           =   3105
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   4
      Top             =   1335
      Width           =   1815
   End
End
Attribute VB_Name = "frmsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------
'Search Form
'This reusable form can be called from anywhere to search a recordset
'This code has been reused from various forms
'----------------------------------------------------------------------

Public srsrs As ADODB.Recordset
Public srsarr As Variant
Public srsdata As Variant
Dim alreadysearched As Boolean
Dim bm As Variant
Dim datatype As DataTypeEnum

Private Sub cmdcancel_Click()
  Unload Me
End Sub

Private Sub cmdsearch_Click()
    On Error GoTo err
    ' determine the data type of the field to be searched
    datatype = srsrs.Fields(srsdata(cbosearch.ListIndex)).Type
    If txtvalue.Text = "" Then
        txtvalue.SetFocus: Exit Sub
    End If
    With srsrs
        If alreadysearched = False Then
            ' if the recordset is searched for the first time for
            ' for the newly entered value then
            .MoveFirst
            ' search according to the datatype determined
            If datatype = adVarWChar Then
                .Find .Fields(srsdata(cbosearch.ListIndex)).Name & " like '*" & txtvalue.Text & "*'"
            ElseIf datatype = adSingle Then
                .Find .Fields(srsdata(cbosearch.ListIndex)).Name & "=" & txtvalue.Text
            ElseIf datatype = adDate Then
                .Find .Fields(srsdata(cbosearch.ListIndex)).Name & "=#" & txtvalue.Text & "#"
            End If
            ' if end of recordset is reached then it means record was not found
            ' then goto record whose position was stored
            If .EOF Then
                .Bookmark = bm
                MsgBox "Could not find '" & txtvalue.Text & "' in '" & cbosearch.Text & "'.", vbExclamation
            Else
                alreadysearched = True
                bm = .Bookmark
                cmdsearch.Caption = "Search Next"
            End If
        Else
            ' continue searching next
            .MoveNext
            If datatype = adVarWChar Then
                .Find .Fields(srsdata(cbosearch.ListIndex)).Name & " like '*" & txtvalue.Text & "*'"
            ElseIf datatype = adSingle Then
                .Find .Fields(srsdata(cbosearch.ListIndex)).Name & "=" & txtvalue.Text
            ElseIf datatype = adDate Then
                .Find .Fields(srsdata(cbosearch.ListIndex)).Name & "=#" & txtvalue.Text & "#"
            End If
            ' if end of recordset is reached it means search is completed
            If .EOF Then
                .Bookmark = bm
                MsgBox "Search completed.", vbInformation
             End If
            bm = .Bookmark
        End If
    End With
    Exit Sub
err:
    MsgBox "please enter a valid value in 'Search for:' text box", vbExclamation
End Sub

Private Sub Form_Load()
    ' fill the combobox with the required fields dynamically
    ' based on the array srsarr which is passed values from the calling module
    FillCombo cbosearch, srsarr, srsdata
    cbosearch.ListIndex = 0
    ' record the position of the current record of the recordset
    bm = srsrs.Bookmark
    alreadysearched = False
End Sub

Private Sub txtvalue_Change()
    alreadysearched = False
    cmdsearch.Caption = "Search"
End Sub


