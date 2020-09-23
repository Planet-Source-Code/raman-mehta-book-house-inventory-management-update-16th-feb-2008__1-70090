VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmpurchases 
   Caption         =   "Purchases"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmpurchases.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame framescroll 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10440
      TabIndex        =   39
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1335
      LargeChange     =   400
      Left            =   5280
      SmallChange     =   200
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   400
      Left            =   2280
      SmallChange     =   200
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13573
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Purchases"
      TabPicture(0)   =   "frmpurchases.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdclose"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdnew"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdsearch"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Total Purchases"
      TabPicture(1)   =   "frmpurchases.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmddelete"
      Tab(1).Control(1)=   "cmdsearch2"
      Tab(1).Control(2)=   "Frame7"
      Tab(1).Control(3)=   "cmddelall"
      Tab(1).Control(4)=   "cmdclose2"
      Tab(1).Control(5)=   "cmdmod"
      Tab(1).Control(6)=   "Frame6"
      Tab(1).Control(7)=   "Frame2"
      Tab(1).Control(8)=   "DataGrid2"
      Tab(1).ControlCount=   9
      Begin VB.CommandButton cmdsearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   4080
         TabIndex        =   1
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdnew 
         Caption         =   "&Add"
         Height          =   375
         Left            =   2640
         TabIndex        =   0
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   5520
         TabIndex        =   2
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmddelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   -71880
         TabIndex        =   17
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdsearch2 
         Caption         =   "Searc&h"
         Height          =   375
         Left            =   -69000
         TabIndex        =   19
         Top             =   6240
         Width           =   975
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -66360
         TabIndex        =   53
         Top             =   5520
         Width           =   3015
         Begin VB.CommandButton cmdnavigate2 
            Height          =   265
            Index           =   0
            Left            =   600
            Picture         =   "frmpurchases.frx":0902
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   0
            Width           =   400
         End
         Begin VB.CommandButton cmdnavigate2 
            Height          =   265
            Index           =   1
            Left            =   1020
            Picture         =   "frmpurchases.frx":0C8C
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   0
            Width           =   400
         End
         Begin VB.CommandButton cmdnavigate2 
            Height          =   265
            Index           =   2
            Left            =   1440
            Picture         =   "frmpurchases.frx":1016
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   0
            Width           =   400
         End
         Begin VB.CommandButton cmdnavigate2 
            Height          =   265
            Index           =   3
            Left            =   1860
            Picture         =   "frmpurchases.frx":13A0
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   0
            Width           =   400
         End
         Begin VB.Label lblttlrec2 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2160
            TabIndex        =   57
            Top             =   360
            Width           =   90
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "of"
            Height          =   195
            Left            =   1560
            TabIndex        =   56
            Top             =   360
            Width           =   135
         End
         Begin VB.Label lblcurrec2 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   720
            TabIndex        =   55
            Top             =   360
            Width           =   90
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Record"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Width           =   525
         End
      End
      Begin VB.CommandButton cmddelall 
         Caption         =   "Delete A&ll"
         Height          =   375
         Left            =   -70440
         TabIndex        =   18
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdclose2 
         Caption         =   "Cl&ose"
         Height          =   375
         Left            =   -67560
         TabIndex        =   20
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdmod 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   -73320
         TabIndex        =   16
         Top             =   6240
         Width           =   975
      End
      Begin VB.Frame Frame6 
         Caption         =   "Filter Records"
         Height          =   2775
         Left            =   -66360
         TabIndex        =   49
         Top             =   2400
         Width           =   3015
         Begin VB.CommandButton cmdremfilter2 
            Caption         =   "R&emove Filter"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1560
            TabIndex        =   27
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox txtvalue2 
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
            TabIndex        =   24
            Top             =   1170
            Width           =   2535
         End
         Begin VB.ComboBox cbofield2 
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
            ItemData        =   "frmpurchases.frx":172A
            Left            =   240
            List            =   "frmpurchases.frx":174C
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1770
            Width           =   2535
         End
         Begin VB.CommandButton cmdfilter2 
            Caption         =   "F&ilter"
            Height          =   375
            Left            =   240
            TabIndex        =   26
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label14 
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
            TabIndex        =   52
            Top             =   930
            Width           =   975
         End
         Begin VB.Label Label13 
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
            TabIndex        =   51
            Top             =   1530
            Width           =   975
         End
         Begin VB.Image Image4 
            Height          =   480
            Left            =   120
            Picture         =   "frmpurchases.frx":17DE
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label12 
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
            TabIndex        =   50
            Top             =   240
            Width           =   2235
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Filter Records"
         Height          =   2775
         Left            =   8640
         TabIndex        =   45
         Top             =   2400
         Width           =   3015
         Begin VB.CommandButton cmdfilter 
            Caption         =   "&Filter"
            Height          =   375
            Left            =   240
            TabIndex        =   8
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
            ItemData        =   "frmpurchases.frx":24A8
            Left            =   240
            List            =   "frmpurchases.frx":24D2
            Style           =   2  'Dropdown List
            TabIndex        =   7
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
            TabIndex        =   6
            Top             =   1170
            Width           =   2535
         End
         Begin VB.CommandButton cmdremfilter 
            Caption         =   "&Remove Filter"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1560
            TabIndex        =   9
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label4 
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
            TabIndex        =   48
            Top             =   240
            Width           =   2235
         End
         Begin VB.Image Image3 
            Height          =   480
            Left            =   120
            Picture         =   "frmpurchases.frx":2562
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
            TabIndex        =   47
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
            TabIndex        =   46
            Top             =   930
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   8640
         TabIndex        =   40
         Top             =   5520
         Width           =   3015
         Begin VB.CommandButton cmdnavigate 
            Height          =   265
            Index           =   3
            Left            =   1860
            Picture         =   "frmpurchases.frx":322C
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   400
         End
         Begin VB.CommandButton cmdnavigate 
            Height          =   265
            Index           =   2
            Left            =   1440
            Picture         =   "frmpurchases.frx":35B6
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   0
            Width           =   400
         End
         Begin VB.CommandButton cmdnavigate 
            Height          =   265
            Index           =   1
            Left            =   1020
            Picture         =   "frmpurchases.frx":3940
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   0
            Width           =   400
         End
         Begin VB.CommandButton cmdnavigate 
            Height          =   265
            Index           =   0
            Left            =   600
            Picture         =   "frmpurchases.frx":3CCA
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   0
            Width           =   400
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Record"
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   525
         End
         Begin VB.Label lblcurrec 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   720
            TabIndex        =   43
            Top             =   360
            Width           =   90
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "of"
            Height          =   195
            Left            =   1560
            TabIndex        =   42
            Top             =   360
            Width           =   135
         End
         Begin VB.Label lblttlrec 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2160
            TabIndex        =   41
            Top             =   360
            Width           =   90
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Sort Records"
         Height          =   1695
         Left            =   -66360
         TabIndex        =   35
         Top             =   480
         Width           =   3015
         Begin VB.OptionButton Option2 
            Alignment       =   1  'Right Justify
            Caption         =   "Descending"
            Height          =   495
            Index           =   1
            Left            =   1680
            TabIndex        =   23
            Top             =   1080
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Alignment       =   1  'Right Justify
            Caption         =   "Ascending"
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   22
            Top             =   1080
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.ComboBox cbosort2 
            Height          =   315
            ItemData        =   "frmpurchases.frx":4054
            Left            =   240
            List            =   "frmpurchases.frx":4076
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label2 
            Caption         =   "Sort Records By:"
            Height          =   255
            Left            =   650
            TabIndex        =   36
            Top             =   360
            Width           =   1215
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   120
            Picture         =   "frmpurchases.frx":4108
            Top             =   200
            Width           =   480
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sort Records"
         Height          =   1695
         Left            =   8640
         TabIndex        =   33
         Top             =   480
         Width           =   3015
         Begin VB.ComboBox cbosort 
            Height          =   315
            ItemData        =   "frmpurchases.frx":49D2
            Left            =   240
            List            =   "frmpurchases.frx":49FC
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   720
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "Ascending"
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   1080
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "Descending"
            Height          =   495
            Index           =   1
            Left            =   1680
            TabIndex        =   5
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   120
            Picture         =   "frmpurchases.frx":4A8C
            Top             =   200
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   "Sort Records By:"
            Height          =   255
            Left            =   650
            TabIndex        =   34
            Top             =   360
            Width           =   1215
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmpurchases.frx":5356
         Height          =   5535
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   9763
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
         DataMember      =   "purchases"
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "O_ID"
            Caption         =   "Bill Number"
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
            DataField       =   "P_DATE"
            Caption         =   "Bill Date"
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
            DataField       =   "T_CODE"
            Caption         =   "Title Code"
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
            DataField       =   "T_NAME"
            Caption         =   "Title Name"
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
         BeginProperty Column05 
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
         BeginProperty Column06 
            DataField       =   "PRICE"
            Caption         =   "Price"
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
         BeginProperty Column07 
            DataField       =   "QTY"
            Caption         =   "Quantity"
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
         BeginProperty Column08 
            DataField       =   "TOTAL"
            Caption         =   "Amount"
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
         BeginProperty Column09 
            DataField       =   "DIS"
            Caption         =   "Discount"
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
         BeginProperty Column10 
            DataField       =   "CREDIT"
            Caption         =   "Credit Amount"
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
         BeginProperty Column11 
            DataField       =   "net amount"
            Caption         =   "Net Amount"
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
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmpurchases.frx":536A
         Height          =   5535
         Left            =   -74760
         TabIndex        =   32
         Top             =   480
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   9763
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
         DataMember      =   "t_purchases"
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "O_ID"
            Caption         =   "Bill Number"
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
            DataField       =   "P_DATE"
            Caption         =   "Bill Date"
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "S_ADDR"
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
         BeginProperty Column05 
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
         BeginProperty Column06 
            DataField       =   "G_TOTAL"
            Caption         =   "Total Amount"
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
         BeginProperty Column07 
            DataField       =   "D_TOTAL"
            Caption         =   "Total Discount"
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
         BeginProperty Column08 
            DataField       =   "C_TOTAL"
            Caption         =   "Total Credit"
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
         BeginProperty Column09 
            DataField       =   "total amount paid"
            Caption         =   "Total Amount Paid"
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
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmpurchases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form shows the total purchases and detailed purchases
' and enables addition, deletions and modifications of the records
Dim vsval As Integer
Dim hsval As Integer
Dim order As String
Dim order2 As String
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Dim WithEvents rs2 As ADODB.Recordset
Attribute rs2.VB_VarHelpID = -1

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdclose2_Click()
    Unload Me
End Sub

Private Sub cmddelall_Click()
    On Error GoTo err
    If MsgBox("Are you sure you want to delete all the records?", vbQuestion + vbYesNo, "Confirm deletion") = vbNo Then
        Exit Sub
    End If
    Data1.conn.BeginTrans
    Set DataGrid1.DataSource = Nothing
    Set DataGrid2.DataSource = Nothing
    Data1.conn.Execute ("delete from purchases")
    Data1.rspurchases.Close
    Data1.rst_purchases.Close
    Data1.rspurchases.Open
    Data1.rst_purchases.Open
    Set DataGrid1.DataSource = Data1
    Set DataGrid2.DataSource = Data1
    Data1.conn.CommitTrans
    enablecontrols False
    enablecontrols2 False
    cmdremfilter.Enabled = False
    cmdremfilter2.Enabled = False
    Exit Sub
err:
    Data1.conn.RollbackTrans
    Data1.conn.Cancel
    Set DataGrid2.DataSource = Data1
    Set DataGrid1.DataSource = Data1
    MsgBox err.Description, vbCritical, "Error"
End Sub

Private Sub cmdfilter_Click()
    With Data1.rspurchases
        .Filter = .Fields(cbofield.ItemData(cbofield.ListIndex)).Name & "='" & txtvalue.Text & "'"
        If .RecordCount = 0 Then enablecontrols False
    End With
    cmdremfilter.Enabled = True
End Sub

Private Sub cmdfilter2_Click()
    With Data1.rst_purchases
        .Filter = .Fields(cbofield2.ItemData(cbofield2.ListIndex)).Name & "='" & txtvalue2.Text & "'"
        If .RecordCount = 0 Then enablecontrols2 False
    End With
    cmdremfilter2.Enabled = True
End Sub

Private Sub cmdsearch2_Click()
    With frmsearch
        Set .srsrs = Data1.rst_purchases
        .srsarr = Array("Bill Number", "Bill Date", "Supplier Code", "Supplier Name", "Address", "Credit Balance", "Total Amount", "Total Discount", "Total Credit", "Total Amount Paid")
        .srsdata = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
        .Show vbModal
    End With
End Sub

Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If purchasesformisopen Then updatelabels
End Sub
Private Sub rs2_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If purchasesformisopen Then updatelabels2
End Sub
Private Sub cbosort_Click()
    With Data1.rspurchases
        .Sort = .Fields(cbosort.ItemData(cbosort.ListIndex)).Name & " " & order
    End With
End Sub
Private Sub cbosort2_Click()
    With Data1.rst_purchases
        .Sort = .Fields(cbosort2.ItemData(cbosort2.ListIndex)).Name & " " & order2
    End With
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
    Set DataGrid2.DataSource = Nothing
    Data1.conn.Execute "delete from purchases where o_id='" & Data1.rst_purchases.Fields(0) & "'"
    With Data1.rst_purchases
        If .RecordCount = 1 Then
            flag = True
            enablecontrols False
            enablecontrols2 False
        Else
            .MoveNext
            If .EOF Then: .MoveLast
            bm = .Bookmark - 1
        End If
        .Close
        .Open
        Data1.rspurchases.Close
        Data1.rspurchases.Open
        If Not flag Then
            .Bookmark = bm
        End If
    End With
    Set DataGrid1.DataSource = Data1
    Set DataGrid2.DataSource = Data1
    Data1.conn.CommitTrans
    Exit Sub
err:
    Data1.conn.RollbackTrans
    Data1.conn.Cancel
    Set DataGrid1.DataSource = Data1
    Set DataGrid2.DataSource = Data1
    MsgBox err.Description, vbCritical, "Error"
End Sub
Private Sub cmdnavigate_Click(Index As Integer)
    With Data1.rspurchases
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
Private Sub cmdnavigate2_Click(Index As Integer)
    With Data1.rst_purchases
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
Private Sub cmdmod_Click()
    frmpurchase_addmod.modstate = True
    frmpurchase_addmod.Caption = "Modify Purchase"
    frmpurchase_addmod.Show vbModal
End Sub

Public Sub cmdnew_Click()
    frmpurchase_addmod.modstate = False
    frmpurchase_addmod.Show vbModal
End Sub
Public Sub cmdremfilter_Click()
    cmdremfilter.Enabled = False
    Data1.rspurchases.Filter = adFilterNone
    If Data1.rspurchases.RecordCount = 0 Then
        enablecontrols False
    ElseIf cmdsearch.Enabled = False Then
        enablecontrols True
    End If
    cmdremfilter.Enabled = False
End Sub
Public Sub cmdremfilter2_Click()
    cmdremfilter2.Enabled = False
    Data1.rst_purchases.Filter = adFilterNone
    If Data1.rst_purchases.RecordCount = 0 Then
        enablecontrols2 False
    ElseIf cmdDelete.Enabled = False Then
        enablecontrols2 True
    End If
    cmdremfilter2.Enabled = False
End Sub
Private Sub cmdsearch_Click()
    With frmsearch
        Set .srsrs = Data1.rspurchases
        .srsarr = Array("Bill Number", "Bill Date", "Title Code", "Title Name", "Supplier Code", "Supplier Name", "Price", "Quantity", "Amount", "Discount", "Credit Amount", "Net Amount")
        .srsdata = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
        .Show vbModal
    End With
End Sub
Private Sub DataGrid2_DblClick()
    If enable = True Then cmdmod_Click
End Sub
Private Sub DataGrid2_Keydown(KeyCode As Integer, Shift As Integer)
    If enable = True Then
        If KeyCode = vbKeyReturn Then cmdmod_Click
    End If
End Sub
Private Sub Form_Load()
    cmdnew.Enabled = enable
    cmdmod.Enabled = enable
    cmdDelete.Enabled = enable
    cmddelall.Enabled = enable
    cmdfilter.Enabled = enable
    cmdfilter2.Enabled = enable
    cbosort.ListIndex = 0
    cbosort2.ListIndex = 0
    cbofield.ListIndex = 0
    cbofield2.ListIndex = 0
    If Data1.rst_purchases.RecordCount = 0 Then
        enablecontrols False
        enablecontrols2 False
    End If
    updatelabels
    updatelabels2
    vsval = 0
    hsval = 0
    Set rs = Data1.rspurchases
    Set rs2 = Data1.rst_purchases
    purchasesformisopen = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    purchasesformisopen = False
    Data1.rspurchases.Close
    Data1.rst_purchases.Close
    If Data1.rspurchases.Filter <> adFilterNone Then Data1.rspurchases.Filter = adFilterNone
    If Data1.rst_purchases.Filter <> adFilterNone Then Data1.rst_purchases.Filter = adFilterNone
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0: order = "ASC"
        Case 1: order = "DESC"
    End Select
    cbosort_Click
End Sub
Private Sub Option2_Click(Index As Integer)
    Select Case Index
        Case 0: order2 = "ASC"
        Case 1: order2 = "DESC"
    End Select
    cbosort2_Click
End Sub

Sub enablecontrols(val As Boolean)
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
Sub enablecontrols2(val As Boolean)
    cmdmod.Enabled = val
    cmdDelete.Enabled = val
    cmddelall.Enabled = val
    cmdsearch2.Enabled = val
    cbosort2.Enabled = val
    Option2(0).Enabled = val
    Option2(1).Enabled = val
    cmdfilter2.Enabled = val
    DataGrid2.Enabled = val
    cmdnavigate2(0).Enabled = val
    cmdnavigate2(1).Enabled = val
    cmdnavigate2(2).Enabled = val
    cmdnavigate2(3).Enabled = val
End Sub
' similar to the customers form
Private Sub Form_Resize()
    Dim hdisp As Integer
    Dim vdisp As Integer
    Dim h As Integer
    Dim w As Integer
    hdisp = Me.Width - 12060
    vdisp = Me.Height - 8310 + 375 + 600
    If hdisp >= 0 And vdisp >= 0 Then
        VScroll1.Visible = False
        HScroll1.Visible = False
        VScroll1.Value = 0
        HScroll1.Value = 0
        framescroll.Visible = False
        positioncontrols
    ElseIf hdisp >= 0 And vdisp < 0 Then
        VScroll1.Visible = True
        HScroll1.Visible = False
        h = 0
        HScroll1.Value = 0
        HScroll1.Max = -hdisp + VScroll1.Width
        framescroll.Visible = False
    ElseIf hdisp < 0 And vdisp >= 0 Then
        HScroll1.Visible = True
        VScroll1.Visible = False
        w = 0
        VScroll1.Value = 0
        VScroll1.Max = -vdisp + HScroll1.Height
        framescroll.Visible = False
    Else
        VScroll1.Visible = True
        HScroll1.Visible = True
        HScroll1.Max = -hdisp + VScroll1.Width
        VScroll1.Max = -vdisp + HScroll1.Height
        h = HScroll1.Height
        w = VScroll1.Width
        framescroll.Visible = True
    End If
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
Private Sub VScroll1_Change()
    Dim inc As Integer
    inc = VScroll1.Value - vsval
    vsval = VScroll1.Value
    SSTab1.Top = SSTab1.Top - inc
End Sub
Private Sub vScroll1_Scroll()
    VScroll1_Change
End Sub
Private Sub HScroll1_Change()
    Dim inc As Integer
    inc = HScroll1.Value - hsval
    hsval = HScroll1.Value
    SSTab1.Left = SSTab1.Left - inc
End Sub
Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub
Private Sub positioncontrols()
    SSTab1.Left = 120
    SSTab1.Top = 120
End Sub
Private Sub updatelabels()
    ' displays the current record no and the total no of records
    ' in the purchases recordset
    With Data1.rspurchases
        If .RecordCount = 0 Then
            lblcurrec.Caption = "0"
        Else
            lblcurrec.Caption = .AbsolutePosition
        End If
        lblttlrec.Caption = .RecordCount
    End With
End Sub
Private Sub updatelabels2()
    ' displays the current record no and the total no of records
    ' in the total purchases recordset
    With Data1.rst_purchases
        If .RecordCount = 0 Then
            lblcurrec2.Caption = "0"
        Else
            lblcurrec2.Caption = .AbsolutePosition
        End If
        lblttlrec2.Caption = .RecordCount
    End With
End Sub
