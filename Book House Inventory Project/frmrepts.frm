VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmrepts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4845
   Icon            =   "frmrepts.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4845
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Customers"
      TabPicture(0)   =   "frmrepts.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdcus"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Suppliers"
      TabPicture(1)   =   "frmrepts.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "cmdsupp"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Titles"
      TabPicture(2)   =   "frmrepts.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "cmdtitles"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Purchases"
      TabPicture(3)   =   "frmrepts.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(1)=   "cbosup"
      Tab(3).Control(2)=   "cmdpurchases"
      Tab(3).Control(3)=   "dtp1"
      Tab(3).Control(4)=   "dtp2"
      Tab(3).Control(5)=   "Label5"
      Tab(3).Control(6)=   "Label2"
      Tab(3).Control(7)=   "Label1"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "Sales"
      TabPicture(4)   =   "frmrepts.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label3"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label4"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label6"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "dtp4"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "dtp3"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "cbocus"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "cmdsales"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Frame5"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "Subjects"
      TabPicture(5)   =   "frmrepts.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdsub"
      Tab(5).Control(1)=   "Frame6"
      Tab(5).ControlCount=   2
      Begin VB.CommandButton cmdsub 
         Caption         =   "View Re&port"
         Height          =   495
         Left            =   -73200
         TabIndex        =   21
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   34
         Top             =   600
         Width           =   1935
         Begin VB.OptionButton Option10 
            Caption         =   "All Subjects"
            Height          =   495
            Left            =   360
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   615
         Left            =   -74760
         TabIndex        =   33
         Top             =   600
         Width           =   3495
         Begin VB.OptionButton Option9 
            Caption         =   "Credit Sales"
            Height          =   255
            Left            =   1920
            TabIndex        =   15
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Option8 
            Caption         =   "All Sales"
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   615
         Left            =   -74760
         TabIndex        =   32
         Top             =   600
         Width           =   3495
         Begin VB.OptionButton Option6 
            Caption         =   "All Purchases"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Credit Purchases"
            Height          =   255
            Left            =   1920
            TabIndex        =   9
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   855
         Left            =   -74760
         TabIndex        =   31
         Top             =   600
         Width           =   1935
         Begin VB.OptionButton Option5 
            Caption         =   "All Titles"
            Height          =   495
            Left            =   360
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1215
         Left            =   -74760
         TabIndex        =   30
         Top             =   600
         Width           =   4335
         Begin VB.OptionButton Option4 
            Caption         =   "Suppliers having Credit Balance > 0"
            Height          =   495
            Left            =   1920
            TabIndex        =   4
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton Option3 
            Caption         =   "All Suppliers"
            Height          =   495
            Left            =   360
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1935
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   4335
         Begin VB.OptionButton Option1 
            Caption         =   "All Customers"
            Height          =   495
            Left            =   360
            TabIndex        =   0
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Customers having Credit Balance > 0"
            Height          =   495
            Left            =   1920
            TabIndex        =   1
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdsales 
         Caption         =   "View &Report"
         Height          =   495
         Left            =   -73200
         TabIndex        =   19
         Top             =   3000
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo cbosup 
         Bindings        =   "frmrepts.frx":0972
         Height          =   315
         Left            =   -74520
         TabIndex        =   10
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "S_CODE"
         Text            =   ""
         Object.DataMember      =   "suppliers"
      End
      Begin VB.CommandButton cmdpurchases 
         Caption         =   "Vie&w Report"
         Height          =   495
         Left            =   -73200
         TabIndex        =   13
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdtitles 
         Caption         =   "Vi&ew Report"
         Height          =   495
         Left            =   -73200
         TabIndex        =   7
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdcus 
         Caption         =   "&View Report"
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdsupp 
         Caption         =   "V&iew Report"
         Height          =   495
         Left            =   -73200
         TabIndex        =   5
         Top             =   3000
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtp1 
         Height          =   300
         Left            =   -74520
         TabIndex        =   11
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   64159745
         CurrentDate     =   39355
      End
      Begin MSComCtl2.DTPicker dtp2 
         Height          =   300
         Left            =   -72000
         TabIndex        =   12
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   64159745
         CurrentDate     =   39355
      End
      Begin MSDataListLib.DataCombo cbocus 
         Bindings        =   "frmrepts.frx":0986
         Height          =   315
         Left            =   -74520
         TabIndex        =   16
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "C_CODE"
         Text            =   ""
         Object.DataMember      =   "customers"
      End
      Begin MSComCtl2.DTPicker dtp3 
         Height          =   300
         Left            =   -74520
         TabIndex        =   17
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   64159745
         CurrentDate     =   39355
      End
      Begin MSComCtl2.DTPicker dtp4 
         Height          =   300
         Left            =   -72000
         TabIndex        =   18
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   64159745
         CurrentDate     =   39355
      End
      Begin VB.Label Label6 
         Caption         =   "Customer Code:"
         Height          =   255
         Left            =   -74520
         TabIndex        =   28
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Supplier Code:"
         Height          =   255
         Left            =   -74520
         TabIndex        =   27
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Between Date"
         Height          =   255
         Left            =   -74520
         TabIndex        =   26
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "and"
         Height          =   255
         Left            =   -72600
         TabIndex        =   25
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "and"
         Height          =   255
         Left            =   -72600
         TabIndex        =   24
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Between Date"
         Height          =   255
         Left            =   -74520
         TabIndex        =   23
         Top             =   2160
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmrepts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this form displays various options for generating reports
Private Sub cmdcus_Click()
    If Option1.Value = True Then
        Data1.rsrptcustomers.Filter = adFilterNone
    Else
        Data1.rsrptcustomers.Filter = "credit>0"
    End If
    Set rptcustomers.DataSource = Data1
    rptcustomers.Show vbModal
End Sub

Private Sub cmdpurchases_Click()
    ' dynamically generate the purchase invoice depending upon the
    ' dates and the supplier code selected by the user
    If Option6.Value = True Then
        If dtp1.Value < dtp2.Value Then
            Data1.rsrptpurchases.Filter = "s_code='" & cbosup.Text & "' and " & "p_date >=#" & dtp1.Value & "# and p_date <=#" & dtp2.Value & "#"
        Else
            Data1.rsrptpurchases.Filter = "s_code='" & cbosup.Text & "' and " & "p_date >=#" & dtp2.Value & "# and p_date <=#" & dtp1.Value & "#"
        End If
    Else
        If dtp1.Value < dtp2.Value Then
            Data1.rsrptpurchases.Filter = "s_code='" & cbosup.Text & "' and " & "p_date >=#" & dtp1.Value & "# and p_date <=#" & dtp2.Value & "# and c_total>0"
        Else
            Data1.rsrptpurchases.Filter = "s_code='" & cbosup.Text & "' and " & "p_date >=#" & dtp2.Value & "# and p_date <=#" & dtp1.Value & "# and c_total>0"
        End If
    End If
    Set rptpurchases.DataSource = Data1
    rptpurchases.Show vbModal
End Sub

Private Sub cmdsales_Click()
    ' dynamically generate the sale invoice depending upon the
    ' dates and the supplier code selected by the user
    If Option8.Value = True Then
        If dtp3.Value < dtp4.Value Then
            Data1.rsrptsales.Filter = "c_code='" & cbocus.Text & "' and " & "s_date >=#" & dtp3.Value & "# and s_date <=#" & dtp4.Value & "#"
        Else
            Data1.rsrptsales.Filter = "c_code='" & cbocus.Text & "' and " & "s_date >=#" & dtp4.Value & "# and s_date <=#" & dtp3.Value & "#"
        End If
    Else
        If dtp3.Value < dtp4.Value Then
            Data1.rsrptsales.Filter = "c_code='" & cbocus.Text & "' and " & "s_date >=#" & dtp3.Value & "# and s_date <=#" & dtp4.Value & "# and c_total>0"
        Else
            Data1.rsrptsales.Filter = "c_code='" & cbocus.Text & "' and " & "s_date >=#" & dtp4.Value & "# and s_date <=#" & dtp3.Value & "# and c_total>0"
        End If
    End If
    Set rptsales.DataSource = Data1
    rptsales.Show vbModal
End Sub

Private Sub cmdsub_Click()
    rptsub.Show vbModal
End Sub

Private Sub cmdsupp_Click()
    If Option3.Value = True Then
        Data1.rsrptsuppliers.Filter = adFilterNone
    Else
        Data1.rsrptsuppliers.Filter = "credit>0"
    End If
    Set rptsuppliers.DataSource = Data1
    rptsuppliers.Show vbModal
End Sub

Private Sub cmdtitles_Click()
    rpttitles.Show vbModal
End Sub

Private Sub Form_Load()
    dtp1.Value = Date
    dtp2.Value = Date
    dtp3.Value = Date
    dtp4.Value = Date
End Sub
