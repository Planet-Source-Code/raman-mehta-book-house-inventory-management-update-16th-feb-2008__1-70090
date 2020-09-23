VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmmain 
   BackColor       =   &H00AA820A&
   Caption         =   "Book House"
   ClientHeight    =   6030
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6540
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "View Book Titles"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "View Customers"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "View Suppliers"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Reports"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit Application"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Log Off"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lock Application"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Calculator"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Notepad"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Calendar"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5655
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1799
            MinWidth        =   1764
            Text            =   "Current User:"
            TextSave        =   "Current User:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2910
            MinWidth        =   2910
            TextSave        =   "2/14/2008"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2910
            MinWidth        =   2910
            TextSave        =   "1:05 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   2910
            MinWidth        =   2910
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2910
            MinWidth        =   2910
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   2910
            MinWidth        =   2910
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":227E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6B3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9846
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":B598
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":B8BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":BBD4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuappl 
      Caption         =   "&Application"
      Begin VB.Menu mnulock 
         Caption         =   "Loc&k"
      End
      Begin VB.Menu mnulogoff 
         Caption         =   "&Log Off"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuusers 
      Caption         =   "&Users"
      Begin VB.Menu mnuviewusers 
         Caption         =   "&View Users"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnucus 
         Caption         =   "&Customers"
      End
      Begin VB.Menu mnusup 
         Caption         =   "&Suppliers"
      End
      Begin VB.Menu mnusub 
         Caption         =   "S&ubjects"
      End
      Begin VB.Menu mnutitles 
         Caption         =   "&Book Titles"
      End
      Begin VB.Menu mnudem 
         Caption         =   "&Demand"
      End
      Begin VB.Menu mnupur 
         Caption         =   "&Purchases"
      End
      Begin VB.Menu mnusales 
         Caption         =   "Sa&les"
      End
   End
   Begin VB.Menu mnunew 
      Caption         =   "&New"
      Begin VB.Menu mnunewcus 
         Caption         =   "&Customer"
      End
      Begin VB.Menu mnunewsup 
         Caption         =   "&Supplier"
      End
      Begin VB.Menu mnunewtitle 
         Caption         =   "&Title"
      End
      Begin VB.Menu mnunewpur 
         Caption         =   "&Purchase"
      End
      Begin VB.Menu mnuewsale 
         Caption         =   "Sa&le"
      End
   End
   Begin VB.Menu mnucredits 
      Caption         =   "&Credits"
      Begin VB.Menu mnucpay 
         Caption         =   "Credit &Payments"
      End
      Begin VB.Menu mnucrec 
         Caption         =   "Credit &Receipts"
      End
   End
   Begin VB.Menu mnuutil 
      Caption         =   "U&tilities"
      Begin VB.Menu mnucalc 
         Caption         =   "&Calculator"
      End
      Begin VB.Menu mnunote 
         Caption         =   "&Notepad"
      End
      Begin VB.Menu mnucal 
         Caption         =   "Calen&dar"
      End
   End
   Begin VB.Menu mnurepts 
      Caption         =   "&Reports"
      Begin VB.Menu mnureptlist 
         Caption         =   "&Show List"
      End
      Begin VB.Menu mnureptcus 
         Caption         =   "All &Customers"
      End
      Begin VB.Menu mnureptsup 
         Caption         =   "All &Suppliers"
      End
      Begin VB.Menu mnurepttitles 
         Caption         =   "All &Titles"
      End
      Begin VB.Menu mnureptsub 
         Caption         =   "&All Subjects"
      End
   End
   Begin VB.Menu mnuwind 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnucas 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuhoriz 
         Caption         =   "Tile Windows &Horizontally"
      End
      Begin VB.Menu mnuvert 
         Caption         =   "Tile Windows &Vertically"
      End
      Begin VB.Menu mnuarr 
         Caption         =   "&Arrange Icons"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
    Me.BackColor = RGB(10, 130, 170)
    ' display the current user and his role on the status bar
    StatusBar1.Panels(1).Text = "Current User:  " & user & "(" & role & ")"
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure you want to quit the Application?", vbQuestion + vbYesNo) = vbNo Then Cancel = True
End Sub

Private Sub mnuarr_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnucal_Click()
    frmcalendar.Show
    frmcalendar.SetFocus
End Sub

Private Sub mnucas_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnucpay_Click()
    Me.Caption = "Loading..."
    Load frmcrpayments
    Me.Caption = "Book House"
    frmcrpayments.Show vbModal
End Sub

Private Sub mnucrec_Click()
    Me.Caption = "Loading..."
    Load frmcrreceipts
    Me.Caption = "Book House"
    frmcrreceipts.Show vbModal
End Sub

Private Sub mnudem_Click()
    Me.Caption = "Loading..."
    Load frmdemand
    Me.Caption = "Book House"
    frmdemand.Show vbModal
End Sub

Private Sub mnuewsale_Click()
    frmsales.cmdnew_Click
End Sub

Private Sub mnuexit_Click()
    Unload Me
End Sub

Private Sub mnuhoriz_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnulock_Click()
    frmunlock.Show vbModal
End Sub

Private Sub mnulogoff_Click()
    If MsgBox("Are you sure you want to Log Off?", vbQuestion + vbYesNo) = vbYes Then
        For i = 1 To 10
            frmmain.Toolbar1.Buttons(i).Enabled = False
        Next i
        StatusBar1.Panels(1).Text = "Current User:"
        For Each X In Forms
            If X.Name <> "frmmain" Then
                Unload X
            End If
        Next
        frmlogin.Show vbModal
    End If
End Sub

Private Sub mnunewpur_Click()
    frmpurchases.cmdnew_Click
End Sub

Private Sub mnunewtitle_Click()
    frmtitles.cmdnewmod_Click 0
End Sub

Private Sub mnunote_Click()
    On Error GoTo err:
    Shell "notepad.exe", vbNormalFocus
    Exit Sub
err:
    MsgBox "Notepad utility cannot be run", vbExclamation, "Error"
End Sub

Private Sub mnucalc_Click()
    On Error GoTo err:
    Shell "calc.exe", vbNormalFocus
    Exit Sub
err:
    MsgBox "Calculator utility cannot be run", vbExclamation, "Error"
End Sub

Private Sub mnucus_Click()
    Me.Caption = "Loading..."
    With frmcustomers
        .Show
        .SetFocus
    End With
    Me.Caption = "Book House"
End Sub

Private Sub mnunewcus_Click()
    frmcustomers.cmdnewmod_Click 0
End Sub

Private Sub mnunewsup_Click()
    frmsuppliers.cmdnewmod_Click 0
End Sub

Private Sub mnupur_Click()
    Me.Caption = "Loading..."
    With frmpurchases
        .Show
        .SetFocus
    End With
    Me.Caption = "Book House"
End Sub


Private Sub mnurepsup_Click()
    rptsuppliers.Show vbModal
End Sub

Private Sub mnureptcus_Click()
    rptcustomers.Show vbModal
End Sub

Private Sub mnureptlist_Click()
    frmrepts.Show vbModal
End Sub

Private Sub mnureptsub_Click()
    rptsub.Show vbModal
End Sub

Private Sub mnureptsup_Click()
    rptsuppliers.Show vbModal
End Sub

Private Sub mnurepttitles_Click()
    rpttitles.Show vbModal
End Sub

Private Sub mnusales_Click()
    Me.Caption = "Loading..."
    With frmsales
        .Show
        .SetFocus
    End With
    Me.Caption = "Book House"
End Sub

Private Sub mnusub_Click()
    Me.Caption = "Loading..."
    Load frmsubjects
    Me.Caption = "Book House"
    frmsubjects.Show vbModal
End Sub

Private Sub mnusup_Click()
    Me.Caption = "Loading..."
    With frmsuppliers
        .Show
        .SetFocus
    End With
    Me.Caption = "Book House"
End Sub

Private Sub mnutitles_Click()
    Me.Caption = "Loading..."
    With frmtitles
        .Show
        .SetFocus
    End With
    Me.Caption = "Book House"
End Sub
Private Sub mnuvert_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuviewusers_Click()
    Me.Caption = "Loading..."
    With frmusers
        .Show
        .SetFocus
    End With
    Me.Caption = "Book House"
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnutitles_Click
        Case 2: mnucus_Click
        Case 3: mnusup_Click
        Case 4: mnureptlist_Click
        Case 5: mnuexit_Click
        Case 6: mnulogoff_Click
        Case 7: mnulock_Click
        Case 8: mnucalc_Click
        Case 9: mnunote_Click
        Case 10: mnucal_Click
    End Select
End Sub
