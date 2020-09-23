VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00AA820A&
   BorderStyle     =   0  'None
   ClientHeight    =   4245
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmsplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3120
      Top             =   1920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00AA820A&
      BorderStyle     =   0  'None
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Image imgLogo 
         Height          =   2985
         Left            =   120
         Picture         =   "frmsplash.frx":000C
         Stretch         =   -1  'True
         Top             =   555
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackColor       =   &H00AA820A&
         Caption         =   "Created by: Raman Mehta MCA 6th Sem. IGNOU Enrollment No. 050747970"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   4920
         TabIndex        =   1
         Top             =   3360
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00AA820A&
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4890
         TabIndex        =   2
         Top             =   2880
         Width           =   1275
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00AA820A&
         Caption         =   "Book House Inventory Management"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   2760
         TabIndex        =   3
         Top             =   540
         Width           =   3975
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this is a splash screen shown for 4 seconds
' before displaying the login screen
Dim cnt As Byte
Option Explicit
Private Sub Timer1_Timer()
    cnt = cnt + 1
    If cnt = 4 Then
        Timer1.Enabled = False
        Unload Me
    End If
End Sub
