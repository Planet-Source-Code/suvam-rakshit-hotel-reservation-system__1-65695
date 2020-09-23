VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form splashScreen 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3210
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "splashScreen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar hotelProgressBar 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   2850
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7080
      Begin VB.Timer progressTimer 
         Interval        =   1000
         Left            =   6360
         Top             =   1680
      End
      Begin VB.Timer loadTimer 
         Interval        =   1000
         Left            =   6360
         Top             =   2160
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "splashScreen.frx":000C
         Stretch         =   -1  'True
         Top             =   315
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright : Product is copyrighted in the year 2006"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   2280
         Width           =   3855
      End
      Begin VB.Label lblCompany 
         Caption         =   "Company : Inditech Solutions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   1800
         Width           =   3975
      End
      Begin VB.Label lblWarning 
         Caption         =   " Warning : Use of pirated copy is illegal."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2400
         TabIndex        =   1
         Top             =   2520
         Width           =   3015
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version : 1.0"
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
         Left            =   4440
         TabIndex        =   4
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Platform : Windows"
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
         Left            =   4440
         TabIndex        =   5
         Top             =   840
         Width           =   2235
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Hotel Reservation System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2355
         TabIndex        =   6
         Top             =   225
         Width           =   4410
      End
   End
End
Attribute VB_Name = "splashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub Form_Load()
    splashScreen.Show
End Sub

Private Sub progressTimer_Timer()
    i = Rnd() * 20
    If hotelProgressBar.Value < 100 Then
        If hotelProgressBar.Value + i < 100 Then
            hotelProgressBar.Value = hotelProgressBar.Value + i
        Else
            hotelProgressBar.Value = 100
        End If
    Else
        HRS.Show
    End If
End Sub
