VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form HRS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hotel Reservation System"
   ClientHeight    =   9735
   ClientLeft      =   2835
   ClientTop       =   870
   ClientWidth     =   9750
   Icon            =   "HRSmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   9750
   Begin TabDlg.SSTab tab1 
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   16325
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "&Booking"
      TabPicture(0)   =   "HRSmain.frx":212A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "addButton"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cancelButton"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "bookingTimer"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Search / Cancellation"
      TabPicture(1)   =   "HRSmain.frx":2146
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "updateAmount"
      Tab(1).Control(1)=   "recordList"
      Tab(1).Control(2)=   "cancelBookingButton"
      Tab(1).Control(3)=   "nextSearchButton"
      Tab(1).Control(4)=   "customerInformationFrame"
      Tab(1).Control(5)=   "searchCustomerFrame"
      Tab(1).Control(6)=   "searchButton"
      Tab(1).Control(7)=   "cancelSearchButton"
      Tab(1).Control(8)=   "srecordLabel"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Room &Information"
      TabPicture(2)   =   "HRSmain.frx":2162
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frm_ViewOpt"
      Tab(2).Control(1)=   "Frame24"
      Tab(2).Control(2)=   "Frame25"
      Tab(2).Control(3)=   "RoomGrid"
      Tab(2).ControlCount=   4
      Begin VB.CommandButton updateAmount 
         Caption         =   "&Update Amount"
         Height          =   375
         Left            =   -70200
         TabIndex        =   169
         Top             =   8760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Timer bookingTimer 
         Interval        =   100
         Left            =   240
         Top             =   8760
      End
      Begin VB.ComboBox recordList 
         Height          =   315
         Left            =   -73440
         Style           =   2  'Dropdown List
         TabIndex        =   168
         Top             =   8760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame Frm_ViewOpt 
         Caption         =   "Room View"
         Height          =   1815
         Left            =   -74640
         TabIndex        =   160
         Top             =   6060
         Width           =   2055
         Begin VB.OptionButton Opt_InUse 
            Caption         =   "Rooms In Use"
            Height          =   255
            Left            =   360
            MaskColor       =   &H8000000F&
            TabIndex        =   163
            Top             =   1365
            Width           =   1455
         End
         Begin VB.OptionButton Opt_Free 
            Caption         =   "Free Rooms"
            Height          =   255
            Left            =   360
            MaskColor       =   &H8000000F&
            TabIndex        =   162
            Top             =   1005
            Width           =   1335
         End
         Begin VB.OptionButton Opt_All 
            Caption         =   "All Rooms"
            Height          =   255
            Left            =   360
            MaskColor       =   &H8000000F&
            TabIndex        =   161
            Top             =   645
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.Label lbl_optheader 
            Caption         =   "Only Show:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   164
            Top             =   345
            Width           =   1095
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "Roomwise Customer Detail"
         Height          =   2895
         Left            =   -72480
         TabIndex        =   147
         Top             =   6060
         Width           =   6615
         Begin VB.Label noCustomerLabel 
            Alignment       =   2  'Center
            Caption         =   "Room Empty"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1920
            TabIndex        =   148
            Top             =   1200
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.Label cNameLabel 
            Caption         =   "Name : "
            Height          =   255
            Left            =   120
            TabIndex        =   159
            Top             =   360
            Width           =   6255
         End
         Begin VB.Label cAddressLabel 
            Caption         =   "Address : "
            Height          =   255
            Left            =   120
            TabIndex        =   158
            Top             =   720
            Width           =   4455
         End
         Begin VB.Label cCountryLabel 
            Caption         =   "Country : "
            Height          =   255
            Left            =   120
            TabIndex        =   157
            Top             =   1080
            Width           =   3975
         End
         Begin VB.Label cStateLabel 
            Caption         =   "State : "
            Height          =   255
            Left            =   120
            TabIndex        =   156
            Top             =   1440
            Width           =   3495
         End
         Begin VB.Label cCityLabel 
            Caption         =   "City : "
            Height          =   255
            Left            =   120
            TabIndex        =   155
            Top             =   1800
            Width           =   3735
         End
         Begin VB.Label cMobileLabel 
            Caption         =   "Contact No. (Mobile) : "
            Height          =   255
            Left            =   120
            TabIndex        =   154
            Top             =   2520
            Width           =   3615
         End
         Begin VB.Label cArrivalTMLabel 
            Caption         =   "Arrival Time : "
            Height          =   255
            Left            =   3960
            TabIndex        =   153
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Label cArrivalDTLabel 
            Caption         =   "Arrival Date : "
            Height          =   255
            Left            =   3960
            TabIndex        =   152
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label cPinLabel 
            Caption         =   "Pin : "
            Height          =   255
            Left            =   120
            TabIndex        =   151
            Top             =   2160
            Width           =   2535
         End
         Begin VB.Label cDeptTMLabel 
            Caption         =   "Departure Time : "
            Height          =   255
            Left            =   3960
            TabIndex        =   150
            Top             =   2400
            Width           =   2535
         End
         Begin VB.Label cDeptDTLabel 
            Caption         =   "Departure Date : "
            Height          =   255
            Left            =   3960
            TabIndex        =   149
            Top             =   2040
            Width           =   2415
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "Room Status"
         Height          =   975
         Left            =   -74640
         TabIndex        =   145
         Top             =   7980
         Width           =   2055
         Begin VB.Label roomStatusLabel 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   146
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.CommandButton cancelBookingButton 
         Caption         =   "Cancel &Booking"
         Height          =   375
         Left            =   -68640
         TabIndex        =   138
         Top             =   8760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton nextSearchButton 
         Caption         =   "&Next"
         Height          =   375
         Left            =   -67080
         TabIndex        =   136
         Top             =   8760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame customerInformationFrame 
         Caption         =   "Search Result / Cancellation Information"
         Height          =   8175
         Left            =   -74760
         TabIndex        =   71
         Top             =   480
         Visible         =   0   'False
         Width           =   9015
         Begin VB.Frame Frame21 
            Caption         =   "Arrival Details"
            Height          =   2295
            Left            =   240
            TabIndex        =   127
            Top             =   4200
            Width           =   5055
            Begin VB.TextBox sCustVehicleNoText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   131
               Top             =   840
               Width           =   1575
            End
            Begin VB.TextBox sCustArrivalText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   130
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox sCustArrDTText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   129
               Top             =   1320
               Width           =   1575
            End
            Begin VB.TextBox sCustArrTMtext 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   128
               Top             =   1800
               Width           =   1575
            End
            Begin VB.Label Label34 
               Caption         =   "Arrival by :"
               Height          =   255
               Left            =   240
               TabIndex        =   135
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label35 
               Caption         =   "Vehicle No :"
               Height          =   255
               Left            =   240
               TabIndex        =   134
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label36 
               Caption         =   "Arrival Date : (Booking Date)"
               Height          =   495
               Left            =   240
               TabIndex        =   133
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label Label37 
               Caption         =   "Arrival Time :"
               Height          =   255
               Left            =   240
               TabIndex        =   132
               Top             =   1920
               Width           =   1095
            End
         End
         Begin VB.Frame Frame22 
            Caption         =   "Departure Details"
            Height          =   1335
            Left            =   240
            TabIndex        =   122
            Top             =   6600
            Width           =   5055
            Begin VB.TextBox sCustDeptDTText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   124
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox sCustDeptTMText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   123
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label Label38 
               Caption         =   "Departure Time :"
               Height          =   255
               Left            =   240
               TabIndex        =   126
               Top             =   840
               Width           =   1455
            End
            Begin VB.Label Label39 
               Caption         =   "Departure Date :"
               Height          =   255
               Left            =   240
               TabIndex        =   125
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.Frame Frame23 
            Caption         =   "Payment Detail"
            Height          =   2295
            Left            =   5520
            TabIndex        =   112
            Top             =   5640
            Width           =   3255
            Begin VB.TextBox sCustAmountText 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               TabIndex        =   119
               Top             =   360
               Width           =   1215
            End
            Begin VB.Frame Frame26 
               Caption         =   "Payment Type"
               Height          =   1215
               Left            =   240
               TabIndex        =   113
               Top             =   960
               Width           =   2775
               Begin VB.TextBox sCustTypeNoText 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   1440
                  TabIndex        =   117
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.OptionButton sCustCashOption 
                  Caption         =   "Cash"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   116
                  Top             =   360
                  Width           =   735
               End
               Begin VB.OptionButton sCustDDOption 
                  Caption         =   "DD"
                  Height          =   255
                  Left            =   840
                  TabIndex        =   115
                  Top             =   360
                  Width           =   615
               End
               Begin VB.OptionButton sCustCreditOption 
                  Caption         =   "Credit card"
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   114
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   1215
               End
               Begin VB.Label sCustTypeNoLabel 
                  Caption         =   "Credit card no :"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   118
                  Top             =   720
                  Width           =   1095
               End
            End
            Begin VB.Label Label41 
               Caption         =   "Amount Payable :"
               Height          =   255
               Left            =   240
               TabIndex        =   121
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label sIncludingFoodLabel 
               Caption         =   "(including food)"
               Height          =   255
               Left            =   1800
               TabIndex        =   120
               Top             =   720
               Visible         =   0   'False
               Width           =   1095
            End
         End
         Begin VB.Frame Frame27 
            Caption         =   "Companion Detail"
            Height          =   1335
            Left            =   5520
            TabIndex        =   107
            Top             =   4200
            Width           =   3255
            Begin VB.OptionButton sCustSpouseOption 
               Caption         =   "Spouse"
               Height          =   255
               Left            =   240
               TabIndex        =   110
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton sCustWSpouseOption 
               Caption         =   "Without spouse"
               Height          =   255
               Left            =   1680
               TabIndex        =   109
               Top             =   360
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.TextBox sCustRelationText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               TabIndex        =   108
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label sRelationLabel 
               Caption         =   "Relationship :"
               Enabled         =   0   'False
               Height          =   255
               Left            =   240
               TabIndex        =   111
               Top             =   840
               Width           =   1095
            End
         End
         Begin VB.Frame Frame28 
            Caption         =   "Food Details"
            Height          =   1335
            Left            =   5520
            TabIndex        =   102
            Top             =   2640
            Width           =   3255
            Begin VB.OptionButton sCustOutsideOption 
               Caption         =   "Outside"
               Height          =   255
               Left            =   1560
               TabIndex        =   105
               Top             =   360
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton sCustHotelOption 
               Caption         =   "Hotel"
               Height          =   255
               Left            =   240
               TabIndex        =   104
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox sCustFoodChoiceText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               TabIndex        =   103
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label sFoodChoiceLabel 
               Caption         =   "Food-Choice :"
               Enabled         =   0   'False
               Height          =   255
               Left            =   240
               TabIndex        =   106
               Top             =   840
               Width           =   1095
            End
         End
         Begin VB.Frame Frame31 
            Caption         =   "Room Details"
            Height          =   2055
            Left            =   5520
            TabIndex        =   94
            Top             =   360
            Width           =   3255
            Begin VB.TextBox sCustRoomNoText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               TabIndex        =   99
               Top             =   360
               Width           =   735
            End
            Begin VB.Frame Frame32 
               Caption         =   "AC Facility :"
               Height          =   615
               Left            =   240
               TabIndex        =   96
               Top             =   1200
               Width           =   2775
               Begin VB.OptionButton sCustACOption 
                  Caption         =   "AC"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   98
                  Top             =   240
                  Width           =   735
               End
               Begin VB.OptionButton sCustNACOption 
                  Caption         =   "Non-AC"
                  Height          =   255
                  Left            =   1560
                  TabIndex        =   97
                  Top             =   240
                  Width           =   975
               End
            End
            Begin VB.TextBox sCustRoomTypeText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               TabIndex        =   95
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label Label48 
               Caption         =   "Room No :"
               Height          =   255
               Left            =   240
               TabIndex        =   101
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label49 
               Caption         =   "Room Type :"
               Height          =   255
               Left            =   240
               TabIndex        =   100
               Top             =   840
               Width           =   975
            End
         End
         Begin VB.Frame Frame33 
            Caption         =   "Customer Personal"
            Height          =   3615
            Left            =   240
            TabIndex        =   72
            Top             =   360
            Width           =   5055
            Begin VB.TextBox sCustPinText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3600
               TabIndex        =   83
               Top             =   2160
               Width           =   1095
            End
            Begin VB.TextBox sCustIDText 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   1800
               TabIndex        =   82
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox sCustNameText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   81
               Top             =   720
               Width           =   2895
            End
            Begin VB.TextBox sCustAddressText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   80
               Top             =   1080
               Width           =   2895
            End
            Begin VB.TextBox sCustStateText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   79
               Top             =   1800
               Width           =   2895
            End
            Begin VB.TextBox sCustCityText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   78
               Top             =   2160
               Width           =   1335
            End
            Begin VB.TextBox sCustResText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   77
               Top             =   2520
               Width           =   1575
            End
            Begin VB.TextBox sCustMobileText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   76
               Top             =   2880
               Width           =   1575
            End
            Begin VB.TextBox sCustEmailText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   75
               Top             =   3240
               Width           =   2895
            End
            Begin VB.Frame Frame34 
               Caption         =   "Frame3"
               Height          =   15
               Left            =   4320
               TabIndex        =   74
               Top             =   4920
               Width           =   135
            End
            Begin VB.TextBox sCustCountryText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   73
               Top             =   1440
               Width           =   1815
            End
            Begin VB.Label Label60 
               Caption         =   "&&"
               Height          =   255
               Left            =   3360
               TabIndex        =   93
               Top             =   2160
               Width           =   135
            End
            Begin VB.Label Label61 
               Caption         =   "Customer ID :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   92
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label62 
               Caption         =   "Customer Name :"
               Height          =   255
               Left            =   240
               TabIndex        =   91
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Label63 
               Caption         =   "Address :"
               Height          =   255
               Left            =   240
               TabIndex        =   90
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label64 
               Caption         =   "Country :"
               Height          =   255
               Left            =   240
               TabIndex        =   89
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label Label65 
               Caption         =   "State :"
               Height          =   255
               Left            =   240
               TabIndex        =   88
               Top             =   1800
               Width           =   1695
            End
            Begin VB.Label Label69 
               Caption         =   "City && Pin :"
               Height          =   255
               Left            =   240
               TabIndex        =   87
               Top             =   2160
               Width           =   1455
            End
            Begin VB.Label Label70 
               Caption         =   "Contact No (Res) :"
               Height          =   255
               Left            =   240
               TabIndex        =   86
               Top             =   2520
               Width           =   1935
            End
            Begin VB.Label Label71 
               Caption         =   "Contact No (Mobile) :"
               Height          =   255
               Left            =   240
               TabIndex        =   85
               Top             =   2880
               Width           =   1575
            End
            Begin VB.Label Label72 
               Caption         =   "E-mail :"
               Height          =   255
               Left            =   240
               TabIndex        =   84
               Top             =   3240
               Width           =   1215
            End
         End
      End
      Begin VB.CommandButton cancelButton 
         Cancel          =   -1  'True
         Caption         =   "Cancel &Booking"
         Height          =   375
         Left            =   7920
         TabIndex        =   58
         Top             =   8760
         Width           =   1335
      End
      Begin VB.CommandButton addButton 
         Caption         =   "Add C&ustomer"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Left            =   6360
         TabIndex        =   57
         Top             =   8760
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Booking Information"
         Height          =   8175
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   9015
         Begin VB.Frame Frame7 
            Caption         =   "Arrival Details"
            Height          =   2295
            Left            =   240
            TabIndex        =   63
            Top             =   4200
            Width           =   5055
            Begin VB.Timer arrivalTimer 
               Interval        =   100
               Left            =   3960
               Top             =   1800
            End
            Begin VB.ComboBox custArrCombo 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "HRSmain.frx":217E
               Left            =   1800
               List            =   "HRSmain.frx":2180
               Sorted          =   -1  'True
               TabIndex        =   12
               Text            =   "Select"
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox custVehicleText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   13
               Top             =   840
               Width           =   1575
            End
            Begin MSComCtl2.DTPicker custArrTMPicker 
               Height          =   375
               Left            =   1800
               TabIndex        =   15
               Top             =   1800
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               Format          =   20381698
               CurrentDate     =   0.5
            End
            Begin MSComCtl2.DTPicker custArrDTPicker 
               Height          =   375
               Left            =   1800
               TabIndex        =   14
               Top             =   1320
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "  dd - mm - yyyy"
               Format          =   20381697
               CurrentDate     =   38806
            End
            Begin VB.Label Label16 
               Caption         =   "Arrival by :"
               Height          =   255
               Left            =   240
               TabIndex        =   67
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label15 
               Caption         =   "Vehicle No :"
               Height          =   255
               Left            =   240
               TabIndex        =   66
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label14 
               Caption         =   "Arrival Date : (Booking Date)"
               Height          =   495
               Left            =   240
               TabIndex        =   65
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label Label13 
               Caption         =   "Arrival Time :"
               Height          =   255
               Left            =   240
               TabIndex        =   64
               Top             =   1920
               Width           =   1095
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Departure Details"
            Height          =   1335
            Left            =   240
            TabIndex        =   60
            Top             =   6600
            Width           =   5055
            Begin MSComCtl2.DTPicker custDeptDTPicker 
               Height          =   375
               Left            =   1800
               TabIndex        =   16
               Top             =   360
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "  dd - mm - yyyy"
               Format          =   20381697
               CurrentDate     =   38806
            End
            Begin MSComCtl2.DTPicker custDeptTMPicker 
               Height          =   375
               Left            =   1800
               TabIndex        =   17
               Top             =   840
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               Format          =   20381698
               CurrentDate     =   0.5
            End
            Begin VB.Label Label20 
               Caption         =   "Departure Time :"
               Height          =   255
               Left            =   240
               TabIndex        =   62
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label Label19 
               Caption         =   "Departure Date :"
               Height          =   255
               Left            =   240
               TabIndex        =   61
               Top             =   480
               Width           =   1455
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Payment Detail"
            Height          =   2295
            Left            =   5520
            TabIndex        =   52
            Top             =   5640
            Width           =   3255
            Begin VB.TextBox amountText 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               TabIndex        =   28
               Top             =   360
               Width           =   1215
            End
            Begin VB.Frame Frame11 
               Caption         =   "Payment Type"
               Height          =   1215
               Left            =   240
               TabIndex        =   53
               Top             =   960
               Width           =   2775
               Begin VB.TextBox typeNoText 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   1440
                  TabIndex        =   32
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.OptionButton cashOption 
                  Caption         =   "Cash"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   29
                  Top             =   360
                  Width           =   735
               End
               Begin VB.OptionButton DDOption 
                  Caption         =   "DD"
                  Height          =   255
                  Left            =   840
                  TabIndex        =   30
                  Top             =   360
                  Width           =   615
               End
               Begin VB.OptionButton creditCardOption 
                  Caption         =   "Credit card"
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   31
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   1215
               End
               Begin VB.Label typeNoLabel 
                  Caption         =   "Credit card no :"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   54
                  Top             =   720
                  Width           =   1335
               End
            End
            Begin VB.Label amountLabel 
               Caption         =   "Amount Payable :"
               Height          =   255
               Left            =   240
               TabIndex        =   56
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label includingFoodLabel 
               Caption         =   "(including food)"
               Height          =   255
               Left            =   1800
               TabIndex        =   55
               Top             =   720
               Visible         =   0   'False
               Width           =   1095
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Companion Detail"
            Height          =   1335
            Left            =   5520
            TabIndex        =   50
            Top             =   4200
            Width           =   3255
            Begin VB.ComboBox relationCombo 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "HRSmain.frx":2182
               Left            =   1560
               List            =   "HRSmain.frx":2184
               Sorted          =   -1  'True
               TabIndex        =   27
               Text            =   "Select"
               Top             =   840
               Width           =   1455
            End
            Begin VB.OptionButton spouseOption 
               Caption         =   "Spouse"
               Height          =   255
               Left            =   240
               TabIndex        =   25
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton withoutSpouseOption 
               Caption         =   "Without spouse"
               Height          =   255
               Left            =   1680
               TabIndex        =   26
               Top             =   360
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.Label relationshipLabel 
               Caption         =   "Relationship :"
               Enabled         =   0   'False
               Height          =   255
               Left            =   240
               TabIndex        =   51
               Top             =   840
               Width           =   1095
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Food Details"
            Height          =   1335
            Left            =   5520
            TabIndex        =   48
            Top             =   2640
            Width           =   3255
            Begin VB.ComboBox foodChoiceCombo 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "HRSmain.frx":2186
               Left            =   1560
               List            =   "HRSmain.frx":2188
               Sorted          =   -1  'True
               TabIndex        =   24
               Text            =   "Select"
               Top             =   840
               Width           =   1455
            End
            Begin VB.OptionButton outsideOption 
               Caption         =   "Outside"
               Height          =   255
               Left            =   1560
               TabIndex        =   23
               Top             =   360
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton hotelOption 
               Caption         =   "Hotel"
               Height          =   255
               Left            =   240
               TabIndex        =   22
               Top             =   360
               Width           =   975
            End
            Begin VB.Label foodChoiceLabel 
               Caption         =   "Food-Choice :"
               Enabled         =   0   'False
               Height          =   255
               Left            =   240
               TabIndex        =   49
               Top             =   840
               Width           =   1095
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Room Details"
            Height          =   2055
            Left            =   5520
            TabIndex        =   44
            Top             =   360
            Width           =   3255
            Begin VB.ComboBox custRoomCombo 
               Height          =   315
               ItemData        =   "HRSmain.frx":218A
               Left            =   1560
               List            =   "HRSmain.frx":218C
               Sorted          =   -1  'True
               TabIndex        =   18
               Text            =   "Select"
               Top             =   360
               Width           =   855
            End
            Begin VB.ComboBox roomTypeCombo 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "HRSmain.frx":218E
               Left            =   1560
               List            =   "HRSmain.frx":21A1
               Sorted          =   -1  'True
               TabIndex        =   19
               Text            =   "Select"
               Top             =   840
               Width           =   1455
            End
            Begin VB.Frame Frame6 
               Caption         =   "AC Facility :"
               Height          =   615
               Left            =   240
               TabIndex        =   45
               Top             =   1200
               Width           =   2775
               Begin VB.OptionButton ACOption 
                  Caption         =   "AC"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   20
                  Top             =   240
                  Width           =   735
               End
               Begin VB.OptionButton nonACOption 
                  Caption         =   "Non-AC"
                  Height          =   255
                  Left            =   1560
                  TabIndex        =   21
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   975
               End
            End
            Begin VB.Label Label7 
               Caption         =   "Room No :"
               Height          =   255
               Left            =   240
               TabIndex        =   47
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label8 
               Caption         =   "Room Type :"
               Height          =   255
               Left            =   240
               TabIndex        =   46
               Top             =   840
               Width           =   975
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Customer Personal"
            Height          =   3615
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   5055
            Begin VB.TextBox custMobCodeText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   9
               Top             =   2880
               Width           =   495
            End
            Begin VB.TextBox custResCodeText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   7
               Top             =   2520
               Width           =   495
            End
            Begin VB.ComboBox custCountryCombo 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "HRSmain.frx":21E0
               Left            =   1800
               List            =   "HRSmain.frx":21E2
               Sorted          =   -1  'True
               TabIndex        =   3
               Text            =   "Select"
               Top             =   1080
               Width           =   2055
            End
            Begin VB.TextBox custPinText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   6
               Top             =   2160
               Width           =   1215
            End
            Begin VB.TextBox custNameText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   1
               Top             =   360
               Width           =   2895
            End
            Begin VB.TextBox custAddressText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   2
               Top             =   720
               Width           =   2895
            End
            Begin VB.TextBox custStateText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   4
               Top             =   1440
               Width           =   2055
            End
            Begin VB.TextBox custCityText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   5
               Top             =   1800
               Width           =   2055
            End
            Begin VB.TextBox custResText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2640
               TabIndex        =   8
               Top             =   2520
               Width           =   1215
            End
            Begin VB.TextBox custMobText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2640
               TabIndex        =   10
               Top             =   2880
               Width           =   1215
            End
            Begin VB.TextBox custEmailText 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   11
               Top             =   3240
               Width           =   2895
            End
            Begin VB.Frame Frame3 
               Caption         =   "Frame3"
               Height          =   15
               Left            =   4320
               TabIndex        =   35
               Top             =   4920
               Width           =   135
            End
            Begin VB.Label Label50 
               Alignment       =   2  'Center
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2280
               TabIndex        =   70
               Top             =   2880
               Width           =   375
            End
            Begin VB.Label Label24 
               Alignment       =   2  'Center
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2400
               TabIndex        =   69
               Top             =   2520
               Width           =   135
            End
            Begin VB.Label Label1 
               Caption         =   "Pin :"
               Height          =   255
               Left            =   240
               TabIndex        =   68
               Top             =   2160
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "Customer Name :"
               Height          =   255
               Left            =   240
               TabIndex        =   43
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label3 
               Caption         =   "Address :"
               Height          =   255
               Left            =   240
               TabIndex        =   42
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label4 
               Caption         =   "Country :"
               Height          =   255
               Left            =   240
               TabIndex        =   41
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label5 
               Caption         =   "State :"
               Height          =   255
               Left            =   240
               TabIndex        =   40
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label Label6 
               Caption         =   "City : "
               Height          =   255
               Left            =   240
               TabIndex        =   39
               Top             =   1800
               Width           =   495
            End
            Begin VB.Label Label9 
               Caption         =   "Contact No (Res) :"
               Height          =   255
               Left            =   240
               TabIndex        =   38
               Top             =   2520
               Width           =   1935
            End
            Begin VB.Label Label10 
               Caption         =   "Contact No (Mobile) :"
               Height          =   255
               Left            =   240
               TabIndex        =   37
               Top             =   2880
               Width           =   1575
            End
            Begin VB.Label Label11 
               Caption         =   "E-mail :"
               Height          =   255
               Left            =   240
               TabIndex        =   36
               Top             =   3240
               Width           =   1215
            End
         End
      End
      Begin VB.Frame searchCustomerFrame 
         Caption         =   "Search/Cancel Booking :"
         Height          =   1575
         Left            =   -72120
         TabIndex        =   139
         Top             =   3420
         Width           =   3735
         Begin VB.ComboBox searchCombo 
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   167
            Text            =   "Select"
            Top             =   960
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker searchDTPicker 
            Height          =   375
            Left            =   1680
            TabIndex        =   166
            Top             =   960
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20381697
            CurrentDate     =   38828
         End
         Begin VB.ComboBox searchCriteriaCombo 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "HRSmain.frx":21E4
            Left            =   1920
            List            =   "HRSmain.frx":21F4
            TabIndex        =   140
            Text            =   "Customer Name"
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label searchLabel 
            Caption         =   "Customer Name :"
            Height          =   255
            Left            =   240
            TabIndex        =   142
            Top             =   990
            Width           =   1335
         End
         Begin VB.Label searchCriteriaLabel 
            Caption         =   "Search / Booking cancellation criteria :"
            Height          =   495
            Left            =   240
            TabIndex        =   141
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton searchButton 
         Caption         =   "&Search / Cancel Booking"
         Height          =   375
         Left            =   -72120
         TabIndex        =   143
         Top             =   5220
         Width           =   2055
      End
      Begin VB.CommandButton cancelSearchButton 
         Caption         =   "&Cancel Operation"
         Height          =   375
         Left            =   -69960
         TabIndex        =   144
         Top             =   5220
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid RoomGrid 
         Height          =   5415
         Left            =   -74640
         TabIndex        =   165
         Top             =   540
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   9551
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         AllowBigSelection=   -1  'True
         Enabled         =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.Label srecordLabel 
         Caption         =   "Customer IDs :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   137
         Top             =   8760
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   59
      Top             =   9480
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   450
      SimpleText      =   "Enter the informations in the relevant fields to reserve room."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   9710
            MinWidth        =   9701
            Text            =   "Enter the informations in the relevant fields to reserve room."
            TextSave        =   "Enter the informations in the relevant fields to reserve room."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "5/22/2006"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "8:17 PM"
         EndProperty
      EndProperty
   End
   Begin VB.Image hotelImage 
      Height          =   9540
      Left            =   -1920
      Picture         =   "HRSmain.frx":222C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11940
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOperations 
         Caption         =   "Operations..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About us..."
         Shortcut        =   {F9}
      End
   End
End
Attribute VB_Name = "HRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbHotel As ADODB.Connection
Dim hotelRecord As ADODB.Recordset
Dim roomInfo As ADODB.Recordset

Dim ConnectionString As String
Dim customerID As String
Dim Rooms As Integer
Dim roomAmount As Double
Dim foodAmount As Double
Dim LineCounter As Integer
Dim recordCnt As Integer
Dim recordPresent As Boolean
Dim addItem As Boolean
Dim tSpent As Date
    
Private Sub ACOption_Click()
    Dim i As Integer
    custRoomCombo.Clear
    hotelRecord.Open "select room_no from room_info where room_type = '" & roomTypeCombo.Text & "' and is_Ac =" & ACOption.Value & " and room_status = false", dbHotel, adOpenStatic, adLockBatchOptimistic
    For i = 0 To hotelRecord.RecordCount - 1
        custRoomCombo.addItem hotelRecord.Fields("room_no").Value, i
        hotelRecord.MoveNext
    Next i
    hotelRecord.Close
    custRoomCombo.Text = "Select"
End Sub

Private Sub addButton_Click()
    Call validateFields
    Call createCustomerID
    
    If addItem = True Then
        hotelRecord.Open ("select * from customer_info"), dbHotel, adOpenKeyset, adLockOptimistic
        With hotelRecord
            .AddNew
                !customer_id = customerID
                !customer_name = custNameText.Text
                !address = custAddressText.Text
                !country = custCountryCombo.Text
                !State = custStateText.Text
                !city = custCityText.Text
                !pin = custPinText.Text
                !contact_res = custResCodeText.Text + custResText.Text
                !contact_mob = custMobCodeText.Text + custMobText.Text
                !email = custEmailText.Text
                !arrival_by = custArrCombo.Text
                !vehicle_no = custVehicleText.Text
                !arrival_date = custArrDTPicker.Value
                !arrival_time = custArrTMPicker.Value
                !departure_date = custDeptDTPicker.Value
                !departure_time = custDeptTMPicker.Value
                !food_choice = foodChoiceCombo.Text
                !is_accompanied = spouseOption.Value
                !relationship = relationCombo.Text
            .Update
        End With
        hotelRecord.Close
    
        roomInfo.Open ("select * from customer_room_info"), dbHotel, adOpenKeyset, adLockOptimistic
        With roomInfo
            .AddNew
                !room_no = custRoomCombo.Text
                !customer_id = customerID
            .Update
        End With
        roomInfo.Close
    
        hotelRecord.Open ("select * from transaction_info"), dbHotel, adOpenKeyset, adLockOptimistic
        With hotelRecord
            .AddNew
                !transaction_id = customerID
                !is_cash = cashOption.Value
                !is_creditcard = creditCardOption.Value
                !is_dd = DDOption.Value
                !creditcard_dd_no = typeNoText.Text
                !amount_payable = Val(amountText.Text)
            .Update
        End With
        hotelRecord.Close
    
        hotelRecord.Open ("select * from payment_info"), dbHotel, adOpenKeyset, adLockOptimistic
        With hotelRecord
            .AddNew
                !transaction_id = customerID
                !customer_id = customerID
            .Update
        End With
        hotelRecord.Close
    
        dbHotel.Execute ("update room_info set room_status = true where room_no ='" & custRoomCombo.Text & "'")
    
        Call erasecustomerInformation
    
        'Reset RoomGrid connection
        roomInfo.Open "Select * from room_info", dbHotel, adOpenStatic, adLockBatchOptimistic
        RenderGrid
    End If
End Sub

Private Sub bookingTimer_Timer()
    Dim i As Integer
    If custNameText.Text <> "" And custAddressText.Text <> "" And (custCountryCombo.Text <> "Select" Or custCountryCombo.Text <> "") And custStateText.Text <> "" And custCityText.Text <> "" And custPinText.Text <> "" And (custRoomCombo.Text <> "Select" Or custRoomCombo.Text <> "") Then
        addButton.Enabled = True
    End If
    If creditCardOption.Value = True Or DDOption.Value = True Then
        If typeNoText.Text <> "" Then
            addButton.Enabled = True
        End If
    End If
End Sub

Private Sub cancelBookingButton_Click()
    Dim i As Integer
    i = MsgBox("Are you sure to delete this record?", vbYesNo, "Cancel Booking...")
    If i = 6 Then
        dbHotel.Execute "delete * from customer_info where customer_id = '" & sCustIDText.Text & "'"
        dbHotel.Execute "delete * from customer_room_info where room_no ='" & sCustRoomNoText.Text & "'"
        dbHotel.Execute "delete * from transaction_info where transaction_id = (select transaction_id from payment_info where customer_id ='" & sCustIDText.Text & "')"
        dbHotel.Execute "delete * from payment_info where customer_id ='" & sCustIDText.Text & "'"
        dbHotel.Execute ("update room_info set room_status = false where room_no ='" & sCustRoomNoText.Text & "'")
    
        recordList.Clear
        hotelRecord.Open "select * from customer_info where customer_name = '" & searchCombo.Text & "'", dbHotel, adOpenStatic, adLockBatchOptimistic
        If hotelRecord.RecordCount > 0 Then
            Call listInformation(hotelRecord)
        Else
            hotelRecord.Close
            Call nextSearchButton_Click
        End If
        
        'Reset RoomGrid connection
        roomInfo.Open "Select * from room_info", dbHotel, adOpenStatic, adLockBatchOptimistic
        RenderGrid
    End If
End Sub

Private Sub cancelButton_Click()
    Call erasecustomerInformation
End Sub

Private Sub cashOption_Click()
    typeNoLabel.Visible = False
    typeNoText.Visible = False
End Sub

Private Sub creditCardOption_Click()
    typeNoLabel.Visible = True
    typeNoLabel.Caption = "Credit Card no :"
    typeNoText.Visible = True
End Sub

Private Sub custArrTMPicker_Click()
    arrivalTimer.Enabled = False
End Sub

Private Sub custArrTMPicker_LostFocus()
    arrivalTimer.Enabled = True
End Sub

Private Sub date_Click()
    Call timeSpent
End Sub

Private Sub custDeptDTPicker_Change()
    amountText.Text = (roomAmount + foodAmount) * (custDeptDTPicker.Value - custArrDTPicker.Value)
End Sub

Private Sub custRoomCombo_click()
    hotelRecord.Open ("select * from room_info where room_no = '" & custRoomCombo.Text & "'")
    If custRoomCombo.Text <> "Select" Then
        roomAmount = hotelRecord.Fields("Fare").Value
    End If
    amountText.Text = (roomAmount + foodAmount) * (custDeptDTPicker.Value - custArrDTPicker.Value)
    roomTypeCombo.Text = hotelRecord.Fields("room_type").Value
    ACOption.Value = hotelRecord.Fields("is_AC").Value
    If ACOption.Value = False Then
        nonACOption.Value = True
    End If
    hotelRecord.Close
End Sub

Private Sub DDOption_Click()
    typeNoLabel.Visible = True
    typeNoLabel.Caption = "Demand Draft no :"
    typeNoText.Visible = True
End Sub

Private Sub foodChoiceCombo_Click()
    hotelRecord.Open ("select price from food where food_type = '" & foodChoiceCombo.Text & "'")
    If foodChoiceCombo.Text <> "Select" Then
        foodAmount = hotelRecord.Fields("price").Value
    End If
    amountText.Text = (roomAmount + foodAmount) * (custDeptDTPicker.Value - custArrDTPicker.Value)
    hotelRecord.Close
End Sub

Private Sub Form_Load()
    custArrDTPicker.Value = Date$
    custDeptDTPicker.Value = Date$
    searchDTPicker.Value = Date$
    custArrTMPicker.Value = Time$
    custDeptTMPicker.Value = Time$
    
    Dim i As Integer
    
    roomAmount = 0
    foodAmount = 0
    
    'Set database connection
    Set dbHotel = CreateObject("ADODB.Connection")
    ConnectionString = "DSN=hotel"
    dbHotel.Open ConnectionString
         
    Set hotelRecord = New ADODB.Recordset
     
    'Set country records
    hotelRecord.Open "select country_name from country", dbHotel, adOpenStatic, adLockBatchOptimistic
    recordCnt = hotelRecord.RecordCount
    For i = 0 To recordCnt - 1
        custCountryCombo.addItem hotelRecord.Fields("country_name").Value, i
        hotelRecord.MoveNext
    Next i
    hotelRecord.Close
    
    'Set arrival type records
    hotelRecord.Open "select arrival_by from arrival", dbHotel, adOpenStatic, adLockBatchOptimistic
    recordCnt = hotelRecord.RecordCount
    For i = 0 To recordCnt - 1
        custArrCombo.addItem hotelRecord.Fields("arrival_by").Value, i
        hotelRecord.MoveNext
    Next i
    hotelRecord.Close
    
    'Set room no records
    hotelRecord.Open "select room_no from room_info where is_ac = " & ACOption.Value & " and room_status = false", dbHotel, adOpenStatic, adLockBatchOptimistic
    recordCnt = hotelRecord.RecordCount
    For i = 0 To recordCnt - 1
        custRoomCombo.addItem hotelRecord.Fields("room_no").Value, i
        hotelRecord.MoveNext
    Next i
    hotelRecord.Close
    
    'Set food type records
    hotelRecord.Open "select food_type from food", dbHotel, adOpenStatic, adLockBatchOptimistic
    recordCnt = hotelRecord.RecordCount
    For i = 0 To recordCnt - 1
        foodChoiceCombo.addItem hotelRecord.Fields("food_type").Value, i
        hotelRecord.MoveNext
    Next i
    hotelRecord.Close
    
    'Set relation type records
    hotelRecord.Open "select relation_type from relationship", dbHotel, adOpenStatic, adLockBatchOptimistic
    recordCnt = hotelRecord.RecordCount
    For i = 0 To recordCnt - 1
        relationCombo.addItem hotelRecord.Fields("relation_type").Value, i
        hotelRecord.MoveNext
    Next i
    hotelRecord.Close
    
    'Set RoomGrid connection
    Set roomInfo = New ADODB.Recordset
    roomInfo.Open "Select * from room_info", dbHotel, adOpenStatic, adLockBatchOptimistic
    RenderGrid
        
    searchCriteriaCombo.ListIndex = 0
    statusBar.Panels.Item(1).Text = "To start operating press F5 or go to File->Operations."
    splashScreen.Hide
    Unload splashScreen
End Sub

Private Sub hotelOption_Click()
    Dim i As Integer
    foodChoiceLabel.Enabled = True
    foodChoiceCombo.Enabled = True
    includingFoodLabel.Visible = True
End Sub

Private Sub createCustomerID()
    Dim length, i As Integer
    Dim cid As String
    customerID = ""
    cid = Str(custArrDTPicker.Value) + Str(custArrTMPicker.Value)
    length = Len(cid)
    For i = 1 To length
        If Asc(Mid$(cid, i, 1)) > 47 And Asc(Mid$(cid, i, 1)) < 58 Then
            customerID = customerID + Mid$(cid, i, 1)
        End If
    Next i
    If (custArrDTPicker.Value - custDeptDTPicker.Value) > 0 Then
        i = MsgBox("Either Departure date or Departure time is wrong", vbInformation, "Wrong Date...")
    End If
    i = MsgBox("The customer ID of " & custNameText.Text & " is  : " & customerID, vbInformation, "Customer ID")
End Sub

Private Sub mnuAbout_Click()
    aboutUs.Show
End Sub

Private Sub mnuExit_Click()
    Unload HRS
    Unload aboutUs
End Sub

Private Sub mnuOperations_Click()
    tab1.Visible = True
    mnuOperations.Enabled = False
    hotelImage.Visible = False
    statusBar.Panels.Item(1).Text = "Enter relevant field informations to reserve, search, cancel room."
End Sub

Private Sub mnuUse_Click()
    
End Sub

Private Sub nextSearchButton_Click()
    searchCustomerFrame.Visible = True
    searchButton.Visible = True
    cancelSearchButton.Visible = True
    recordList.Visible = False
    customerInformationFrame.Visible = False
    nextSearchButton.Visible = False
    srecordLabel.Visible = False
    updateAmount.Visible = False
    cancelBookingButton.Visible = False
    recordList.Clear
    searchCriteriaCombo_Click
End Sub

Private Sub nonACOption_Click()
    Dim i As Integer
    custRoomCombo.Clear
    hotelRecord.Open "select room_no from room_info where room_type = '" & roomTypeCombo.Text & "' and is_Ac =" & ACOption.Value & " and room_status = false", dbHotel, adOpenStatic, adLockBatchOptimistic
    For i = 0 To hotelRecord.RecordCount - 1
        custRoomCombo.addItem hotelRecord.Fields("room_no").Value, i
        hotelRecord.MoveNext
    Next i
    hotelRecord.Close
    custRoomCombo.Text = "Select"
End Sub

Private Sub Opt_All_Click()
    Dim i As Integer
    roomInfo.Open "Select * from room_info", dbHotel, adOpenStatic, adLockBatchOptimistic
    If roomInfo.RecordCount > 0 Then
        RenderGrid
    Else
        i = MsgBox("There are no room in the hotel.", vbInformation, "Room availability")
        roomInfo.Close
    End If
End Sub

Private Sub Opt_Free_Click()
    Dim i As Integer
    roomInfo.Open "Select * from room_info where room_status = false", dbHotel, adOpenStatic, adLockBatchOptimistic
    If roomInfo.RecordCount > 0 Then
        RenderGrid
    Else
        i = MsgBox("No FREE rooms are avilable", vbInformation, "Room availability")
        roomInfo.Close
    End If
End Sub

Private Sub Opt_InUse_Click()
    Dim i As Integer
    roomInfo.Open "Select * from room_info where room_status = true", dbHotel, adOpenStatic, adLockBatchOptimistic
    If roomInfo.RecordCount > 0 Then
        RenderGrid
    Else
        i = MsgBox("No room is In USE", vbInformation, "Room availability")
        roomInfo.Close
    End If
End Sub

Private Sub outsideOption_Click()
    foodChoiceLabel.Enabled = False
    foodChoiceCombo.Enabled = False
    includingFoodLabel.Visible = False
    foodAmount = 0
    amountText.Text = (roomAmount + foodAmount) * (custDeptDTPicker.Value - custArrDTPicker.Value)
End Sub

Private Sub recordList_Click()
    hotelRecord.Open "select * from customer_info where customer_id = '" & recordList.List(recordList.ListIndex) & "'", dbHotel, adOpenStatic, adLockBatchOptimistic
    Call validateRecord(hotelRecord)
    If recordPresent = True Then
        Call customerInformation(hotelRecord)
        roomInfo.Open "select * from room_info where room_no = (select room_no from customer_room_info where customer_id = '" & recordList.List(recordList.ListIndex) & "')", dbHotel, adOpenStatic, adLockBatchOptimistic
        Call roomInformation(roomInfo)
        hotelRecord.Open "select * from transaction_info where transaction_id = (select transaction_id from payment_info where customer_id = '" & recordList.List(recordList.ListIndex) & "')", dbHotel, adOpenStatic, adLockBatchOptimistic
        Call paymentInformation(hotelRecord)
    End If
End Sub

Private Sub roomTypeCombo_Click()
    Dim i As Integer
    custRoomCombo.Clear
    hotelRecord.Open "select room_no from room_info where room_type = '" & roomTypeCombo.Text & "' and is_Ac =" & ACOption.Value & " and room_status = false", dbHotel, adOpenStatic, adLockBatchOptimistic
    For i = 0 To hotelRecord.RecordCount - 1
        custRoomCombo.addItem hotelRecord.Fields("room_no").Value, i
        hotelRecord.MoveNext
    Next i
    hotelRecord.Close
    custRoomCombo.Text = "Select"
End Sub

Private Sub searchButton_Click()
    Dim i As Integer
    recordPresent = False
    If searchCriteriaCombo.ListIndex = 3 Then
        hotelRecord.Open "select * from customer_info where arrival_date = #" & searchDTPicker.Value & "#", dbHotel, adOpenStatic, adLockBatchOptimistic
        Call validateRecord(hotelRecord)
        If recordPresent = True Then
            recordList.Visible = True
            srecordLabel.Visible = True
            Call listInformation(hotelRecord)
        End If
    Else
        If searchCombo.Text = "Select" Or searchCombo.Text = "" Then
            i = MsgBox("Enter any string in search field.", vbInformation, "Search...")
        Else
            'Take information from database
            Select Case searchCriteriaCombo.ListIndex
        
            Case 0: hotelRecord.Open "select * from customer_info where customer_name = '" & searchCombo.Text & "'", dbHotel, adOpenStatic, adLockBatchOptimistic
                    Call validateRecord(hotelRecord)
                    If recordPresent = True Then
                        recordList.Visible = True
                        srecordLabel.Visible = True
                        Call listInformation(hotelRecord)
                    End If
        
            Case 1: hotelRecord.Open "select * from customer_info where customer_id = '" & searchCombo.Text & "'", dbHotel, adOpenStatic, adLockBatchOptimistic
                    Call validateRecord(hotelRecord)
                    If recordPresent = True Then
                        recordList.Visible = False
                        srecordLabel.Visible = False
                        Call customerInformation(hotelRecord)
                        roomInfo.Open "select * from room_info where room_no = (select room_no from customer_room_info where customer_id = '" & searchCombo.Text & "')", dbHotel, adOpenStatic, adLockBatchOptimistic
                        Call roomInformation(roomInfo)
                        hotelRecord.Open "select * from transaction_info where transaction_id = (select transaction_id from payment_info where customer_id = '" & searchCombo.Text & "')", dbHotel, adOpenStatic, adLockBatchOptimistic
                        Call paymentInformation(hotelRecord)
                    End If
    
            Case 2: hotelRecord.Open "select * from customer_info where customer_id = (select customer_id from customer_room_info where room_no = '" & searchCombo.Text & "')", dbHotel, adOpenStatic, adLockBatchOptimistic
                    Call validateRecord(hotelRecord)
                    If recordPresent = True Then
                        recordList.Visible = False
                        srecordLabel.Visible = False
                        Call customerInformation(hotelRecord)
                        roomInfo.Open "select * from room_info where room_no ='" & searchCombo.Text & "'", dbHotel, adOpenStatic, adLockBatchOptimistic
                        Call roomInformation(roomInfo)
                        hotelRecord.Open "select * from transaction_info where transaction_id = (select transaction_id from payment_info where customer_id = '" & sCustIDText.Text & "')", dbHotel, adOpenStatic, adLockBatchOptimistic
                        Call paymentInformation(hotelRecord)
                    End If
            End Select
        End If
    End If
    
        If recordPresent = True Then
            searchCustomerFrame.Visible = False
            searchButton.Visible = False
            cancelSearchButton.Visible = False
            customerInformationFrame.Visible = True
            nextSearchButton.Visible = True
            updateAmount.Visible = True
            cancelBookingButton.Visible = True
        End If
    recordPresent = False
End Sub

Private Sub searchCriteriaCombo_Click()
    Dim Index, flag As Integer
    Dim i, j, k As Integer
    flag = 0
    k = 0
    searchCombo.Clear
    Select Case searchCriteriaCombo.ListIndex
        Case 0: searchLabel.Caption = "Customer Name : "
                searchDTPicker.Visible = False
                searchCombo.Visible = True
                hotelRecord.Open "select customer_name from customer_info", dbHotel, adOpenStatic, adLockBatchOptimistic
                recordCnt = hotelRecord.RecordCount
                For i = 0 To recordCnt - 1
                    For j = 0 To i
                        If searchCombo.List(0) = "" Then
                            searchCombo.addItem hotelRecord.Fields("customer_name").Value, k
                            k = k + 1
                            flag = 1
                        Else
                            If searchCombo.List(j) = hotelRecord.Fields("customer_name").Value Then
                                flag = 1
                            End If
                        End If
                    Next j
                    If flag = 0 Then
                        searchCombo.addItem hotelRecord.Fields("customer_name").Value, k
                        k = k + 1
                    End If
                    flag = 0
                    hotelRecord.MoveNext
                Next i
                hotelRecord.Close
        Case 1: searchLabel.Caption = "Customer ID : "
                searchDTPicker.Visible = False
                searchCombo.Visible = True
                hotelRecord.Open "select customer_id from customer_info", dbHotel, adOpenStatic, adLockBatchOptimistic
                recordCnt = hotelRecord.RecordCount
                For i = 0 To recordCnt - 1
                    searchCombo.addItem hotelRecord.Fields("customer_id").Value, i
                    hotelRecord.MoveNext
                Next i
                hotelRecord.Close
        Case 2: searchLabel.Caption = "Room no : "
                searchDTPicker.Visible = False
                searchCombo.Visible = True
                hotelRecord.Open "select room_no from customer_room_info", dbHotel, adOpenStatic, adLockBatchOptimistic
                recordCnt = hotelRecord.RecordCount
                For i = 0 To recordCnt - 1
                    searchCombo.addItem hotelRecord.Fields("room_no").Value, i
                    hotelRecord.MoveNext
                Next i
                hotelRecord.Close
        Case 3: searchLabel.Caption = "Booking Date : "
                searchDTPicker.Visible = True
                searchCombo.Visible = False
    End Select
    searchCombo.Text = "Select"
End Sub

Private Sub spouseOption_Click()
    Dim i As Integer
    relationshipLabel.Enabled = True
    relationCombo.Enabled = True
End Sub

Private Sub arrivalTimer_Timer()
    custArrTMPicker.Value = Time$
End Sub

Private Sub tab1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuFile
    End If
End Sub

Private Sub updateAmount_Click()
    Dim i As Integer
    i = InputBox("Enter amount to be paid : ", "Make Payment")
    dbHotel.Execute "update transaction_info set amount_payable = " & (Val(sCustAmountText.Text) - i) & " where transaction_id = '" & sCustIDText.Text & "'"
    sCustAmountText.Text = Val(sCustAmountText.Text) - i
End Sub

Private Sub withoutSpouseOption_Click()
    relationshipLabel.Enabled = False
    relationCombo.Enabled = False
End Sub

Public Function RenderGrid()
    Dim i As Integer
    
    recordCnt = roomInfo.RecordCount
    LineCounter = 1

    RoomGrid.Clear
    RoomGrid.Cols = 5
    
    'Until max record for rooms
    RoomGrid.Rows = recordCnt + 1
    
    RoomGrid.FixedCols = 1
    RoomGrid.FixedRows = 1
    
    RoomGrid.ColWidth(0) = 1000
    RoomGrid.ColWidth(1) = 3487
    RoomGrid.ColWidth(2) = 1000
    RoomGrid.ColWidth(3) = 1500
    RoomGrid.ColWidth(4) = 1500
    
 
    For i = 0 To 4
   
        RoomGrid.Row = 0
        RoomGrid.Col = i
        RoomGrid.CellFontBold = True
    
    Next i
    
    'Bold the room numberoomInfo
    For i = 1 To recordCnt
    
        RoomGrid.Row = i
        RoomGrid.Col = 0
        RoomGrid.Text = i
        RoomGrid.CellFontBold = True
        RoomGrid.CellAlignment = 3
    
    Next i
    
    'Set Alignment for each row.
    For i = 1 To recordCnt
    
        RoomGrid.Row = i
        RoomGrid.Col = 1
        RoomGrid.CellAlignment = 0
        RoomGrid.Col = 2
        RoomGrid.CellAlignment = 0
        RoomGrid.Col = 3
        RoomGrid.CellAlignment = 3
        RoomGrid.Col = 4
        RoomGrid.CellAlignment = 3
        
    Next i

    RoomGrid.TextMatrix(0, 0) = "Room No"
    RoomGrid.TextMatrix(0, 1) = "Room Type"
    RoomGrid.TextMatrix(0, 2) = "Fare"
    RoomGrid.TextMatrix(0, 3) = "Air Conditioned"
    RoomGrid.TextMatrix(0, 4) = "Room Status"
    
    Rooms = recordCnt
    
    For i = 1 To recordCnt
    
        RoomGrid.TextMatrix(LineCounter, 0) = roomInfo.Fields("room_no").Value
        RoomGrid.TextMatrix(LineCounter, 1) = roomInfo.Fields("room_type").Value
        RoomGrid.TextMatrix(LineCounter, 2) = roomInfo.Fields("fare").Value
        RoomGrid.TextMatrix(LineCounter, 3) = roomInfo.Fields("is_AC").Value
            
        If roomInfo.Fields("room_status").Value = True Then
            
            RoomGrid.Row = LineCounter
            RoomGrid.Col = 4
            RoomGrid.Text = "IN USE"
            RoomGrid.CellFontBold = True
            RoomGrid.CellBackColor = &HC0C0FF
            RoomGrid.CellForeColor = vbBlack
        
        Else

            RoomGrid.Row = LineCounter
            RoomGrid.Col = 4
            RoomGrid.Text = "FREE"
            RoomGrid.CellFontBold = True
            RoomGrid.CellBackColor = &HC0FFC0
            RoomGrid.CellForeColor = vbBlack
            
        End If
               
        LineCounter = LineCounter + 1
        roomInfo.MoveNext
        
    Next i
    
    roomInfo.Close
    
End Function

Private Sub RoomGrid_Click()
    RoomGrid.Row = RoomGrid.RowSel
    RoomGrid.Col = 0
    noCustomerLabel.Visible = False
    cNameLabel.Visible = True
    cAddressLabel.Visible = True
    cCountryLabel.Visible = True
    cStateLabel.Visible = True
    cCityLabel.Visible = True
    cPinLabel.Visible = True
    cMobileLabel.Visible = True
    cArrivalDTLabel.Visible = True
    cArrivalTMLabel.Visible = True
    cDeptDTLabel.Visible = True
    cDeptTMLabel.Visible = True
    
    'Initialising caption of labels
    cNameLabel.Caption = "Name : "
    cAddressLabel.Caption = "Address : "
    cCountryLabel.Caption = "Country : "
    cStateLabel.Caption = "State : "
    cCityLabel.Caption = "City : "
    cPinLabel.Caption = "Pin : "
    cMobileLabel.Caption = "Contact no.(Mobile) : "
    cArrivalDTLabel.Caption = "Arrival Date : "
    cArrivalTMLabel.Caption = "Arrival Time : "
    cDeptDTLabel.Caption = "Departure Date : "
    cDeptTMLabel.Caption = "Departure Time : "
    
    roomInfo.Open ("Select * from room_info where room_no = '" & RoomGrid.Text & "'")
    
    If roomInfo.Fields("room_status").Value = True Then
    
        hotelRecord.Open ("Select * from customer_info where customer_id = (Select customer_id from customer_room_info where room_no = '" & RoomGrid.Text & "')"), dbHotel, adOpenStatic, adLockBatchOptimistic
        
        cNameLabel.Caption = cNameLabel.Caption & hotelRecord.Fields("customer_name").Value
        cAddressLabel.Caption = cAddressLabel.Caption & hotelRecord.Fields("address").Value
        cCountryLabel.Caption = cCountryLabel.Caption & hotelRecord.Fields("country").Value
        cStateLabel.Caption = cStateLabel.Caption & hotelRecord.Fields("state").Value
        cCityLabel.Caption = cCityLabel.Caption & hotelRecord.Fields("city").Value
        cPinLabel.Caption = cPinLabel.Caption & hotelRecord.Fields("PIN").Value
        cMobileLabel.Caption = cMobileLabel.Caption & hotelRecord.Fields("contact_mob").Value
        cArrivalDTLabel.Caption = cArrivalDTLabel.Caption & hotelRecord.Fields("arrival_date").Value
        cArrivalTMLabel.Caption = cArrivalTMLabel.Caption & hotelRecord.Fields("arrival_time").Value
        cDeptDTLabel.Caption = cDeptDTLabel.Caption & hotelRecord.Fields("departure_date").Value
        cDeptTMLabel.Caption = cDeptTMLabel.Caption & hotelRecord.Fields("departure_time").Value
        
        hotelRecord.Close
    Else
    
        noCustomerLabel.Visible = True
        cNameLabel.Visible = False
        cAddressLabel.Visible = False
        cCountryLabel.Visible = False
        cStateLabel.Visible = False
        cCityLabel.Visible = False
        cPinLabel.Visible = False
        cMobileLabel.Visible = False
        cArrivalDTLabel.Visible = False
        cArrivalTMLabel.Visible = False
        cDeptDTLabel.Visible = False
        cDeptTMLabel.Visible = False
    End If
    
    If roomInfo.Fields("room_status").Value = True Then
    
        roomStatusLabel.Caption = "IN USE"
        roomStatusLabel.BackColor = "&HC0C0FF"
        
    Else
    
        roomStatusLabel.Caption = "FREE"
        roomStatusLabel.BackColor = "&HC0FFC0"
        
    End If
    
    roomInfo.Close
    
End Sub

Public Function erasecustomerInformation()
    custNameText.Text = ""
    custAddressText.Text = ""
    custCountryCombo.Text = "Select"
    custStateText.Text = ""
    custCityText.Text = ""
    custPinText.Text = ""
    custResText.Text = ""
    custMobText.Text = ""
    custResCodeText.Text = ""
    custMobCodeText.Text = ""
    custEmailText.Text = ""
    custArrCombo.Text = "Select"
    custVehicleText.Text = ""
    custArrDTPicker.Value = Date$
    custArrTMPicker.Value = Time$
    custDeptDTPicker.Value = Date$
    custDeptTMPicker.Value = Time$
    foodChoiceCombo.Text = "Select"
    custRoomCombo.Text = "Select"
    spouseOption.Value = False
    relationCombo.Text = "Select"
    cashOption.Value = False
    creditCardOption.Value = False
    DDOption.Value = False
    typeNoText.Text = ""
    amountText.Text = ""
End Function

Private Function timeSpent()
    tSpent = (custDeptDTPicker.Value - custArrDTPicker.Value) * 24 * 3600 + (custDeptTMPicker.Value - custArrTMPicker.Value)
    MsgBox (custArrTMPicker.Value)
    MsgBox (custDeptTMPicker.Value - custArrTMPicker.Value)
End Function

Public Function customerInformation(hotelRecord As ADODB.Recordset)
    sCustIDText.Text = hotelRecord.Fields("customer_id").Value
    sCustNameText.Text = hotelRecord.Fields("customer_name").Value
    sCustAddressText.Text = hotelRecord.Fields("address").Value
    sCustCountryText.Text = hotelRecord.Fields("country").Value
    sCustStateText.Text = hotelRecord.Fields("state").Value
    sCustCityText.Text = hotelRecord.Fields("city").Value
    sCustPinText.Text = hotelRecord.Fields("pin").Value
    sCustResText.Text = hotelRecord.Fields("contact_res").Value
    sCustMobileText.Text = hotelRecord.Fields("contact_mob").Value
    sCustEmailText.Text = hotelRecord.Fields("email").Value
    sCustArrivalText.Text = hotelRecord.Fields("arrival_by").Value
    sCustVehicleNoText.Text = hotelRecord.Fields("vehicle_no").Value
    sCustEmailText.Text = hotelRecord.Fields("email").Value
    sCustArrDTText.Text = hotelRecord.Fields("arrival_date").Value
    sCustArrTMtext.Text = hotelRecord.Fields("arrival_time").Value
    sCustDeptDTText.Text = hotelRecord.Fields("departure_date").Value
    sCustDeptTMText.Text = hotelRecord.Fields("departure_time").Value
    sCustFoodChoiceText.Text = hotelRecord.Fields("food_choice").Value
    If sCustFoodChoiceText.Text <> "" Then
        sFoodChoiceLabel.Enabled = True
        sCustHotelOption.Value = True
    Else
        sFoodChoiceLabel.Enabled = False
        sCustHotelOption.Value = False
    End If
    sCustSpouseOption.Value = hotelRecord.Fields("is_accompanied").Value
    If sCustSpouseOption.Value = True Then
        sRelationLabel.Enabled = True
    Else
        sRelationLabel.Enabled = False
    End If
        sCustRelationText.Text = hotelRecord.Fields("relationship").Value
        hotelRecord.Close
End Function

Public Function roomInformation(roomInfo As ADODB.Recordset)
    sCustRoomNoText.Text = roomInfo.Fields("room_no").Value
    sCustRoomTypeText.Text = roomInfo.Fields("room_type").Value
    sCustACOption.Value = roomInfo.Fields("is_AC").Value
    If sCustACOption.Value = False Then
        sCustNACOption.Value = True
    End If
    roomInfo.Close
End Function

Public Function paymentInformation(hotelRecord As ADODB.Recordset)
    sCustAmountText.Text = hotelRecord.Fields("amount_payable").Value
    sCustCashOption.Value = hotelRecord.Fields("is_cash").Value
    sCustDDOption.Value = hotelRecord.Fields("is_dd").Value
    sCustCreditOption.Value = hotelRecord.Fields("is_creditcard").Value
    If sCustDDOption.Value = True Or sCustCreditOption.Value = True Then
        sCustTypeNoText.Visible = True
        sCustTypeNoText.Text = hotelRecord.Fields("creditcard_dd_no").Value
        sCustTypeNoLabel.Visible = True
    Else
        sCustTypeNoText.Visible = False
        sCustTypeNoLabel.Visible = False
    End If
    If sCustHotelOption.Value = True Then
        sIncludingFoodLabel.Visible = True
    Else
        sIncludingFoodLabel.Visible = False
    End If
    hotelRecord.Close
End Function

Public Function listInformation(hotelRecord As ADODB.Recordset)
    Dim i As Integer
    For i = 0 To hotelRecord.RecordCount - 1
        recordList.addItem hotelRecord.Fields("customer_id").Value, i
        hotelRecord.MoveNext
    Next i
    hotelRecord.Close
    recordList.Text = recordList.List(0)
    
    hotelRecord.Open "select * from customer_info where customer_id = '" & recordList.List(0) & "'", dbHotel, adOpenStatic, adLockBatchOptimistic
    Call customerInformation(hotelRecord)
    roomInfo.Open "select * from room_info where room_no = (select room_no from customer_room_info where customer_id = '" & recordList.List(0) & "')", dbHotel, adOpenStatic, adLockBatchOptimistic
    Call roomInformation(roomInfo)
    hotelRecord.Open "select * from transaction_info where transaction_id = (select transaction_id from payment_info where customer_id = '" & recordList.List(0) & "')", dbHotel, adOpenStatic, adLockBatchOptimistic
    Call paymentInformation(hotelRecord)
End Function

Public Function validateRecord(hotelRecord As ADODB.Recordset)
    Dim i As Integer
    If hotelRecord.RecordCount = 0 Then
        i = MsgBox("Record not found.", vbInformation, "No Record...")
        hotelRecord.Close
        recordPresent = False
    Else
        recordPresent = True
    End If
End Function

Public Function validateFields()
    Dim i As Integer
    If hotelOption.Value = True Then
        If foodChoiceCombo.Text = "Select" Or foodChoiceCombo.Text = "" Then
            i = MsgBox("Choose food type.", vbInformation, "Food Choice...")
            addItem = False
        Else
            addItem = True
        End If
    End If
    If spouseOption.Value = True Then
        If relationCombo.Text = "" Or relationCombo.Text = "Select" Then
            i = MsgBox("Choose relationship.", vbInformation, "Relation...")
            addItem = False
        Else
            addItem = True
        End If
    End If
End Function
