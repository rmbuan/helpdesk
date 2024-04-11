VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9915
   ClientLeft      =   4215
   ClientTop       =   1020
   ClientWidth     =   11205
   Icon            =   "Frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   9645
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14579
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "23/06/00"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "10:45 AM"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab stContainer 
      Height          =   9495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   16748
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Problem Report"
      TabPicture(0)   =   "Frmmain.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmPR"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Look Problem(s)"
      TabPicture(1)   =   "Frmmain.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmSPSearch"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Print Report"
      TabPicture(2)   =   "Frmmain.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmPRepCP"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "New Problem(s)"
      TabPicture(3)   =   "Frmmain.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frmNPMain"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Admin Fuction(s)"
      TabPicture(4)   =   "Frmmain.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "frmAFMain"
      Tab(4).ControlCount=   1
      Begin VB.Frame frmAFMain 
         Height          =   9015
         Left            =   -74880
         TabIndex        =   50
         Top             =   360
         Width           =   10815
         Begin VB.TextBox txtFindAdmin 
            Height          =   375
            Left            =   7560
            TabIndex        =   124
            Text            =   "1"
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CommandButton cmdFindAdmin 
            Caption         =   "Find Text"
            Height          =   375
            Left            =   6480
            TabIndex        =   123
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton cmdSortAdmin 
            Caption         =   "Sort"
            Height          =   375
            Left            =   9720
            TabIndex        =   122
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton cmdNewUser 
            Caption         =   "Create New User"
            Height          =   375
            Left            =   9240
            TabIndex        =   121
            Top             =   360
            Width           =   1455
         End
         Begin VB.Frame frmAFCP 
            Caption         =   "Completed Problems"
            Height          =   1215
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   5055
            Begin VB.TextBox txtCompletedByTech 
               Height          =   285
               Left            =   2640
               TabIndex        =   87
               Top             =   720
               Width           =   2295
            End
            Begin VB.ComboBox cmbAFCCP 
               Height          =   315
               Left            =   120
               TabIndex        =   53
               Top             =   720
               Width           =   2295
            End
            Begin VB.Label lblAFBTech 
               AutoSize        =   -1  'True
               Caption         =   "By Tech"
               Height          =   195
               Left            =   2640
               TabIndex        =   55
               Top             =   360
               Width           =   600
            End
            Begin VB.Label lblAFCCP 
               AutoSize        =   -1  'True
               Caption         =   "Current Completed Problems"
               Height          =   195
               Left            =   120
               TabIndex        =   54
               Top             =   360
               Width           =   1995
            End
         End
         Begin MSFlexGridLib.MSFlexGrid MFGAdmin 
            Bindings        =   "Frmmain.frx":04CE
            Height          =   7095
            Left            =   120
            TabIndex        =   120
            Top             =   1560
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   12515
            _Version        =   393216
            FixedCols       =   0
            WordWrap        =   -1  'True
            TextStyle       =   2
            TextStyleFixed  =   2
            GridLines       =   2
            AllowUserResizing=   3
         End
         Begin MSComctlLib.ProgressBar pBarAdmin 
            Height          =   255
            Left            =   120
            TabIndex        =   125
            Top             =   8640
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
      End
      Begin VB.Frame frmNPMain 
         Height          =   9015
         Left            =   -74880
         TabIndex        =   40
         Top             =   360
         Width           =   10815
         Begin VB.TextBox txtNPPStatus 
            Height          =   375
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtNPDTReq 
            Height          =   375
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtNPPPriority 
            Height          =   375
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   600
            Width           =   1335
         End
         Begin VB.Frame frmNPFixEdit 
            Caption         =   "This frame has the options of the current problem that the technician needs to fix"
            Height          =   7815
            Left            =   120
            TabIndex        =   43
            Top             =   1080
            Visible         =   0   'False
            Width           =   10575
            Begin VB.CommandButton cmdNPFixEdit 
               Caption         =   "Fix / Edit"
               Height          =   375
               Left            =   9240
               TabIndex        =   119
               Top             =   720
               Width           =   1215
            End
            Begin VB.Frame frmMemo 
               Caption         =   "Explaination for the fixes"
               Height          =   3255
               Left            =   120
               TabIndex        =   117
               Top             =   4440
               Width           =   10335
               Begin RichTextLib.RichTextBox RTBFixes 
                  Height          =   2895
                  Left            =   120
                  TabIndex        =   118
                  ToolTipText     =   "Put in your fixes notes here!"
                  Top             =   240
                  Width           =   10095
                  _ExtentX        =   17806
                  _ExtentY        =   5106
                  _Version        =   393217
                  Enabled         =   -1  'True
                  TextRTF         =   $"Frmmain.frx":04E4
               End
            End
            Begin VB.Frame frmDOA 
               Height          =   2055
               Left            =   120
               TabIndex        =   114
               Top             =   2280
               Width           =   10335
               Begin VB.TextBox txtDOUP 
                  Height          =   1575
                  Left            =   120
                  MultiLine       =   -1  'True
                  TabIndex        =   115
                  ToolTipText     =   "Enter a brief description of the problem in your own words."
                  Top             =   360
                  Width           =   10095
               End
               Begin VB.Label lblDOUP 
                  AutoSize        =   -1  'True
                  Caption         =   "Description of User Problem"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   116
                  ToolTipText     =   "Enter a brief description of the problem in your own words."
                  Top             =   120
                  Width           =   1965
               End
            End
            Begin VB.TextBox txtOS 
               Height          =   375
               Left            =   5160
               TabIndex        =   113
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox txtSoftware 
               Height          =   375
               Left            =   5160
               TabIndex        =   111
               Top             =   840
               Width           =   1815
            End
            Begin VB.TextBox txtHardware 
               Height          =   375
               Left            =   5160
               TabIndex        =   109
               Top             =   360
               Width           =   1815
            End
            Begin VB.TextBox txtCompLoc 
               Height          =   375
               Left            =   1800
               TabIndex        =   107
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox txtCampusFix 
               Height          =   375
               Left            =   1800
               TabIndex        =   105
               Top             =   1800
               Width           =   1815
            End
            Begin VB.TextBox txtCompID 
               Height          =   375
               Left            =   1800
               TabIndex        =   103
               Top             =   840
               Width           =   1815
            End
            Begin VB.TextBox txtReqBy 
               Height          =   375
               Left            =   1800
               TabIndex        =   101
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label lblOS 
               AutoSize        =   -1  'True
               Caption         =   "OS:"
               Height          =   195
               Left            =   4080
               TabIndex        =   112
               Top             =   1440
               Width           =   270
            End
            Begin VB.Label lblSoftware 
               AutoSize        =   -1  'True
               Caption         =   "Software:"
               Height          =   195
               Left            =   4080
               TabIndex        =   110
               Top             =   960
               Width           =   675
            End
            Begin VB.Label lblHardware 
               AutoSize        =   -1  'True
               Caption         =   "Hardware:"
               Height          =   195
               Left            =   4080
               TabIndex        =   108
               Top             =   480
               Width           =   735
            End
            Begin VB.Label lblCompLoc 
               AutoSize        =   -1  'True
               Caption         =   "Computer Location:"
               Height          =   195
               Left            =   240
               TabIndex        =   106
               Top             =   1440
               Width           =   1380
            End
            Begin VB.Label lblCampusFix 
               AutoSize        =   -1  'True
               Caption         =   "Campus:"
               Height          =   195
               Left            =   240
               TabIndex        =   104
               Top             =   1920
               Width           =   615
            End
            Begin VB.Label lblCompID 
               AutoSize        =   -1  'True
               Caption         =   "Computer ID:"
               Height          =   195
               Left            =   240
               TabIndex        =   102
               Top             =   960
               Width           =   930
            End
            Begin VB.Label lblReqBy 
               AutoSize        =   -1  'True
               Caption         =   "Requested by:"
               Height          =   195
               Left            =   240
               TabIndex        =   100
               Top             =   480
               Width           =   1035
            End
         End
         Begin VB.ComboBox cmbNPAllProblems 
            Height          =   315
            Left            =   120
            TabIndex        =   42
            Top             =   600
            Width           =   3015
         End
         Begin VB.Label lblNPPStatus 
            AutoSize        =   -1  'True
            Caption         =   "Problem Status"
            Height          =   195
            Left            =   3360
            TabIndex        =   48
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label lblNPDTReq 
            AutoSize        =   -1  'True
            Caption         =   "Date/Time Requested"
            Height          =   195
            Left            =   6240
            TabIndex        =   46
            Top             =   240
            Width           =   1590
         End
         Begin VB.Label lblNPPPriority 
            AutoSize        =   -1  'True
            Caption         =   "Problem Priority"
            Height          =   195
            Left            =   4800
            TabIndex        =   44
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lblNPAllProblems 
            AutoSize        =   -1  'True
            Caption         =   "List of all new Problems"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   1650
         End
      End
      Begin VB.Frame frmPRepCP 
         Caption         =   "Current Problems to Print"
         Height          =   9015
         Left            =   -74880
         TabIndex        =   29
         Top             =   360
         Width           =   10815
         Begin VB.CommandButton cmbPRepCancel 
            Caption         =   "Cancel"
            Height          =   495
            Left            =   8640
            TabIndex        =   34
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmbPRepPSetup 
            Caption         =   "Printer Setup"
            Height          =   495
            Left            =   6960
            TabIndex        =   33
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmbPRepPreview 
            Caption         =   "Preview"
            Height          =   495
            Left            =   5280
            TabIndex        =   99
            Top             =   360
            Width           =   1695
         End
         Begin RichTextLib.RichTextBox rtbPRView 
            Height          =   7935
            Left            =   120
            TabIndex        =   51
            Top             =   960
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   13996
            _Version        =   393217
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"Frmmain.frx":059E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmbPRepPrint 
            Caption         =   "Print Now"
            Height          =   495
            Left            =   3600
            TabIndex        =   32
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox cmbPRepPrintSel 
            Height          =   315
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label lblPRepPrint 
            AutoSize        =   -1  'True
            Caption         =   "Select the Problem you want to print:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   2595
         End
      End
      Begin VB.Frame frmSPSearch 
         Height          =   9015
         Left            =   -74880
         TabIndex        =   27
         Top             =   360
         Width           =   10815
         Begin VB.Frame frmSPResult 
            Caption         =   "Result of search"
            Height          =   8775
            Left            =   120
            TabIndex        =   28
            Top             =   120
            Width           =   10575
            Begin VB.CommandButton cmdSort 
               Caption         =   "Sort"
               Height          =   375
               Left            =   3360
               TabIndex        =   98
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton cmdFindTxt 
               Caption         =   "Find Text"
               Height          =   375
               Left            =   120
               TabIndex        =   97
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox txtSearch 
               Height          =   375
               Left            =   1200
               TabIndex        =   96
               Text            =   "1"
               Top             =   240
               Width           =   1575
            End
            Begin VB.Data dataDAO 
               Caption         =   "Data1"
               Connect         =   "Access"
               DatabaseName    =   "D:\Projects\Software\HelpDesk\HelpDesk.mdb"
               DefaultCursorType=   0  'DefaultCursor
               DefaultType     =   2  'UseODBC
               Exclusive       =   0   'False
               Height          =   345
               Left            =   9240
               Options         =   0
               ReadOnly        =   0   'False
               RecordsetType   =   1  'Dynaset
               RecordSource    =   "tblProb"
               Top             =   8400
               Visible         =   0   'False
               Width           =   1215
            End
            Begin MSFlexGridLib.MSFlexGrid MFGPrint 
               Bindings        =   "Frmmain.frx":065A
               Height          =   7695
               Left            =   120
               TabIndex        =   94
               Top             =   720
               Width           =   10335
               _ExtentX        =   18230
               _ExtentY        =   13573
               _Version        =   393216
               FixedCols       =   0
               AllowUserResizing=   1
            End
            Begin MSComctlLib.ProgressBar pBarPrint 
               Height          =   255
               Left            =   120
               TabIndex        =   95
               Top             =   8400
               Width           =   10335
               _ExtentX        =   18230
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   1
            End
         End
      End
      Begin VB.Frame frmPR 
         Height          =   9015
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   10815
         Begin MSComDlg.CommonDialog CD 
            Left            =   360
            Top             =   8520
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   375
            Left            =   6480
            TabIndex        =   93
            Top             =   8520
            Width           =   1335
         End
         Begin VB.Frame frmUserItems 
            Caption         =   "User Info"
            Height          =   3015
            Left            =   120
            TabIndex        =   62
            Top             =   120
            Width           =   10575
            Begin VB.Frame frmInfo 
               Height          =   2655
               Left            =   120
               TabIndex        =   71
               Top             =   240
               Width           =   3255
               Begin VB.TextBox txtPassword 
                  Height          =   405
                  IMEMode         =   3  'DISABLE
                  Left            =   1200
                  Locked          =   -1  'True
                  PasswordChar    =   "*"
                  TabIndex        =   88
                  ToolTipText     =   "Enter your Full Name here for Identification"
                  Top             =   960
                  Width           =   1575
               End
               Begin VB.TextBox txtFirstName 
                  Height          =   405
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   76
                  ToolTipText     =   "Enter your First Name here for Identification"
                  Top             =   1920
                  Width           =   1575
               End
               Begin VB.TextBox txtLastName 
                  Height          =   405
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   74
                  ToolTipText     =   "Enter your Last Name here for Identification"
                  Top             =   1440
                  Width           =   1575
               End
               Begin VB.TextBox txtPRFullName 
                  Height          =   405
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   72
                  ToolTipText     =   "Enter your Full Name here for Identification"
                  Top             =   480
                  Width           =   1575
               End
               Begin VB.Label lblPassword 
                  AutoSize        =   -1  'True
                  Caption         =   "Password:"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   89
                  ToolTipText     =   "Enter your Full Name here for Identification"
                  Top             =   1080
                  Width           =   735
               End
               Begin VB.Label lblFirstName 
                  AutoSize        =   -1  'True
                  Caption         =   "First Name:"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   77
                  ToolTipText     =   "Enter your Full Name here for Identification"
                  Top             =   2040
                  Width           =   795
               End
               Begin VB.Label lblLastName 
                  AutoSize        =   -1  'True
                  Caption         =   "Last Name:"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   75
                  ToolTipText     =   "Enter your Full Name here for Identification"
                  Top             =   1560
                  Width           =   810
               End
               Begin VB.Label lblPRLogin 
                  AutoSize        =   -1  'True
                  Caption         =   "Login:"
                  Height          =   195
                  Left            =   600
                  TabIndex        =   73
                  ToolTipText     =   "Enter your Full Name here for Identification"
                  Top             =   600
                  Width           =   435
               End
            End
            Begin VB.Frame frmAddress 
               Height          =   1455
               Left            =   3960
               TabIndex        =   78
               Top             =   1440
               Width           =   6255
               Begin VB.TextBox txtPostalCode 
                  Height          =   405
                  Left            =   3960
                  Locked          =   -1  'True
                  TabIndex        =   84
                  ToolTipText     =   "Enter your Student ID here for Identification"
                  Top             =   840
                  Width           =   2055
               End
               Begin VB.TextBox txtProvince 
                  Height          =   405
                  Left            =   3960
                  Locked          =   -1  'True
                  TabIndex        =   83
                  ToolTipText     =   "Enter your Full Name here for Identification"
                  Top             =   360
                  Width           =   2055
               End
               Begin VB.TextBox txtCity 
                  Height          =   405
                  Left            =   720
                  Locked          =   -1  'True
                  TabIndex        =   80
                  ToolTipText     =   "Enter your Student ID here for Identification"
                  Top             =   840
                  Width           =   2055
               End
               Begin VB.TextBox txtStreet 
                  Height          =   405
                  Left            =   720
                  Locked          =   -1  'True
                  TabIndex        =   79
                  ToolTipText     =   "Enter your Full Name here for Identification"
                  Top             =   360
                  Width           =   2055
               End
               Begin VB.Label lblPostalCode 
                  AutoSize        =   -1  'True
                  Caption         =   "Postal Code:"
                  Height          =   195
                  Left            =   2880
                  TabIndex        =   86
                  ToolTipText     =   "Enter your Student ID here for Identification"
                  Top             =   960
                  Width           =   900
               End
               Begin VB.Label lblProvince 
                  AutoSize        =   -1  'True
                  Caption         =   "Province:"
                  Height          =   195
                  Left            =   3120
                  TabIndex        =   85
                  ToolTipText     =   "Enter your Full Name here for Identification"
                  Top             =   480
                  Width           =   675
               End
               Begin VB.Label lblCity 
                  AutoSize        =   -1  'True
                  Caption         =   "City:"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   82
                  ToolTipText     =   "Enter your Student ID here for Identification"
                  Top             =   960
                  Width           =   300
               End
               Begin VB.Label lblStreeMain 
                  AutoSize        =   -1  'True
                  Caption         =   "Street:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   81
                  ToolTipText     =   "Enter your Full Name here for Identification"
                  Top             =   480
                  Width           =   465
               End
            End
            Begin VB.TextBox txtURL 
               Height          =   405
               Left            =   7800
               Locked          =   -1  'True
               TabIndex        =   69
               ToolTipText     =   "Enter your URL"
               Top             =   840
               Width           =   2655
            End
            Begin VB.TextBox txtPRStID 
               Height          =   405
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   65
               ToolTipText     =   "Enter your Student ID here for Identification"
               Top             =   840
               Width           =   2055
            End
            Begin VB.TextBox txtPREmail 
               Height          =   405
               Left            =   8160
               Locked          =   -1  'True
               TabIndex        =   64
               ToolTipText     =   "Enter your Full Name here for Identification"
               Top             =   360
               Width           =   2295
            End
            Begin VB.TextBox txtPRPhoneNo 
               Height          =   405
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   63
               ToolTipText     =   "Enter your Full Name here for Identification"
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label lblUrlMain 
               AutoSize        =   -1  'True
               Caption         =   "URL:"
               Height          =   195
               Left            =   7320
               TabIndex        =   70
               ToolTipText     =   "Enter your Student ID here for Identification"
               Top             =   960
               Width           =   375
            End
            Begin VB.Label lblPRStID 
               AutoSize        =   -1  'True
               Caption         =   "Student ID:"
               Height          =   195
               Left            =   3600
               TabIndex        =   68
               ToolTipText     =   "Enter your Student ID here for Identification"
               Top             =   960
               Width           =   810
            End
            Begin VB.Label lblPREmail 
               AutoSize        =   -1  'True
               Caption         =   "Email:"
               Height          =   195
               Left            =   7560
               TabIndex        =   67
               ToolTipText     =   "Enter your Full Name here for Identification"
               Top             =   480
               Width           =   420
            End
            Begin VB.Label lblPRPhoneNo 
               AutoSize        =   -1  'True
               Caption         =   "Phone No:"
               Height          =   195
               Left            =   3600
               TabIndex        =   66
               ToolTipText     =   "Enter your Full Name here for Identification"
               Top             =   480
               Width           =   765
            End
         End
         Begin VB.Frame frmPRCompLocation 
            Height          =   1095
            Left            =   120
            TabIndex        =   6
            Top             =   3120
            Width           =   5175
            Begin VB.ComboBox cmbPRCompID 
               Height          =   315
               ItemData        =   "Frmmain.frx":0670
               Left            =   240
               List            =   "Frmmain.frx":0672
               TabIndex        =   56
               Text            =   "1"
               Top             =   600
               Width           =   2295
            End
            Begin VB.ComboBox cmbPRCompLoc 
               Height          =   315
               ItemData        =   "Frmmain.frx":0674
               Left            =   2640
               List            =   "Frmmain.frx":0676
               TabIndex        =   35
               Text            =   "Room 415"
               Top             =   600
               Width           =   2295
            End
            Begin VB.Label lblPRCompNum 
               AutoSize        =   -1  'True
               Caption         =   "Computer ID Number:"
               Height          =   195
               Left            =   240
               TabIndex        =   36
               ToolTipText     =   "Enter the Computer ID Number"
               Top             =   240
               Width           =   1530
            End
            Begin VB.Label lblPRCompLocation 
               AutoSize        =   -1  'True
               Caption         =   "Where is the computer located?"
               Height          =   195
               Left            =   2640
               TabIndex        =   7
               Top             =   240
               Width           =   2265
            End
         End
         Begin VB.CommandButton cmdPRReport 
            Caption         =   "Report Problem"
            Height          =   375
            Left            =   9360
            TabIndex        =   26
            Top             =   8520
            Width           =   1335
         End
         Begin VB.CommandButton cmdPRClearAll 
            Caption         =   "Clear All"
            Height          =   375
            Left            =   7920
            TabIndex        =   25
            Top             =   8520
            Width           =   1335
         End
         Begin VB.Frame frmPRProblem 
            Height          =   2055
            Left            =   120
            TabIndex        =   3
            Top             =   6360
            Width           =   10575
            Begin VB.TextBox txtPRProblem 
               Height          =   1575
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   5
               ToolTipText     =   "Enter a brief description of the problem in your own words."
               Top             =   360
               Width           =   10335
            End
            Begin VB.Label lblPRProblem 
               AutoSize        =   -1  'True
               Caption         =   "Enter a brief description of the problem in your own words."
               Height          =   195
               Left            =   120
               TabIndex        =   4
               ToolTipText     =   "Enter a brief description of the problem in your own words."
               Top             =   120
               Width           =   4080
            End
         End
         Begin VB.Frame frmPRPriority 
            Height          =   1095
            Left            =   5400
            TabIndex        =   37
            Top             =   3120
            Width           =   1575
            Begin VB.ComboBox cmbPRPriority 
               Height          =   315
               ItemData        =   "Frmmain.frx":0678
               Left            =   240
               List            =   "Frmmain.frx":0685
               TabIndex        =   39
               Text            =   "Low"
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label lblPRPriority 
               AutoSize        =   -1  'True
               Caption         =   "Priority Level:"
               Height          =   195
               Left            =   240
               TabIndex        =   38
               ToolTipText     =   "Enter your Full Name here for Identification"
               Top             =   240
               Width           =   945
            End
         End
         Begin VB.Frame frmCampus 
            Height          =   1095
            Left            =   7080
            TabIndex        =   90
            Top             =   3120
            Width           =   1575
            Begin VB.ComboBox cmbCampus 
               Height          =   315
               ItemData        =   "Frmmain.frx":069C
               Left            =   120
               List            =   "Frmmain.frx":069E
               TabIndex        =   91
               Text            =   "CASA LOMA"
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label lblCampus 
               AutoSize        =   -1  'True
               Caption         =   "Campus Area:"
               Height          =   195
               Left            =   120
               TabIndex        =   92
               ToolTipText     =   "Enter your Full Name here for Identification"
               Top             =   240
               Width           =   990
            End
         End
         Begin VB.Frame frmTimeDate 
            Height          =   1095
            Left            =   8760
            TabIndex        =   59
            Top             =   3120
            Width           =   1935
            Begin VB.TextBox txtPRTimeReq 
               Height          =   375
               Left            =   240
               Locked          =   -1  'True
               TabIndex        =   60
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label lblPRTimeReq 
               AutoSize        =   -1  'True
               Caption         =   "Date/Time Requested"
               Height          =   195
               Left            =   240
               TabIndex        =   61
               Top             =   240
               Width           =   1590
            End
         End
         Begin VB.Frame frmPROSandSoftware 
            Height          =   2175
            Left            =   7800
            TabIndex        =   20
            Top             =   4200
            Width           =   2895
            Begin VB.ComboBox cmbPRSoftware 
               Height          =   315
               ItemData        =   "Frmmain.frx":06A0
               Left            =   240
               List            =   "Frmmain.frx":06F2
               TabIndex        =   24
               Text            =   "Notepad"
               Top             =   1440
               Width           =   2415
            End
            Begin VB.ComboBox cmbPROS 
               Height          =   315
               ItemData        =   "Frmmain.frx":0868
               Left            =   240
               List            =   "Frmmain.frx":089F
               TabIndex        =   22
               Text            =   "Dos all versions"
               Top             =   600
               Width           =   2415
            End
            Begin VB.Label lblPRSoftware 
               Caption         =   "What software is faulty?"
               Height          =   255
               Left            =   240
               TabIndex        =   23
               Top             =   1080
               Width           =   2175
            End
            Begin VB.Label lblPROS 
               AutoSize        =   -1  'True
               Caption         =   "What OS is the computer using?"
               Height          =   195
               Left            =   240
               TabIndex        =   21
               Top             =   240
               Width           =   2295
            End
         End
         Begin VB.Frame frmPRHardware 
            Height          =   2175
            Left            =   120
            TabIndex        =   8
            Top             =   4200
            Width           =   7575
            Begin VB.OptionButton optPRScanner 
               Caption         =   "Scanner"
               Height          =   255
               Left            =   2640
               TabIndex        =   58
               Top             =   1080
               Width           =   1095
            End
            Begin VB.OptionButton optPRPrinter 
               Caption         =   "Printer"
               Height          =   255
               Left            =   1680
               TabIndex        =   57
               Top             =   1080
               Width           =   975
            End
            Begin VB.TextBox txtPROthers 
               Height          =   375
               Left            =   3840
               TabIndex        =   19
               Top             =   1440
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.OptionButton optPROthers 
               Caption         =   "Others, Please specify:"
               Height          =   255
               Left            =   3840
               TabIndex        =   18
               Top             =   1080
               Width           =   2055
            End
            Begin VB.OptionButton optPRSoundCard 
               Caption         =   "Sound Card"
               Height          =   255
               Left            =   6000
               TabIndex        =   17
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton optPRVideoCard 
               Caption         =   "Video Card"
               Height          =   255
               Left            =   240
               TabIndex        =   16
               Top             =   1080
               Width           =   1095
            End
            Begin VB.OptionButton optPRCDRom 
               Caption         =   "CD-ROM"
               Height          =   255
               Left            =   4200
               TabIndex        =   15
               Top             =   600
               Width           =   1095
            End
            Begin VB.OptionButton optPRRAM 
               Caption         =   "RAM"
               Height          =   255
               Left            =   5280
               TabIndex        =   14
               Top             =   600
               Width           =   855
            End
            Begin VB.OptionButton optPRMouse 
               Caption         =   "Mouse"
               Height          =   255
               Left            =   3240
               TabIndex        =   13
               Top             =   600
               Width           =   1095
            End
            Begin VB.OptionButton optPRKeyboard 
               Caption         =   "Keyboard"
               Height          =   255
               Left            =   2040
               TabIndex        =   12
               Top             =   600
               Width           =   1335
            End
            Begin VB.OptionButton optPRCPU 
               Caption         =   "CPU"
               Height          =   255
               Left            =   1200
               TabIndex        =   11
               Top             =   600
               Width           =   975
            End
            Begin VB.OptionButton optPRMonitor 
               Caption         =   "Monitor"
               Height          =   255
               Left            =   240
               TabIndex        =   10
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label lblPRHardware 
               AutoSize        =   -1  'True
               Caption         =   "What hardware is faulty? ie. RAM, Monitor, CPU, Keyboard, Mouse"
               Height          =   195
               Left            =   240
               TabIndex        =   9
               Top             =   240
               Width           =   4740
            End
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuBackUp 
         Caption         =   "&Back Up"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      WindowList      =   -1  'True
      Begin VB.Menu mnuSndMail 
         Caption         =   "&Send Mail"
      End
      Begin VB.Menu mnuTipOfTheDay 
         Caption         =   "&Tip of the day"
      End
      Begin VB.Menu mnuNewUser 
         Caption         =   "&New User"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' declare all variable explicitly

' get OS help from user32 dll
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub cmbAFCCP_Click()
    
    Dim IntNum As Integer
    
    IntNum = cmbAFCCP.List(cmbAFCCP.ListIndex)
    
    Set rstProb = db.OpenRecordset("SELECT * FROM tblProb")
        
    Do Until rstProb.EOF
        If (rstProb!PROB_ID = IntNum) Then
            txtCompletedByTech.Text = rstProb!TECH_NAME
        End If
        rstProb.MoveNext
    Loop

End Sub

Private Sub cmbNPAllProblems_Click()

    Dim IntNum As Integer
    
    frmNPFixEdit.Visible = True
    
    IntNum = cmbNPAllProblems.List(cmbNPAllProblems.ListIndex)
    
    Set rstProb = db.OpenRecordset("SELECT * FROM tblProb")
        
    Do Until rstProb.EOF
        If (rstProb!PROB_ID = IntNum) Then
            
            If (rstProb!PROB_STATUS = "A") Then
                txtNPPStatus.Text = "Active"
            End If
            
            If (rstProb!PROB_PRIORITY = "L") Then
                txtNPPPriority.Text = "Low"
            ElseIf (rstProb!PROB_PRIORITY = "M") Then
                txtNPPPriority.Text = "Medium"
            Else
                txtNPPPriority.Text = "High"
                MsgBox "Priority Level is HIGH", vbInformation, "PRIORITY LEVEL"
            End If
                        
            txtNPDTReq.Text = rstProb!PROB_REQBYDATE ' req. date
            Set rstUserRT = db.OpenRecordset("SELECT * FROM tblUser")
            
            Do Until rstUserRT.EOF
                If (rstUserRT!USER_ID = rstProb!USER_ID) Then
                    txtReqBy.Text = rstUserRT!USER_LOGIN
                End If
                rstUserRT.MoveNext
            Loop
            
            txtCompID.Text = rstProb!COMP_ID
            
            Set rstCampusRT = db.OpenRecordset("SELECT * FROM tblCampus")
            
            Do Until rstCampusRT.EOF
                If (rstCampusRT!CAMPUS_ID = rstProb!CAMPUS_ID) Then
                    txtCampusFix.Text = rstCampusRT!CAMPUS_NAME
                End If
                rstCampusRT.MoveNext
            Loop
            
            txtCompLoc.Text = rstProb!ROOM
            
            txtSoftware.Text = rstProb!PROB_SOFTWARE
            txtHardware.Text = rstProb!PROB_HARDWARE
            txtOS.Text = rstProb!PROB_OS
            txtDOUP.Text = rstProb!PROB_DESC
        End If
        rstProb.MoveNext
    Loop

End Sub

' This command button will cancel your print. Tech/Admin/user access
Private Sub cmbPRepCancel_Click()
    
    If MsgBox("Are you sure you want to clear this current report?", vbYesNo, "Control info") = vbYes Then
        rtbPRView.Text = ""
    Else
        MsgBox "Select a new problem by " & vbCrLf & _
        "by clikcing on the Combo Box", vbInformation, "Control info"
    End If
    
End Sub

' This command button will give you a chance to make some changes to the data before you print it. Tech/Admin/user access
Private Sub cmbPRepPreview_Click()

        Load frmPrinter
        frmPrinter.cmdClose.Visible = True
        frmPrinter.rtbPRView.Text = rtbPRView.Text
        frmPrinter.Show 1

End Sub

' This command button will print the current problem data. Tech/Admin/user access
Private Sub cmbPRepPrint_Click()
    
On Error GoTo PrinterError
    
    If MsgBox("Are you sure you want to print this report?", vbYesNo, "Print") = vbYes Then
    
        Load frmPrinter
        frmPrinter.rtbPRView.Text = rtbPRView.Text
        frmPrinter.PrintForm
        Unload frmPrinter
        MsgBox "Report has been printed"
        
    Else
        MsgBox "Please review the selected report, Printed has been cancelled"
    End If
    
ProcExit:
    Screen.MousePointer = vbDefault ' set mouse pointer to vbdefault
    Exit Sub ' exit sub
PrinterError:
    MsgBox (Err.Description), , "Printer Error"
    Resume ProcExit

End Sub

Private Sub cmbPRepPrintSel_Click()

    Dim IntNum As Integer
    
    IntNum = cmbPRepPrintSel.List(cmbPRepPrintSel.ListIndex)
    
    Set rstProb = db.OpenRecordset("SELECT * FROM tblProb")
        
    rtbPRView.AutoVerbMenu = True
    rtbPRView.SetFocus
    
    Do Until rstProb.EOF
        If (rstProb!PROB_ID = IntNum) Then
            rtbPRView.Text = "Problem ID: " & rstProb!PROB_ID & vbCrLf & _
            "Requested By: " & rstProb!USER_ID & vbCrLf & _
            "Campus Location: " & rstProb!CAMPUS_ID & vbCrLf & _
            "Room: " & rstProb!ROOM & vbCrLf & _
            "Tech Responsible: " & rstProb!TECH_NAME & vbCrLf & _
            "Computer Id: " & rstProb!COMP_ID & vbCrLf & _
            "Hardware Problem: " & rstProb!PROB_HARDWARE & vbCrLf & _
            "Using Operating System: " & rstProb!PROB_OS
            
            If (rstProb!PROB_STATUS = "A") Then
                rtbPRView.Text = rtbPRView.Text & vbCrLf & _
                "Status: is Active"
            Else
                rtbPRView.Text = rtbPRView.Text & vbCrLf & _
                "Status: is Not Active"
            End If
            
            rtbPRView.Text = rtbPRView.Text & vbCrLf & _
            "Problem Description: " & rstProb!PROB_DESC
            
            If (rstProb!PROB_PRIORITY = "L") Then
                rtbPRView.Text = rtbPRView.Text & vbCrLf & _
                "Priority Level: Low"
            ElseIf (rstProb!PROB_PRIORITY = "M") Then
                rtbPRView.Text = rtbPRView.Text & vbCrLf & _
                "Priority Level: Medium"
            Else
                rtbPRView.Text = rtbPRView.Text & vbCrLf & _
                "Priority Level: High"
                MsgBox "Priority Level is HIGH", vbInformation, "PRIORITY LEVEL"
            End If
            
            rtbPRView.Text = rtbPRView.Text & vbCrLf & _
            "Problem Requested Date: " & rstProb!PROB_REQBYDATE
            
            If (rstProb!PROB_RESPSTATUS = False) Then
                rtbPRView.Text = rtbPRView.Text & vbCrLf & _
                "Response Status: No"
            Else
                rtbPRView.Text = rtbPRView.Text & vbCrLf & _
                "Response Status: Yes"
            End If
              
            rtbPRView.Text = rtbPRView.Text & vbCrLf & _
            "Problem Response Date: " & rstProb!PROB_RESPDATE & vbCrLf & _
            "Problem Closed Date: " & rstProb!PROB_DATECLOSED & vbCrLf & _
            "Fix Notes: " & rstProb!PROB_FIXEDNOTES
            
        End If
        rstProb.MoveNext
    Loop

End Sub

' This button setups the printer that you want to use. Tech/Admin/user access
Private Sub cmbPRepPSetup_Click()
    
    CD.ShowPrinter
    
End Sub

Private Sub cmbPROS_Change()

    ' Lets set the default value to Dos all versions if the user
    ' Plays around with our control
    cmbPROS.Text = "Dos all versions"

End Sub

Private Sub cmbPRCompLoc_Change()
    
    ' Lets set the default value to Open Access Labs if the user
    ' Plays around with our control
    cmbPRCompLoc.Text = "Open Access Labs"
    
End Sub

Private Sub cmbPRPriority_Change()

    ' Lets set the default value to Low if the user
    ' Plays around with our control
    cmbPRPriority.Text = "Low"
    
End Sub

Private Sub cmbPRSoftware_Change()
    
    ' Lets set the default value to Notepad if the user
    ' Plays around with our control
    cmbPRSoftware.Text = "Notepad"

End Sub

' Exit
Private Sub cmdExit_Click()

    End
    
End Sub

Private Sub cmdFindAdmin_Click()

    Dim i As Integer
    Dim j As Integer
    
    'Select entire grid and remove bold formatting
    '(to remove the results of previous finds)
    MFGPrint.FillStyle = flexFillRepeat
    MFGAdmin.Col = 0
    MFGAdmin.Row = 0
    MFGAdmin.ColSel = MFGAdmin.Cols - 1
    MFGAdmin.RowSel = MFGAdmin.Rows - 1
    MFGAdmin.CellFontBold = False
    
    'Initialize ProgressBar to track search
    pBarAdmin.Min = 0
    pBarAdmin.Max = MFGAdmin.Rows - 1
    pBarAdmin.Visible = True
    
    'Search the grid cell by cell for find text
    MFGAdmin.FillStyle = flexFillSingle
    For i = 0 To MFGAdmin.Cols - 1
        For j = 1 To MFGAdmin.Rows - 1
        'Display current row location on ProgressBar
        pBarAdmin.Value = j
            'If current cell matches find text box
            If InStr(MFGAdmin.TextMatrix(j, i), _
            txtSearch.Text) Then
                '...select cell and format bold
                MFGAdmin.Col = i
                MFGAdmin.Row = j
                MFGAdmin.CellFontBold = True
            End If
        Next j
    Next i
    pBarAdmin.Visible = False 'hide ProgressBar
    
End Sub

Private Sub cmdFindTxt_Click()
    
    Dim i As Integer
    Dim j As Integer
    
    'Select entire grid and remove bold formatting
    '(to remove the results of previous finds)
    MFGPrint.FillStyle = flexFillRepeat
    MFGPrint.Col = 0
    MFGPrint.Row = 0
    MFGPrint.ColSel = MFGPrint.Cols - 1
    MFGPrint.RowSel = MFGPrint.Rows - 1
    MFGPrint.CellFontBold = False
    
    'Initialize ProgressBar to track search
    pBarPrint.Min = 0
    pBarPrint.Max = MFGPrint.Rows - 1
    pBarPrint.Visible = True
    
    'Search the grid cell by cell for find text
    MFGPrint.FillStyle = flexFillSingle
    For i = 0 To MFGPrint.Cols - 1
        For j = 1 To MFGPrint.Rows - 1
        'Display current row location on ProgressBar
        pBarPrint.Value = j
            'If current cell matches find text box
            If InStr(MFGPrint.TextMatrix(j, i), _
            txtSearch.Text) Then
                '...select cell and format bold
                MFGPrint.Col = i
                MFGPrint.Row = j
                MFGPrint.CellFontBold = True
            End If
        Next j
    Next i
    pBarPrint.Visible = False 'hide ProgressBar
    
End Sub

' Create a new user
Private Sub cmdNewUser_Click()

    Load frmNewUser
    frmNewUser.Show 1
    
End Sub

' This command button will let you fix the current problem that you have selected. Tech/Admin only
Private Sub cmdNPFixEdit_Click()
    
    Set rstProb = db.OpenRecordset("SELECT * FROM tblProb")
    
    Do Until rstProb.EOF
        If (cmbNPAllProblems.List(cmbNPAllProblems.ListIndex) = rstProb!PROB_ID) Then
            
            With rstProb
                rstProb.Edit
                rstProb!TECH_NAME = Login
                rstProb!PROB_STATUS = "F"
                rstProb!PROB_RESPSTATUS = True
                rstProb!PROB_RESPDATE = Now()
                rstProb!PROB_DATECLOSED = Now()
                rstProb!PROB_FIXEDNOTES = RTBFixes.Text
                rstProb.Update
            End With
        
        End If
        rstProb.MoveNext
    Loop
        
    MsgBox "Problem Fixed.", vbInformation, "Problems"
    
    frmNPFixEdit.Visible = False
    cmbNPAllProblems.Text = ""
    txtNPPStatus.Text = ""
    txtNPPPriority.Text = ""
    txtNPDTReq.Text = ""
    RTBFixes.Text = ""
    
    Set rstProb = db.OpenRecordset("SELECT * FROM tblProb")
        
    cmbNPAllProblems.Clear
    
    Do Until rstProb.EOF
        If (rstProb!PROB_STATUS = "A") Then
            If (cmbNPAllProblems.Text = "") Then
                cmbNPAllProblems.Text = rstProb!PROB_ID
            End If
            cmbNPAllProblems.AddItem rstProb!PROB_ID
        End If
        rstProb.MoveNext
    Loop
    
    dataDAO.Refresh
            
End Sub

' This command button function will erase all the contents of the current problem if the user decides to clear it. Tech/Admin/user access
Private Sub cmdPRClearAll_Click()

    If MsgBox("Are you sure you want to delete this current report?", vbYesNo, "Clear Report") = vbYes Then
        txtPROthers.Text = ""
        txtPROthers.Visible = False
        optPRMonitor.Value = True
        cmbPRCompID.Text = "1"
        cmbPRCompLoc.Text = "Room 415"
        cmbPRPriority.Text = "Low"
        cmbCampus.Text = "CASA LOMA"
        txtPRTimeReq.Text = Now
        cmbPROS.Text = "Dos all versions"
        cmbPRSoftware.Text = "Notepad"
        txtPRProblem.Text = ""
    Else
        MsgBox "Please review your report!", vbInformation, "Clear Report"
    End If

End Sub

' This command button will send the current problem to the database. Tech/Admin/user access
Private Sub cmdPRReport_Click()
        
    Set rstProb = db.OpenRecordset("Select * FROM tblProb")
    Set rstCampus = db.OpenRecordset("Select * FROM tblCampus WHERE CAMPUS_NAME = '" & cmbCampus.Text & "'")
    
    Do Until rstProb.EOF
        rstProb.MoveLast
        If (rstProb!PROB_ID <> 0) Then
        
            probNum = rstProb!PROB_ID + 1
        
        Else
            
            probNum = 0 + 1
        
        End If
        rstProb.MoveNext
    Loop
    
    If MsgBox("Are you sure you want to file this report?", vbYesNo, "Save Report") = vbYes Then
                
        With rstProb
            
            .AddNew
            !PROB_ID = probNum
            !USER_ID = UserId
            !TECH_NAME = ""
            !ROOM = cmbPRCompLoc.Text
            
            Do Until rstCampus.EOF
                !CAMPUS_ID = rstCampus!CAMPUS_ID
                rstCampus.MoveNext
            Loop
            
            !COMP_ID = cmbPRCompID.Text
            
            If (optPRMonitor.Value = True) Then
                !PROB_HARDWARE = "Monitor"
            ElseIf (optPRCPU.Value = True) Then
                !PROB_HARDWARE = "CPU"
            ElseIf (optPRKeyboard.Value = True) Then
                !PROB_HARDWARE = "Keyboard"
            ElseIf (optPRMouse.Value = True) Then
                !PROB_HARDWARE = "Mouse"
            ElseIf (optPRCDRom.Value = True) Then
                !PROB_HARDWARE = "CD-ROM"
            ElseIf (optPRRAM.Value = True) Then
                !PROB_HARDWARE = "RAM"
            ElseIf (optPRSoundCard.Value = True) Then
                !PROB_HARDWARE = "Sound Card"
            ElseIf (optPRVideoCard.Value = True) Then
                !PROB_HARDWARE = "Video Card"
            ElseIf (optPRPrinter.Value = True) Then
                !PROB_HARDWARE = "Printer"
            ElseIf (optPRScanner.Value = True) Then
                !PROB_HARDWARE = "Scanner"
            Else
                !PROB_HARDWARE = txtPROthers.Text
            End If
            
            !PROB_OS = cmbPROS.Text
            !PROB_SOFTWARE = cmbPRSoftware.Text
            !PROB_STATUS = "A"
            !PROB_DESC = txtPRProblem.Text
                        
            If (cmbPRPriority.Text = "Low") Then
                !PROB_PRIORITY = "L"
            ElseIf (cmbPRPriority.Text = "Medium") Then
                !PROB_PRIORITY = "M"
            Else
                !PROB_PRIORITY = "H"
            End If
                        
            !PROB_REQBYDATE = txtPRTimeReq.Text
            !PROB_RESPSTATUS = False
            
            .Update
        
        End With
        
        txtPROthers.Text = ""
        txtPROthers.Visible = False
        optPRMonitor.Value = True
        cmbPRCompID.Text = "1"
        cmbPRCompLoc.Text = "Room 415"
        cmbPRPriority.Text = "Low"
        cmbCampus.Text = "CASA LOMA"
        txtPRTimeReq.Text = Now
        cmbPROS.Text = "Dos all versions"
        cmbPRSoftware.Text = "Notepad"
        txtPRProblem.Text = ""
        
        MsgBox "Report has been successfully saved!" & vbCrLf & "Please write this problem # " & probNum & " down." & vbCrLf & "For future reference.", vbInformation, "Report Saved"
    
    Else
        
        MsgBox "Please review your report!", vbInformation, "Save Report"
    
    End If
    
End Sub

Private Sub cmdSort_Click()
    
    'Set column 2 (LastName) as the sort key
    MFGPrint.Col = 2
    'Sort grid in ascending order
    MFGPrint.Sort = 1
    
End Sub

Private Sub cmdSortAdmin_Click()
    
    'Set column 2 (LastName) as the sort key
    MFGAdmin.Col = 2
    'Sort grid in ascending order
    MFGAdmin.Sort = 1
    
End Sub


Private Sub Form_Load()
    
    ' set the stContainer to visible
    stContainer.Visible = True
    
    ' check if access is U enable and disable tabs
    If Access = "U" Then
        stContainer.TabVisible(0) = True
        stContainer.TabVisible(1) = True
        stContainer.TabVisible(2) = True
        stContainer.TabVisible(3) = False
        stContainer.TabVisible(4) = False
    End If
    ' check if access is T then enable and disable tabs
    If Access = "T" Then
        stContainer.TabVisible(0) = True
        stContainer.TabVisible(1) = True
        stContainer.TabVisible(2) = True
        stContainer.TabVisible(3) = True
        stContainer.TabVisible(4) = False
    End If
    ' check if access is A then enable all tabs
    If Access = "A" Then
        stContainer.TabVisible(0) = True
        stContainer.TabVisible(1) = True
        stContainer.TabVisible(2) = True
        stContainer.TabVisible(3) = True
        stContainer.TabVisible(4) = True
        mnuNewUser.Visible = True
    End If
    
    txtPRTimeReq.Text = Now()
    'Date & " " & Time
    
    Set rstComputer = db.OpenRecordset("SELECT * FROM tblComp")
    Set rstRoom = db.OpenRecordset("SELECT * FROM tblRoom")
    Set rstUserSelected = db.OpenRecordset("SELECT * FROM tblUser WHERE USER_LOGIN = '" & Login & "'")
    Set rstCampus = db.OpenRecordset("SELECT * FROM tblCampus")
    
    Do Until rstComputer.EOF
        cmbPRCompID.AddItem rstComputer!COMP_ID
        rstComputer.MoveNext
    Loop
       
    Do Until rstRoom.EOF
        cmbPRCompLoc.AddItem rstRoom!ROOM_NAME
        rstRoom.MoveNext
    Loop
       
    Do Until rstCampus.EOF
        cmbCampus.AddItem rstCampus!CAMPUS_NAME
        rstCampus.MoveNext
    Loop
        
    optPRMonitor.Value = True
    
    Do Until rstUserSelected.EOF
        txtPRStID.Text = rstUserSelected!USER_STUDENTID
        txtLastName.Text = rstUserSelected!USER_LASTNAME
        txtFirstName.Text = rstUserSelected!USER_FIRSTNAME
        txtPRPhoneNo.Text = rstUserSelected!USER_PHONE
        txtPREmail.Text = rstUserSelected!USER_EMAIL
        txtURL.Text = rstUserSelected!USER_URL
        txtStreet.Text = rstUserSelected!USER_STREET
        txtCity.Text = rstUserSelected!USER_CITY
        txtProvince.Text = rstUserSelected!USER_PROVINCE
        txtPostalCode.Text = rstUserSelected!USER_POSTAL
        txtPassword.Text = rstUserSelected!USER_PASSWORD
        UserId = rstUserSelected!USER_ID
        rstUserSelected.MoveNext
    Loop

End Sub

' Unload this form
Private Sub Form_Unload(Cancel As Integer)
    
    Dim i As Integer
    
    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    
    ' End the program
    End
    
End Sub

Private Sub mnuBackUp_Click()

On Error GoTo ProcError

Dim prompt As String
Dim FileName As String

    prompt = "Would you like to create a backup copy of the database?"
    If MsgBox(prompt, vbOKCancel, "db.DatabaseName") = vbOK Then  'copy the database if user clicks OK
        FileName = InputBox("Enter the complete path" & vbCrLf & _
                            "plus the filename of the database " & vbCrLf & _
                            "for the backup copy." & vbCrLf & _
                            "Ex. c:\temp.mdb")
        If FileName <> "" Then _
            FileCopy App.Path & "\" & "HelpDesk.mdb", FileName
    End If

ProcExit:
    Screen.MousePointer = vbDefault ' set mouse pointer to vbdefault
    Exit Sub ' exit sub
ProcError:
    MsgBox Err.Description, vbExclamation ' tell user about the error
    Resume ProcExit ' resume with out quitting
    
End Sub

' show the about form
Private Sub mnuHelpAbout_Click()
    
    frmAbout.Show vbModal, Me
    
End Sub


Private Sub mnuHelpSearchForHelpOn_Click()
    
    Dim nRet As Integer

    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
    
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    
    Else
        
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        
        If Err Then
            
            MsgBox Err.Description
        
        End If
    
    End If

End Sub

Private Sub mnuHelpContents_Click()
    
    Dim nRet As Integer

    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuNewUser_Click()
    
    Load frmNewUser
    frmNewUser.Show 1
    
End Sub

' load the email form
Private Sub mnuSndMail_Click()

    Load frmMail
    frmMail.Show 1
    
End Sub

Private Sub mnuTipOfTheDay_Click()

    Load frmTip
    frmTip.Show 1
    
End Sub

' load the options form
' this form has the tips on/off switch
Private Sub mnuToolsOptions_Click()
    
    Load frmOptions
    frmOptions.Show 1
   
End Sub

' close the program
Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me
    End

End Sub

Private Sub optPRCDRom_Click()
    
    txtPROthers.Visible = False
    
End Sub

Private Sub optPRCPU_Click()
    
    txtPROthers.Visible = False
    
End Sub

Private Sub optPRKeyboard_Click()
    
    txtPROthers.Visible = False
    
End Sub

Private Sub optPRMonitor_Click()
    
    txtPROthers.Visible = False
    
End Sub

Private Sub optPRMouse_Click()
    
    txtPROthers.Visible = False
    
End Sub

Private Sub optPROthers_Click()

    txtPROthers.Visible = True
    
End Sub

Private Sub optPRPrinter_Click()
    
    txtPROthers.Visible = False

End Sub

Private Sub optPRRAM_Click()
    
    txtPROthers.Visible = False
    
End Sub

Private Sub optPRScanner_Click()
    
    txtPROthers.Visible = False

End Sub

Private Sub optPRSoundCard_Click()
    
    txtPROthers.Visible = False
    
End Sub

Private Sub optPRVideoCard_Click()
    
    txtPROthers.Visible = False
    
End Sub

Private Sub stContainer_Click(PreviousTab As Integer)
        
    Set rstProb = db.OpenRecordset("SELECT * FROM tblProb")
        
    cmbPRepPrintSel.Clear
    cmbPRepPrintSel.Text = "1"
    
    Do Until rstProb.EOF
        cmbPRepPrintSel.AddItem rstProb!PROB_ID
        rstProb.MoveNext
    Loop
    
    Set rstProb = db.OpenRecordset("SELECT * FROM tblProb")
        
    cmbNPAllProblems.Clear
    
    Do Until rstProb.EOF
        If (rstProb!PROB_STATUS = "A") Then
            If (cmbNPAllProblems.Text = "") Then
                cmbNPAllProblems.Text = rstProb!PROB_ID
            End If
            cmbNPAllProblems.AddItem rstProb!PROB_ID
        End If
        rstProb.MoveNext
    Loop
    
    Set rstProb = db.OpenRecordset("SELECT * FROM tblProb")
        
    cmbAFCCP.Clear
    
    Do Until rstProb.EOF
        If (rstProb!PROB_STATUS = "F") Then
            If (cmbAFCCP.Text = "") Then
                cmbAFCCP.Text = rstProb!PROB_ID
            End If
            cmbAFCCP.AddItem rstProb!PROB_ID
        End If
        rstProb.MoveNext
    Loop
    
    dataDAO.Refresh
    
End Sub
