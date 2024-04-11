VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   4170
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComctlLib.ProgressBar pbSplash 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "loading............"
      Top             =   3600
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
   Begin VB.Image imgGBC 
      Height          =   1785
      Left            =   4080
      Picture         =   "frmSplash.frx":4B146
      Top             =   240
      Width           =   3990
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Height          =   675
      Left            =   120
      TabIndex        =   3
      Tag             =   "Warning"
      Top             =   2880
      Width           =   8055
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Tag             =   "Version"
      Top             =   1320
      Width           =   930
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jeff Software"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      TabIndex        =   1
      Tag             =   "CompanyProduct"
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   900
      Left            =   480
      TabIndex        =   0
      Tag             =   "Product"
      Top             =   1920
      Width           =   3480
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' Explicitly define variables

' set the caption of lblVersion to application version
' set the caption of lblProductName to the program name
Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    
End Sub

