VERSION 5.00
Begin VB.Form frmNewUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New User/Tech/Admin"
   ClientHeight    =   4650
   ClientLeft      =   4335
   ClientTop       =   5925
   ClientWidth     =   10635
   Icon            =   "frmNewUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   10635
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9240
      TabIndex        =   18
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Frame frmUserItems 
      Caption         =   "Add New Users / Techs / Admins"
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.OptionButton optTech 
         Caption         =   "Technician"
         Height          =   375
         Left            =   4800
         TabIndex        =   15
         Top             =   3000
         Width           =   1455
      End
      Begin VB.OptionButton optUser 
         Caption         =   "User"
         Height          =   375
         Left            =   2880
         TabIndex        =   14
         Top             =   3000
         Width           =   1335
      End
      Begin VB.OptionButton optAdmin 
         Caption         =   "Administrator"
         Height          =   375
         Left            =   6840
         TabIndex        =   16
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox txtPhoneNo 
         Height          =   405
         Left            =   4560
         TabIndex        =   5
         ToolTipText     =   "Enter your Full Name here for Identification"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtEmail 
         Height          =   405
         Left            =   8160
         TabIndex        =   7
         ToolTipText     =   "Enter your Full Name here for Identification"
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtStID 
         Height          =   405
         Left            =   4560
         TabIndex        =   6
         ToolTipText     =   "Enter your Student ID here for Identification"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtURL 
         Height          =   405
         Left            =   7800
         TabIndex        =   8
         ToolTipText     =   "Enter your URL"
         Top             =   840
         Width           =   2655
      End
      Begin VB.Frame frmAddress 
         Height          =   1455
         Left            =   3960
         TabIndex        =   24
         Top             =   1440
         Width           =   6255
         Begin VB.TextBox txtStreet 
            Height          =   405
            Left            =   720
            TabIndex        =   9
            ToolTipText     =   "Enter your Full Name here for Identification"
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtCity 
            Height          =   405
            Left            =   720
            TabIndex        =   10
            ToolTipText     =   "Enter your Student ID here for Identification"
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox txtProvince 
            Height          =   405
            Left            =   3960
            TabIndex        =   11
            ToolTipText     =   "Enter your Full Name here for Identification"
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtPostalCode 
            Height          =   405
            Left            =   3960
            TabIndex        =   12
            ToolTipText     =   "Enter your Student ID here for Identification"
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Street:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            ToolTipText     =   "Enter your Full Name here for Identification"
            Top             =   480
            Width           =   465
         End
         Begin VB.Label lblCity 
            AutoSize        =   -1  'True
            Caption         =   "City:"
            Height          =   195
            Left            =   240
            TabIndex        =   27
            ToolTipText     =   "Enter your Student ID here for Identification"
            Top             =   960
            Width           =   300
         End
         Begin VB.Label lblProvince 
            AutoSize        =   -1  'True
            Caption         =   "Province:"
            Height          =   195
            Left            =   3120
            TabIndex        =   26
            ToolTipText     =   "Enter your Full Name here for Identification"
            Top             =   480
            Width           =   675
         End
         Begin VB.Label lblPostalCode 
            AutoSize        =   -1  'True
            Caption         =   "Postal Code:"
            Height          =   195
            Left            =   2880
            TabIndex        =   25
            ToolTipText     =   "Enter your Student ID here for Identification"
            Top             =   960
            Width           =   900
         End
      End
      Begin VB.Frame frmInfo 
         Height          =   2655
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3255
         Begin VB.TextBox txtLogin 
            Height          =   405
            Left            =   1200
            TabIndex        =   1
            ToolTipText     =   "Enter your Full Name here for Identification"
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtLastName 
            Height          =   405
            Left            =   1200
            TabIndex        =   3
            ToolTipText     =   "Enter your Last Name here for Identification"
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox txtFirstName 
            Height          =   405
            Left            =   1200
            TabIndex        =   4
            ToolTipText     =   "Enter your First Name here for Identification"
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtPassword 
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   1200
            PasswordChar    =   "*"
            TabIndex        =   2
            ToolTipText     =   "Enter your Full Name here for Identification"
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblPRLogin 
            AutoSize        =   -1  'True
            Caption         =   "Login:"
            Height          =   195
            Left            =   600
            TabIndex        =   23
            ToolTipText     =   "Enter your Full Name here for Identification"
            Top             =   600
            Width           =   435
         End
         Begin VB.Label lblLastName 
            AutoSize        =   -1  'True
            Caption         =   "Last Name:"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            ToolTipText     =   "Enter your Full Name here for Identification"
            Top             =   1560
            Width           =   810
         End
         Begin VB.Label lblFirstName 
            AutoSize        =   -1  'True
            Caption         =   "First Name:"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            ToolTipText     =   "Enter your Full Name here for Identification"
            Top             =   2040
            Width           =   795
         End
         Begin VB.Label lblPassword 
            AutoSize        =   -1  'True
            Caption         =   "Password:"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            ToolTipText     =   "Enter your Full Name here for Identification"
            Top             =   1080
            Width           =   735
         End
      End
      Begin VB.Label lblPRPhoneNo 
         AutoSize        =   -1  'True
         Caption         =   "Phone No:"
         Height          =   195
         Left            =   3600
         TabIndex        =   32
         ToolTipText     =   "Enter your Full Name here for Identification"
         Top             =   480
         Width           =   765
      End
      Begin VB.Label lblPREmail 
         AutoSize        =   -1  'True
         Caption         =   "Email:"
         Height          =   195
         Left            =   7560
         TabIndex        =   31
         ToolTipText     =   "Enter your Full Name here for Identification"
         Top             =   480
         Width           =   420
      End
      Begin VB.Label lblPRStID 
         AutoSize        =   -1  'True
         Caption         =   "Student ID:"
         Height          =   195
         Left            =   3600
         TabIndex        =   30
         ToolTipText     =   "Enter your Student ID here for Identification"
         Top             =   960
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "URL:"
         Height          =   195
         Left            =   7320
         TabIndex        =   29
         ToolTipText     =   "Enter your Student ID here for Identification"
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.Frame frmCampus 
      Height          =   1095
      Left            =   0
      TabIndex        =   33
      Top             =   3480
      Width           =   1575
      Begin VB.ComboBox cmbCampus 
         Height          =   315
         ItemData        =   "frmNewUser.frx":0442
         Left            =   120
         List            =   "frmNewUser.frx":0444
         TabIndex        =   13
         Text            =   "CASA LOMA"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblCampus 
         AutoSize        =   -1  'True
         Caption         =   "Campus Area:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Enter your Full Name here for Identification"
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   35
      Top             =   4200
      Width           =   1335
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmbCampus_Click()
    
    If (txtLogin.Text <> "" And txtPassword.Text <> "" And txtLastName.Text <> "" And txtFirstName.Text <> "" _
        And txtPhoneNo.Text <> "" And txtStID.Text <> "" And txtEmail.Text <> "" And txtURL.Text <> "" _
        And txtStreet.Text <> "" And txtCity.Text <> "" And txtProvince.Text <> "" And txtPostalCode.Text <> "" _
        And cmbCampus.Text <> "") Then
        
        cmdSave.Enabled = True ' enable cmbSize
        cmdClose.Enabled = True
    
    End If

End Sub

Private Sub cmdClearAll_Click()

    txtLogin.Text = ""
    txtPassword.Text = ""
    txtLastName.Text = ""
    txtFirstName.Text = ""
    txtPhoneNo.Text = ""
    txtStID.Text = ""
    txtEmail.Text = ""
    txtURL.Text = ""
    txtStreet.Text = ""
    txtCity.Text = ""
    txtProvince.Text = ""
    txtPostalCode.Text = ""
    cmbCampus.Text = "CASA LOMA"
    
    optUser.Value = True
    
    cmdSave.Enabled = False
    
End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdSave_Click()
    
    Dim STID As Long
    
    STID = Mid(txtStID.Text, 1, 9)
    
    Set rstUser = db.OpenRecordset("SELECT * FROM tblUser")
    
    If MsgBox("Do you really want to add this new user?", vbYesNo, "Add User") = vbYes Then
        
        With rstUser
            .AddNew
            
            Select Case cmbCampus
                Case "CASA LOMA"
                    !CAMPUS_ID = 1
                Case "ST JAMES"
                    !CAMPUS_ID = 2
                Case "NIGHTINGALE"
                    !CAMPUS_ID = 3
                Case "HOSPITALITY"
                    !CAMPUS_ID = 4
            End Select
            
            
            !USER_STUDENTID = STID
            !USER_LOGIN = txtLogin.Text
            !USER_PASSWORD = txtPassword.Text
            !USER_STATUS = "A"
            !USER_LASTNAME = txtLastName.Text
            !USER_FIRSTNAME = txtFirstName.Text
            !USER_PHONE = txtPhoneNo.Text
            !USER_EMAIL = txtEmail.Text
            !USER_URL = txtURL.Text
            !USER_STREET = txtStreet.Text
            !USER_CITY = txtCity.Text
            !USER_PROVINCE = txtProvince.Text
            !USER_POSTAL = txtPostalCode.Text
            !USER_COUNTRY = "CANADA"
            
            If (optUser.Value = True) Then
                MsgBox "Adding new User..... Please wait.......", vbInformation, "Adding....."
            ElseIf (optTech.Value = True) Then
                MsgBox "Adding new Tech..... Please wait.......", vbInformation, "Adding....."
                !USER_TECH = True
            ElseIf (optAdmin.Value = True) Then
                MsgBox "Adding new Admin..... Please wait.......", vbInformation, "Adding....."
                !USER_ADMIN = True
            End If
            
            .Update
        End With
    
        MsgBox "Added " & txtLogin.Text & " to the userlist"
        
    Else
        MsgBox "Please review the new user for further changes"
    End If

End Sub

Private Sub Form_Load()

    optUser.Value = True
    
    Set rstCampus = db.OpenRecordset("SELECT * FROM tblCampus")
    
    Do Until rstCampus.EOF
        cmbCampus.AddItem rstCampus!CAMPUS_NAME
        rstCampus.MoveNext
    Loop
    
    If Access = "U" Then
        optAdmin.Visible = False
        optTech.Visible = False
        txtLogin.Text = Login
        txtLogin.Locked = True
        Me.BorderStyle = vbBSNone
        cmdClose.Enabled = False
        cmdClearAll.Enabled = False
    End If
    
End Sub

Private Sub optAdmin_Click()

    If (txtLogin.Text <> "" And txtPassword.Text <> "" And txtLastName.Text <> "" And txtFirstName.Text <> "" _
        And txtPhoneNo.Text <> "" And txtStID.Text <> "" And txtEmail.Text <> "" And txtURL.Text <> "" _
        And txtStreet.Text <> "" And txtCity.Text <> "" And txtProvince.Text <> "" And txtPostalCode.Text <> "" _
        And cmbCampus.Text <> "") Then
        
        cmdSave.Enabled = True ' enable cmbSize
        cmdClose.Enabled = True
    
    End If

End Sub

Private Sub optTech_Click()

    If (txtLogin.Text <> "" And txtPassword.Text <> "" And txtLastName.Text <> "" And txtFirstName.Text <> "" _
        And txtPhoneNo.Text <> "" And txtStID.Text <> "" And txtEmail.Text <> "" And txtURL.Text <> "" _
        And txtStreet.Text <> "" And txtCity.Text <> "" And txtProvince.Text <> "" And txtPostalCode.Text <> "" _
        And cmbCampus.Text <> "") Then
        
        cmdSave.Enabled = True ' enable cmbSize
        cmdClose.Enabled = True
    
    End If

End Sub

Private Sub optUser_Click()

    If (txtLogin.Text <> "" And txtPassword.Text <> "" And txtLastName.Text <> "" And txtFirstName.Text <> "" _
        And txtPhoneNo.Text <> "" And txtStID.Text <> "" And txtEmail.Text <> "" And txtURL.Text <> "" _
        And txtStreet.Text <> "" And txtCity.Text <> "" And txtProvince.Text <> "" And txtPostalCode.Text <> "" _
        And cmbCampus.Text <> "") Then
        
        cmdSave.Enabled = True ' enable cmbSize
        cmdClose.Enabled = True
    
    End If

End Sub

Private Sub txtCity_Change()

    If (txtLogin.Text <> "" And txtPassword.Text <> "" And txtLastName.Text <> "" And txtFirstName.Text <> "" _
        And txtPhoneNo.Text <> "" And txtStID.Text <> "" And txtEmail.Text <> "" And txtURL.Text <> "" _
        And txtStreet.Text <> "" And txtCity.Text <> "" And txtProvince.Text <> "" And txtPostalCode.Text <> "" _
        And cmbCampus.Text <> "") Then
        
        cmdSave.Enabled = True ' enable cmbSize
        cmdClose.Enabled = True
    
    End If

End Sub

Private Sub txtEmail_Change()
    
    If (txtLogin.Text <> "" And txtPassword.Text <> "" And txtLastName.Text <> "" And txtFirstName.Text <> "" _
        And txtPhoneNo.Text <> "" And txtStID.Text <> "" And txtEmail.Text <> "" And txtURL.Text <> "" _
        And txtStreet.Text <> "" And txtCity.Text <> "" And txtProvince.Text <> "" And txtPostalCode.Text <> "" _
        And cmbCampus.Text <> "") Then
        
        cmdSave.Enabled = True ' enable cmbSize
        cmdClose.Enabled = True
    
    End If

End Sub

Private Sub txtFirstName_Change()

    If (txtLogin.Text <> "" And txtPassword.Text <> "" And txtLastName.Text <> "" And txtFirstName.Text <> "" _
        And txtPhoneNo.Text <> "" And txtStID.Text <> "" And txtEmail.Text <> "" And txtURL.Text <> "" _
        And txtStreet.Text <> "" And txtCity.Text <> "" And txtProvince.Text <> "" And txtPostalCode.Text <> "" _
        And cmbCampus.Text <> "") Then
        
        cmdSave.Enabled = True ' enable cmbSize
        cmdClose.Enabled = True
    
    End If

End Sub

Private Sub txtLastName_Change()

    If (txtLogin.Text <> "" And txtPassword.Text <> "" And txtLastName.Text <> "" And txtFirstName.Text <> "" _
        And txtPhoneNo.Text <> "" And txtStID.Text <> "" And txtEmail.Text <> "" And txtURL.Text <> "" _
        And txtStreet.Text <> "" And txtCity.Text <> "" And txtProvince.Text <> "" And txtPostalCode.Text <> "" _
        And cmbCampus.Text <> "") Then
        
        cmdSave.Enabled = True ' enable cmbSize
        cmdClose.Enabled = True
    
    End If

End Sub

Private Sub txtLogin_Change()

    If (txtLogin.Text <> "" And txtPassword.Text <> "" And txtLastName.Text <> "" And txtFirstName.Text <> "" _
        And txtPhoneNo.Text <> "" And txtStID.Text <> "" And txtEmail.Text <> "" And txtURL.Text <> "" _
        And txtStreet.Text <> "" And txtCity.Text <> "" And txtProvince.Text <> "" And txtPostalCode.Text <> "" _
        And cmbCampus.Text <> "") Then
        
        cmdSave.Enabled = True ' enable cmbSize
        cmdClose.Enabled = True
        
    End If

End Sub

Private Sub txtPassword_Change()

    If (txtLogin.Text <> "" And txtPassword.Text <> "" And txtLastName.Text <> "" And txtFirstName.Text <> "" _
        And txtPhoneNo.Text <> "" And txtStID.Text <> "" And txtEmail.Text <> "" And txtURL.Text <> "" _
        And txtStreet.Text <> "" And txtCity.Text <> "" And txtProvince.Text <> "" And txtPostalCode.Text <> "" _
        And cmbCampus.Text <> "") Then
        
        cmdSave.Enabled = True ' enable cmbSize
        cmdClose.Enabled = True
    
    End If

End Sub

Private Sub txtPhoneNo_Change()

    If (txtLogin.Text <> "" And txtPassword.Text <> "" And txtLastName.Text <> "" And txtFirstName.Text <> "" _
        And txtPhoneNo.Text <> "" And txtStID.Text <> "" And txtEmail.Text <> "" And txtURL.Text <> "" _
        And txtStreet.Text <> "" And txtCity.Text <> "" And txtProvince.Text <> "" And txtPostalCode.Text <> "" _
        And cmbCampus.Text <> "") Then
        
        cmdSave.Enabled = True ' enable cmbSize
        cmdClose.Enabled = True
    
    End If

End Sub

Private Sub txtPostalCode_Change()
    
    If (txtLogin.Text <> "" And txtPassword.Text <> "" And txtLastName.Text <> "" And txtFirstName.Text <> "" _
        And txtPhoneNo.Text <> "" And txtStID.Text <> "" And txtEmail.Text <> "" And txtURL.Text <> "" _
        And txtStreet.Text <> "" And txtCity.Text <> "" And txtProvince.Text <> "" And txtPostalCode.Text <> "" _
        And cmbCampus.Text <> "") Then
        
        cmdSave.Enabled = True ' enable cmbSize
        cmdClose.Enabled = True
    
    End If

End Sub

Private Sub txtProvince_Change()

    If (txtLogin.Text <> "" And txtPassword.Text <> "" And txtLastName.Text <> "" And txtFirstName.Text <> "" _
        And txtPhoneNo.Text <> "" And txtStID.Text <> "" And txtEmail.Text <> "" And txtURL.Text <> "" _
        And txtStreet.Text <> "" And txtCity.Text <> "" And txtProvince.Text <> "" And txtPostalCode.Text <> "" _
        And cmbCampus.Text <> "") Then
        
        cmdSave.Enabled = True ' enable cmbSize
        cmdClose.Enabled = True
    
    End If

End Sub

Private Sub txtStID_Change()

    If (txtLogin.Text <> "" And txtPassword.Text <> "" And txtLastName.Text <> "" And txtFirstName.Text <> "" _
        And txtPhoneNo.Text <> "" And txtStID.Text <> "" And txtEmail.Text <> "" And txtURL.Text <> "" _
        And txtStreet.Text <> "" And txtCity.Text <> "" And txtProvince.Text <> "" And txtPostalCode.Text <> "" _
        And cmbCampus.Text <> "") Then
        
        cmdSave.Enabled = True ' enable cmbSize
        cmdClose.Enabled = True
    
    End If

End Sub

Private Sub txtStreet_Change()

    If (txtLogin.Text <> "" And txtPassword.Text <> "" And txtLastName.Text <> "" And txtFirstName.Text <> "" _
        And txtPhoneNo.Text <> "" And txtStID.Text <> "" And txtEmail.Text <> "" And txtURL.Text <> "" _
        And txtStreet.Text <> "" And txtCity.Text <> "" And txtProvince.Text <> "" And txtPostalCode.Text <> "" _
        And cmbCampus.Text <> "") Then
        
        cmdSave.Enabled = True ' enable cmbSize
        cmdClose.Enabled = True
    
    End If

End Sub

Private Sub txtURL_Change()

    If (txtLogin.Text <> "" And txtPassword.Text <> "" And txtLastName.Text <> "" And txtFirstName.Text <> "" _
        And txtPhoneNo.Text <> "" And txtStID.Text <> "" And txtEmail.Text <> "" And txtURL.Text <> "" _
        And txtStreet.Text <> "" And txtCity.Text <> "" And txtProvince.Text <> "" And txtPostalCode.Text <> "" _
        And cmbCampus.Text <> "") Then
        
        cmdSave.Enabled = True ' enable cmbSize
        cmdClose.Enabled = True
    
    End If

End Sub
