VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help Desk Login"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   FillColor       =   &H80000017&
   ForeColor       =   &H80000017&
   Icon            =   "Frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.Frame frmButtons 
      Height          =   1215
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   360
         Left            =   240
         TabIndex        =   3
         Tag             =   "OK"
         Top             =   240
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         Height          =   360
         Left            =   240
         TabIndex        =   4
         Tag             =   "Cancel"
         Top             =   720
         Width           =   1140
      End
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Left            =   1425
      TabIndex        =   0
      Top             =   480
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1425
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2325
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Please login using your own student login name"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "Please login using your own student login name"
      Top             =   120
      Width           =   3345
   End
   Begin VB.Label lblLogin 
      Caption         =   "Login:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' declares all variables

' get the current login username
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Private Sub Form_Load()

    Dim sBuffer As String
    Dim lSize As Long

    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        txtLogin.Text = Left$(sBuffer, lSize)
    Else
        txtLogin.Text = vbNullString
    End If
    
End Sub

Private Sub cmdCancel_Click()
    ' sets OK variable to false
    OK = False
    ' ends the program
    End

End Sub

' cmdOK_Click()
' Purpose:
'   verify the login name
'   if admin set him up as admin
'   if tech set him up as tech
'   if others then make him a user
Private Sub cmdOK_Click()
    
    Set rstUser = db.OpenRecordset("SELECT * FROM tblUser")
    
    Do Until rstUser.EOF
    
        If UCase(txtLogin.Text) = UCase(rstUser!USER_LOGIN) Then
                
            If (rstUser!USER_ADMIN) Then
        
                If UCase(txtPassword.Text) = UCase(rstUser!USER_PASSWORD) Then
                    ' If correct set this settings and exit the current form
                    Access = "A"
                    Login = txtLogin.Text
                    ExitDown Me
                    Me.Hide
                    Exit Sub
                End If
                
                MsgBox "Password Incorrect, Please Enter the correct Password", vbInformation, "Password Incorect"
                txtPassword.Text = ""
                txtPassword.SetFocus
                
                Exit Sub
            
            End If
           
            If (rstUser!USER_TECH) Then
                ' Verify password entry here.....
                ' If correct set this settings and exit the current form
                If UCase(txtPassword.Text) = UCase(rstUser!USER_PASSWORD) Then
                    ' If correct set this settings and exit the current form
                    Access = "T"
                    Login = txtLogin.Text
                    ExitDown Me
                    Me.Hide
                    Exit Sub
                End If
                
                MsgBox "Password Incorrect, Please Enter the correct Password", vbInformation, "Password Incorect"
                txtPassword.Text = ""
                txtPassword.SetFocus
                
                Exit Sub
                
            Else
                If (txtPassword.Text = rstUser!USER_PASSWORD) Then
                
                    MsgBox "Welcome " & txtLogin.Text, vbInformation, "Welcome " & txtLogin.Text
                    Login = txtLogin.Text
                    Access = "U"
                    ExitDown Me
                    Me.Hide
                    Exit Sub
                
                End If
                
                MsgBox "Password Incorrect, Please Enter the correct Password", vbInformation, "Password Incorect"
                txtPassword.Text = ""
                txtPassword.SetFocus
            
                Exit Sub
            
            End If
        
        End If
        
        rstUser.MoveNext
    
    Loop
    
    MsgBox "User Login " & txtLogin.Text & " Not Found", vbInformation, "Welcome " & txtLogin.Text
    
    MsgBox "Welcome new user " & txtLogin.Text & vbCrLf & _
            "You need to setup your new account now" & vbCrLf & _
            "Please make sure you put in all the needed info" & vbCrLf & _
            "Before you can proceed.", vbInformation, "Welcome " & txtLogin.Text
    
    Login = txtLogin.Text
    Access = "U"
        
    Load frmNewUser
    frmNewUser.Show 1
    
    ExitDown Me
    Me.Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Unloads the form
    End
    
End Sub
