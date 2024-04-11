Attribute VB_Name = "HelpDeskModule"
Option Explicit ' Explicitly define variables

Public fMainForm As frmMain ' set the fMainForm to frmMain for public access
Public OK As Boolean ' set OK as boolean for public access
Public Access As String ' set access as string for public access
Public Login As String

Public db As DAO.Database
Public rstComputer As DAO.Recordset
Public rstRoom As DAO.Recordset
Public rstUser As DAO.Recordset
Public rstUserSelected As DAO.Recordset
Public rstProb As DAO.Recordset
Public rstCampus As DAO.Recordset
Public rstCampusRT As DAO.Recordset
Public rstUserRT As DAO.Recordset

Public UserId As Integer
Public probNum As Integer
Public searchRes As Integer

Dim X As Integer ' set X as integer
Dim i As Integer ' set i as integer
Dim fLogin As New frmLogin ' set fLogin as frmLogin
Dim ShowAtStartup As Long ' dim showatstartup as long

' This is the Sub where our program runs
' it calls the login form
' then the splash screen form
' then the Main form
Sub Main()
    
    Set db = OpenDatabase(App.Path & "\" & "HelpDesk.mdb") ' Connects to the database
    
    fLogin.Show vbModal ' display the fLogin form
    ' check the user access
    ' this uses Access global variable which came from fLogin
    i = MsgBox("Your access is " & Access, vbInformation, "Access")
    ' show frmSplash form
    frmSplash.Show
    ' refresh the form
    frmSplash.Refresh
    ' set the form with bevel on the inner edge
    FormInnerBevel frmSplash, 4
    ' set the form with bevel on the outer edge
    FormOuterBevel frmSplash, 5
    ' for next for the splash screen's progressbar
    For X = 1 To 1000
            frmSplash.pbSplash.Value = X
    Next
    ' disable the progressbar
    frmSplash.pbSplash.Enabled = False
    Set fMainForm = New frmMain ' set the fMainForm to frmMain
    Load fMainForm ' load fMainForm
    fMainForm.Caption = "Help Desk Application Version " & App.Major & "." & App.Minor & "." & App.Revision ' set the form title to application name version
    fMainForm.txtPRFullName.Text = Login
    Unload frmSplash ' unload splash form
    fMainForm.Show ' show Main form
    fMainForm.Refresh ' refresh Main form
    ' See if we should show tips on startup
    ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
    ' if not 0 then load frmtip
    ' and show it
    If ShowAtStartup <> 0 Then
        Load frmTip
        frmTip.Show
    End If
    
End Sub
