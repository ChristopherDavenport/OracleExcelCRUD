VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Login"
   ClientHeight    =   3540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4140
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Must Include ECSession Class Module To Utilize This

Private Sub BtnLogin_Click()
    Set ECSession = New ECSession
    
    ECSession.Initialize txtLoginUserName1.Text, txtLoginPass.Text, comboLoginDSN.Value
    
    Clear_Passwords
    
    If (ECSession.Validated = True) Then
        Unload Me
    Else
        Set ECSession = Nothing
    End If
    
End Sub

Private Sub btnResetPassword_Click()
    Set ECSession = New ECSession
    ECSession.Initialize txtResetUsername.Text, txtCurrentPass.Text, ComboReset.Value
    If (ECSession.Validated = True) Then
        ECSession.Reset_Password txtNewPass1, txtNewPass2
    Else
        Set ECSession = Nothing
        Unload Me
    End If
    
    Clear_Passwords
    
    
End Sub


Private Sub UserForm_Activate()
    'Add All Items to Combo Boxes
    comboLoginDSN.AddItem "PROD"
    comboLoginDSN.AddItem "TRNG"
    comboLoginDSN.AddItem "TEST"
    comboLoginDSN.AddItem "PROD12C"
    ComboReset.AddItem "PROD"
    ComboReset.AddItem "TRNG"
    ComboReset.AddItem "TEST"
    ComboReset.AddItem "PROD12C"
End Sub
'Section for Interoperability of Fields
'Corresponding Fields on Seperate Pages Change on Update
Private Sub txtLoginUserName1_Change()
    txtResetUsername.Text = txtLoginUserName1.Text
End Sub
Private Sub txtResetUsername_Change()
    txtLoginUserName1.Text = txtResetUsername.Text
End Sub
Private Sub txtCurrentPass_Change()
    txtLoginPass.Text = txtCurrentPass.Text
End Sub
Private Sub txtLoginPass_Change()
    txtCurrentPass.Text = txtLoginPass.Text
End Sub

Private Sub Clear_Passwords()
    txtCurrentPass.Text = ""
    txtNewPass1.Text = ""
    txtNewPass2.Text = ""
End Sub
