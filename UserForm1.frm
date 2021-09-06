VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Local Active Directory Search"
   ClientHeight    =   2400
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5100
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Enter()

Dim Name As String
Dim Email As String
Dim NetId As String
Dim Inputs As Variant
Dim count As Integer
Dim n As Integer
Dim terminate As Boolean
Dim Tws As Workbook
Set Tws = ThisWorkbook
    
    
    If Me.TextBox2 = "Check if entry is valid" Then Me.TextBox2 = ""
    If Me.TextBox3 = "Check if entry is valid" Then Me.TextBox3 = ""
    If Me.TextBox4 = "Check if entry is valid" Then Me.TextBox4 = ""
    
    Name = Me.TextBox2
    Email = Me.TextBox3
    NetId = Me.TextBox4
        
        count = 0
        terminate = False
        Call CheckBeforeRun(Name, Email, NetId, terminate)
        If terminate = True Then GoTo SubEnd
        
Call getNetIdFromName(Name, Email, NetId)

Me.TextBox4.Text = NetId
Me.TextBox3.Text = Email
Me.TextBox2.Text = Name

    If NetId = "Check if entry is valid" And NetId <> "" Then
            Me.TextBox4.ForeColor = &HFF&
        Else
            Me.TextBox4.ForeColor = &H80000012
    End If
    If Email = "Check if entry is valid" And Email <> "" Then
            Me.TextBox3.ForeColor = &HFF&
        Else
            Me.TextBox3.ForeColor = &H80000012
    End If
    If Name = "Check if entry is valid" And Name <> "" Then
            Me.TextBox2.ForeColor = &HFF&
        Else
            Me.TextBox2.ForeColor = &H80000012
    End If
    
SubEnd:
End Sub

Private Sub CommandButton2_Click()

Dim Tws As Workbook
Set Tws = ThisWorkbook

Me.TextBox4.Text = Clear
Me.TextBox4.ForeColor = &H80000012
Me.TextBox3.Text = Clear
Me.TextBox3.ForeColor = &H80000012
Me.TextBox2.Text = Clear
Me.TextBox2.ForeColor = &H80000012

End Sub
Sub UserForm_initialize()

UserForm1.Show vbontop

End Sub

Private Sub UserForm_Terminate()

Dim Tws As Workbook
Set Tws = ThisWorkbook

    Call Tws.CloseLDAP
    
End Sub

Sub CheckBeforeRun(Name, Email, NetId, terminate)

'input values count ( if more than 1 - reset the fields )
    Inputs = Array(Name, Email, NetId)
        For n = 0 To UBound(Inputs)
            If Inputs(n) <> "" Then count = count + 1
        Next n
       
    
    If count > 1 Then
        terminate = True
        Call CommandButton2_Click
        MsgBox "Only 1 input value allowed - please try again", , "Error"
        Exit Sub
    End If

 'checks if dipslay name is missing ", "

    If Name <> "" Then
        If InStr(1, Name, ", ", vbTextCompare) = 0 Then
            terminate = True
            Call CommandButton2_Click
            MsgBox "Please check Full Name input. Correct format is : ""lastname, firstname""", , "Input Error"
            Exit Sub
        End If
    End If

'-------------------------------------------------------------------------
'           On mouse move ( over the userform ) do something
'-------------------------------------------------------------------------

'Private Sub userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'
'Dim Tws As Workbook
'Set Tws = ThisWorkbook
'
'If (X > 130 Or X < 15) Then
'    If Tws.Application.WindowState <> xlMinimized Then Tws.Application.WindowState = xlMinimized
'End If
'If (Y > 40 Or Y < 10) Then
'    If Tws.Application.WindowState <> xlMinimized Then Tws.Application.WindowState = xlMinimized
'End If
'With UserForm1
'    .StartUpPosition = manual
'    .Show vbontop
'    End With
'End Sub

'Check any "Check if entry is valid"


    
End Sub
