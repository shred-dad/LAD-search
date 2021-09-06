VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'otwarty gdy sa inne okna excela = OK
'gdy tylko LDAP jest otwarty - otwarcie nowego excela laduje go w tym samym procesie
'znalec jaka opcje zeby byo on window minimized

Private Sub Workbook_Open()

Dim Tws As Workbook
Set Tws = ThisWorkbook
Dim protected As Boolean

    If Tws.Windows.Item(1).Application.ProtectedViewWindows.count > 0 Then
        protected = True
            Call OpenNewExcelInstance(protected)
    End If
    
    'Tws.Windows.Item(1).Visible = True
    
    If Tws.Application.Windows.count > 1 Then
        If Tws.ReadOnly = True Then
            Tws.Windows.Item(1).WindowState = xlMinimized
        End If
        Tws.Windows.Item(1).Visible = False
        
        Call OpenNewExcelInstance(protected)
       
    Else
        Tws.Windows.Item(1).Visible = False
        Tws.Application.WindowState = xlMinimized
        'Tws.Windows.Application.DisplayAlerts = False ' might be a problem with different PC setups
    End If
        
        Call RunLDAPScript(Tws)
End Sub

Sub RunLDAPScript(Tws)

    With UserForm1.TextBox2
        .SetFocus
    End With
    
End Sub

Sub CloseLDAP()
    
    Dim Tws As Workbook
    Set Tws = ThisWorkbook
    
    If Tws.Application.Windows.count > 1 Then
        Tws.Windows.Application.DisplayAlerts = False
        Tws.Windows.Item(1).Close SaveChanges:=False
    Else
        Tws.Application.Quit
    End If
    
End Sub

Sub OpenNewExcelInstance(protected)
 
Dim xlApp As Excel.Application
Set xlApp = New Excel.Application
    xlApp.DisplayAlerts = False
    
    If protected = True Then
        xlApp.ProtectedViewWindows.Open Filename:=ThisWorkbook.FullName
        xlApp.ActiveProtectedViewWindow.Edit
        
    ElseIf protected = False Then
    
        xlApp.Workbooks.Open Filename:=ThisWorkbook.FullName, ReadOnly:=True, IgnoreReadOnlyRecommended:=True
    
    End If

xlApp.Visible = True
Set xlApp = Nothing

    Call CloseLDAP
    
End Sub

Sub workbook_change()
    Dim Tws As Workbook
    Set Tws = ThisWorkbook
    MsgBox "works"
        Tws.Application.WindowState = xlMinimized
        UserForm1.Show vbontop
        
End Sub

Private Sub Workbook_BeforeClose(cancel As Boolean)
        
        Dim Tws As Workbook
        Set Tws = ThisWorkbook

        Tws.Windows.Item(1).Visible = False
        Tws.Application.WindowState = xlMinimized
        Tws.Windows.Application.DisplayAlerts = False
        
        
End Sub





