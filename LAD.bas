Attribute VB_Name = "LAD"
Public Const LAD = "LDAP://YourDomainHere.com"

Sub getNetIdFromName(Name As String, Email As String, NetId As String)

    Dim dbConnection As Object
    Dim dbRecordset As Object
    Dim sqlString As String
    Dim sqlResult()
    
On Error GoTo 0
On Error GoTo DebugSub:


'Search With Name

    If Name <> "" And Email = "" And NetId = "" Then
        
'get NetId
            sqlString = "SELECT samAccountName " _
                        & "FROM '" & LAD & "' " _
                        & "WHERE displayName = '" & Name & "'"
                
                Debug.Print sqlString
                
                Set dbConnection = CreateObject("ADODB.Connection")
                dbConnection.Provider = "ADsDSOObject"
                dbConnection.Properties("ADSI Flag") = 1
                dbConnection.Open "Ads Provider"
            
                Set dbRecordset = dbConnection.Execute(sqlString)
                sqlResult() = dbRecordset.GetRows()
                
                NetId = CStr(sqlResult(0, 0))
                
            
'get Email
            sqlString = "SELECT mail " _
                        & "FROM '" & LAD & "' " _
                        & "WHERE displayName = '" & Name & "'"
                        
                Set dbConnection = CreateObject("ADODB.Connection")
                dbConnection.Provider = "ADsDSOObject"
                dbConnection.Properties("ADSI Flag") = 1
                dbConnection.Open "Ads Provider"
            
                Set dbRecordset = dbConnection.Execute(sqlString)
            
                sqlResult() = dbRecordset.GetRows()
                Email = CStr(sqlResult(0, 0))
        
'Search with EMail

    ElseIf Name = "" And Email <> "" And NetId = "" Then
   
'get netid
            sqlString = "SELECT samAccountName " _
                        & "FROM '" & LAD & "' " _
                        & "WHERE mail = '" & Email & "'"
        
                Set dbConnection = CreateObject("ADODB.Connection")
                dbConnection.Provider = "ADsDSOObject"
                dbConnection.Properties("ADSI Flag") = 1
                dbConnection.Open "Ads Provider"
            
                Set dbRecordset = dbConnection.Execute(sqlString)
            
                sqlResult() = dbRecordset.GetRows()
                NetId = CStr(sqlResult(0, 0))
        
'get Name
            sqlString = "SELECT displayName " _
                        & "FROM '" & LAD & "' " _
                        & "WHERE mail = '" & Email & "'"
        
                Set dbConnection = CreateObject("ADODB.Connection")
                dbConnection.Provider = "ADsDSOObject"
                dbConnection.Properties("ADSI Flag") = 1
                dbConnection.Open "Ads Provider"
            
                Set dbRecordset = dbConnection.Execute(sqlString)
            
                sqlResult() = dbRecordset.GetRows()
                Name = CStr(sqlResult(0, 0))
    
'Search with NetID
    
    ElseIf Name = "" And Email = "" And NetId <> "" Then

'get Name
                sqlString = "SELECT displayName " _
                            & "FROM '" & LAD & "' " _
                            & "WHERE samAccountName = '" & NetId & "'"
            
                Set dbConnection = CreateObject("ADODB.Connection")
                dbConnection.Provider = "ADsDSOObject"
                dbConnection.Properties("ADSI Flag") = 1
                dbConnection.Open "Ads Provider"
            
                Set dbRecordset = dbConnection.Execute(sqlString)

                sqlResult() = dbRecordset.GetRows()
                Name = CStr(sqlResult(0, 0))
        
        
'get Email
                sqlString = "SELECT mail " _
                            & "FROM '" & LAD & "' " _
                            & "WHERE samAccountName = '" & NetId & "'"
        
                Set dbConnection = CreateObject("ADODB.Connection")
                dbConnection.Provider = "ADsDSOObject"
                dbConnection.Properties("ADSI Flag") = 1
                dbConnection.Open "Ads Provider"
            
                Set dbRecordset = dbConnection.Execute(sqlString)
            
                sqlResult() = dbRecordset.GetRows()
                Email = CStr(sqlResult(0, 0))

    End If
  
GoTo exitsub:
  
  
                '-------------------------------------------------------------------------------------------------
DebugSub:
                '-------------------------------------------------------------------------------------------------
                'MsgBox (Err)
                    If Err = 94 Or 3021 Then
                    
                        If Name <> "" Then Name = "Check if entry is valid"
                        If Email <> "" Then Email = "Check if entry is valid"
                        If NetId <> "" Then NetId = "Check if entry is valid"
                            
                            GoTo exitsub:  'create ErrorSub
                    End If
                MsgBox ("Error code : " & Err & ".provide this code to maciej.dlugosz1@delphi.com")
                'Resume 0
                
                '-------------------------------------------------------------------------------------------------

                '-------------------------------------------------------------------------------------------------

    
exitsub:



    'clear variables :
    Set dbConnection = Nothing
    Set dbRecordset = Nothing
    sqlString = ""

End Sub

