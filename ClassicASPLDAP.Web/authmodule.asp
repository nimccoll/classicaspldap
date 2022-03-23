<%
'//===============================================================================
'// Microsoft FastTrack for Azure
'// Azure AD Classic ASP Group Authorization Sample
'// *** Must be used in conjunction with Azure App Service authentication ***
'//===============================================================================
'// Copyright © Microsoft Corporation.  All rights reserved.
'// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY
'// OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
'// LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
'// FITNESS FOR A PARTICULAR PURPOSE.
'//===============================================================================
%>
<%
    Function IsUserInGroup(groupName)
        Dim WshShell
        Dim WshSysEnv
        Dim authUser
        Dim samAccountName
        Dim ldapConnectionString
        Dim ldapConnection
        Dim ldapRecordset
        Dim providerType
        Dim ldapQuery
        Dim memGroups
        Dim userGroups

        If IsNull(Session("Groups")) Or Session("Groups") = "" Then
            ' Groups have not been checked previously - lookup groups for current user via LDAP query to domain controller
            Set WshShell = Server.CreateObject("WScript.Shell")
            Set WshSysEnv = WshShell.Environment("PROCESS")

            ' Azure App Service Authentication should populate the AUTH_USER server variable with the current user's UPN
            ' If it is not populated, check one of these other variables as one of them should contain the current user
            ' LOGON_USER
            ' REMOTE_USER
            ' HTTP_X_MS_CLIENT_PRINCIPAL_NAME
            ' HTTP_X_MS_CLIENT_PRINCIPAL_ID
            ' HTTP_X_MS_CLIENT_PRINCIPAL
            If Request.ServerVariables("AUTH_USER") = "" Then
                If WshSysEnv("APPSETTING_AUTH_USER") = "" Then ' Environment variable setting for local development only
                    authUser = ""
                Else
                    authUser = WshSysEnv("APPSETTING_AUTH_USER")
                End If
            Else
                authUser = Request.ServerVariables("AUTH_USER")
            End If

            If authUser = "" Then
                ' Authentication failed - stop processing request
                Response.Write("You do not have access to this page")
                Response.End
            Else
                On Error Resume Next ' Enable custom error handling

                ' Get the user's SAM account name
                samAccountName = Left(authUser, (Instr(1, authUser, "@") - 1))

                ' Create an LDAP connection
                ldapConnectionString = "LDAP://" + WshSysEnv("APPSETTING_DOMAIN_NAME") + "/" + WshSysEnv("APPSETTING_CONTAINER")
                Set ldapConnection = Server.CreateObject("ADODB.Connection") 
                Set ldapRecordset = Server.CreateObject("ADODB.Recordset")
                ldapConnection.Provider = "ADsDSOObject" 
                ldapConnection.Properties("User ID") = WshSysEnv("APPSETTING_DOMAIN_USERNAME")
                ldapConnection.Properties("Password") = WshSysEnv("APPSETTING_DOMAIN_PASSWORD")
                ldapConnection.Properties("Encrypt Password") = True
                providerType = "Active Directory Provider" 
                ldapConnection.Open providerType, WshSysEnv("APPSETTING_DOMAIN_USERNAME"), WshSysEnv("APPSETTING_DOMAIN_PASSWORD")
                If Err.Number <> 0 Then
                    Response.Write("Error Opening LDAP Connection: " & Err.Description)
                    Response.Write("<br/>")
                    Err.Clear
                End If

                ' Query the domain controller for the current user's list of groups            
                ldapQuery = "SELECT sn, givenName, memberOf FROM '" + ldapConnectionString + "' WHERE sAMAccountName = '" + samAccountName + "'"
                ldapRecordset.Open ldapQuery, ldapConnection, 1, 1
                If Err.Number <> 0 Then
                    Response.Write("Error Executing LDAP Query: " & Err.Description)
                    Response.Write("<br/>")
                    Err.Clear
                End If

                ' Extract groups from result set and store in Session state
                While ldapRecordset.EOF = False
                    memgroups = ldapRecordset.Fields("memberOf")
                    If Not IsNull(memgroups) Then
                        userGroups = LCase(Join(memgroups))
                    Else
                        userGroups = ""
                    End If
                    Session("Groups") = userGroups
                    ldapRecordset.MoveNext 
                Wend

                On Error Goto 0 ' Disable custom error handling
            End If
        Else
            ' Groups have been previously retrieved - pull from Session state
            userGroups = Session("Groups")
        End If

        ' Is the current user in the specified group?
        If InStr(1, userGroups, LCase(groupName)) > 0 Then
            IsUserInGroup = True
        Else
            IsUserInGroup = False
        End If
    End Function
%>