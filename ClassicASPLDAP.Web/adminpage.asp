<!--#include file="authmodule.asp"-->
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
    If Not IsUserInGroup("domain admins") Then
        ' Authorization failed user is not in group - stop processing request
        Response.Write("You do not have access to this page")
        Response.End
    End If
%>
<html>
    <body>
        <h1>This is the Admin Page</h1>
        <p><b>This page is only available to members of the Domain Admins group.</b></p>
        <div>
        <%
            ' Display all of the ASP server variables
            For Each strKey In Request.ServerVariables
                Response.Write("<b>Name:</b> " + strKey + " <b>Value:</b> " + Request.ServerVariables(strKey) + "<br/>")
            Next

            ' Display App Service App Settings by using Environment variable references
            Set WshShell = Server.CreateObject("WScript.Shell")
            Set WshSysEnv = WshShell.Environment("PROCESS")
            Response.Write("<b>Name:</b>APPSETTING_DOMAIN_NAME <b>Value:</b> " + WshSysEnv("APPSETTING_DOMAIN_NAME") +"<br/>")
            Response.Write("<b>Name:</b>APPSETTING_CONTAINER <b>Value:</b> " + WshSysEnv("APPSETTING_CONTAINER") +"<br/>")
            Response.Write("<b>Name:</b>APPSETTING_DOMAIN_USERNAME <b>Value:</b> " + WshSysEnv("APPSETTING_DOMAIN_USERNAME") +"<br/>")
            Response.Write("<b>Name:</b>APPSETTING_DOMAIN_PASSWORD <b>Value:</b> " + WshSysEnv("APPSETTING_DOMAIN_PASSWORD") +"<br/>")
        %>
        </div>
    </body>
</html>