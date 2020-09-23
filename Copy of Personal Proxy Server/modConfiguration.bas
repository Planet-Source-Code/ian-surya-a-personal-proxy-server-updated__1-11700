Attribute VB_Name = "modConfiguration"
Option Explicit

Public LocalIP As String

Public Const ConfigFileName = "PPS.INI"
Public Const ConfigSectionName = "Configuration"

Public Const vbTransparant = &H8000000F
Public Const C_USER_AGENT_IE5 = "Mozilla/4.0 (compatible; MSIE 5.5; Windows 98; Win 9x 4.90)"
Public Const C_USER_AGENT_PPS = "Personal Proxy Server"
Public Const C_PERSONAL_PROXY = "Tristan's Personal Proxy Server"

Public UserAgent As String
Public PersonalProxyName As String
Public LocalComputerName As String

Public MaximumConnection As Long
Public ListeningPort As Long

Public UseProxy As Boolean
Public UseNTLM As Boolean
Public netProxy As CProxy

Public UseAuthentication As Boolean

Public ChildOpen As Boolean
Public ProxySet As Boolean

Public UserOpen As Boolean
Public LoginOpen As Boolean
Public LoginSet As Boolean

Public Filter_Reload As Boolean
Public Filter_Disable_Cookie As Boolean
Public Filter_Hide_Server As Boolean
Public Filter_Hide_Proxy As Boolean
Public Filter_Hide_UserAgent As Boolean

Public Sub LoadConfigurationScreen(ParentForm As Form, ProxySetting As CProxy)
    ChildOpen = True
    ProxySet = False
    ParentForm.Enabled = False
    With ProxySetting
        frmConfiguration.OpenCommunication _
                            ParentForm, 0, _
                            IIf(UseProxy, vbChecked, vbUnchecked), _
                            .Server, _
                            .Port, _
                            MaximumConnection, _
                            ListeningPort, _
                            IIf(Filter_Reload, vbChecked, vbUnchecked), _
                            IIf(Filter_Disable_Cookie, vbChecked, vbUnchecked), _
                            IIf(Filter_Hide_Server, vbChecked, vbUnchecked), _
                            IIf(Filter_Hide_Proxy, vbChecked, vbUnchecked), _
                            IIf(Filter_Hide_UserAgent, vbChecked, vbUnchecked), _
                            PersonalProxyName, _
                            LocalComputerName, _
                            UserAgent, _
                            IIf(UseAuthentication, vbChecked, vbUnchecked)

        While ChildOpen
            DoEvents
        Wend
        If ProxySet Then
            If frmConfiguration.chkUseProxy = vbChecked Then
                UseProxy = True
                ProxySetting.Server = frmConfiguration.txtProxySetting(0).Text
                ProxySetting.Port = Val(frmConfiguration.txtProxySetting(1).Text)
            Else
                UseProxy = False
                ProxySetting.Server = ""
                ProxySetting.Port = 0
            End If
            MaximumConnection = Val(frmConfiguration.txtProxySetting(2))
            ListeningPort = Val(frmConfiguration.txtProxySetting(3))
            If frmConfiguration.chkConfig(0) = vbChecked Then
                Filter_Reload = True
            Else
                Filter_Reload = False
            End If
            If frmConfiguration.chkConfig(1) = vbChecked Then
                Filter_Disable_Cookie = True
            Else
                Filter_Disable_Cookie = False
            End If
            If frmConfiguration.chkConfig(2) = vbChecked Then
                Filter_Hide_Server = True
                PersonalProxyName = frmConfiguration.txtProxySetting(4)
            Else
                Filter_Hide_Server = False
            End If
            If frmConfiguration.chkConfig(3) = vbChecked Then
                Filter_Hide_Proxy = True
                LocalComputerName = frmConfiguration.txtProxySetting(5)
            Else
                Filter_Hide_Proxy = False
            End If
            If frmConfiguration.chkConfig(4) = vbChecked Then
                Filter_Hide_UserAgent = True
                UserAgent = frmConfiguration.txtProxySetting(6)
            Else
                Filter_Hide_UserAgent = False
            End If
            If frmConfiguration.chkConfig(5) = vbChecked Then
                UseAuthentication = True
            Else
                UseAuthentication = False
            End If
            SaveProxyConfiguration
        End If
    End With
    Unload frmConfiguration
    ParentForm.Enabled = True
End Sub

Public Sub SaveProxyConfiguration()
    WriteInitString ConfigSectionName, "MaximumConnection", CStr(MaximumConnection), App.Path & "\" & ConfigFileName
    WriteInitString ConfigSectionName, "ListeningPort", CStr(ListeningPort), App.Path & "\" & ConfigFileName
    WriteInitString ConfigSectionName, "LocalHostName", LocalComputerName, App.Path & "\" & ConfigFileName
    WriteInitString ConfigSectionName, "UserAgent", UserAgent, App.Path & "\" & ConfigFileName
    WriteInitString ConfigSectionName, "PersonalProxyName", PersonalProxyName, App.Path & "\" & ConfigFileName
    WriteInitString ConfigSectionName, "ForceReload", CStr(CInt(Filter_Reload)), App.Path & "\" & ConfigFileName
    WriteInitString ConfigSectionName, "DisableCookie", CStr(CInt(Filter_Disable_Cookie)), App.Path & "\" & ConfigFileName
    WriteInitString ConfigSectionName, "HideServer", CStr(CInt(Filter_Hide_Server)), App.Path & "\" & ConfigFileName
    WriteInitString ConfigSectionName, "HideProxy", CStr(CInt(Filter_Hide_Proxy)), App.Path & "\" & ConfigFileName
    WriteInitString ConfigSectionName, "HideUserAgent", CStr(CInt(Filter_Hide_UserAgent)), App.Path & "\" & ConfigFileName
    WriteInitString ConfigSectionName, "UseProxy", CStr(CInt(UseProxy)), App.Path & "\" & ConfigFileName
    WriteInitString ConfigSectionName, "ProxyServer", netProxy.Server, App.Path & "\" & ConfigFileName
    WriteInitString ConfigSectionName, "ProxyPort", CStr(netProxy.Port), App.Path & "\" & ConfigFileName
    WriteInitString ConfigSectionName, "UseAuthentication", CStr(CInt(UseAuthentication)), App.Path & "\" & ConfigFileName
End Sub

Public Sub LoadProxyConfiguration()
Dim tmpLong As Long, tmpString As String

    If FileLen(App.Path & "\" & ConfigFileName) = 0 Then LoadDefaultConfiguration
    tmpLong = Val(GetInitString(ConfigSectionName, "MaximumConnection", App.Path & "\" & ConfigFileName))
    MaximumConnection = tmpLong
    tmpLong = Val(GetInitString(ConfigSectionName, "ListeningPort", App.Path & "\" & ConfigFileName))
    ListeningPort = tmpLong
    tmpString = GetInitString(ConfigSectionName, "LocalHostName", App.Path & "\" & ConfigFileName)
    If tmpString <> "" Then LocalComputerName = tmpString
    tmpString = GetInitString(ConfigSectionName, "UserAgent", App.Path & "\" & ConfigFileName)
    If tmpString <> "" Then UserAgent = tmpString
    tmpString = GetInitString(ConfigSectionName, "PersonalProxyName", App.Path & "\" & ConfigFileName)
    If tmpString <> "" Then PersonalProxyName = tmpString
    tmpLong = Val(GetInitString(ConfigSectionName, "ForceReload", App.Path & "\" & ConfigFileName))
    Filter_Reload = tmpLong
    tmpLong = Val(GetInitString(ConfigSectionName, "DisableCookie", App.Path & "\" & ConfigFileName))
    Filter_Disable_Cookie = tmpLong
    tmpLong = Val(GetInitString(ConfigSectionName, "HideServer", App.Path & "\" & ConfigFileName))
    Filter_Hide_Server = tmpLong
    tmpLong = Val(GetInitString(ConfigSectionName, "HideProxy", App.Path & "\" & ConfigFileName))
    Filter_Hide_Proxy = tmpLong
    tmpLong = Val(GetInitString(ConfigSectionName, "HideUserAgent", App.Path & "\" & ConfigFileName))
    Filter_Hide_UserAgent = tmpLong
    tmpLong = Val(GetInitString(ConfigSectionName, "UseProxy", App.Path & "\" & ConfigFileName))
    UseProxy = tmpLong
    tmpString = GetInitString(ConfigSectionName, "ProxyServer", App.Path & "\" & ConfigFileName)
    If tmpString <> "" Then netProxy.Server = tmpString
    tmpLong = Val(GetInitString(ConfigSectionName, "ProxyPort", App.Path & "\" & ConfigFileName))
    netProxy.Port = tmpLong
    tmpLong = Val(GetInitString(ConfigSectionName, "UseAuthentication", App.Path & "\" & ConfigFileName))
    UseAuthentication = tmpLong
End Sub

Private Sub LoadDefaultConfiguration()
    MaximumConnection = 10000
    ListeningPort = 8080
    UserAgent = C_USER_AGENT_PPS
    PersonalProxyName = C_PERSONAL_PROXY & "/" & App.Major & "." & App.Minor
    Filter_Reload = False
    Filter_Disable_Cookie = False
    Filter_Hide_Server = False
    Filter_Hide_Proxy = False
    Filter_Hide_UserAgent = False
    UseAuthentication = False
End Sub
