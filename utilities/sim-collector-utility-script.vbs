' Copyright (c) AppDynamics, Inc., and its affiliates, 2014, 2015
' All Rights Reserved
' THIS IS UNPUBLISHED PROPRIETARY CODE OF APPDYNAMICS, INC.
' The copyright notice above does not evidence any
' actual or intended publication of such source code.

' Require explicitly declaring variables
Option Explicit

' Force this script to run using cscript
Sub forceCScriptExecution
    Const QUOTE = """"
    Dim Arg, Str
    If Not LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then
        For Each Arg In WScript.Arguments
            If InStr( Arg, " " ) Then Arg = QUOTE & Arg & QUOTE
            Str = Str & " " & Arg
        Next
        CreateObject( "WScript.Shell" ).Run _
            "cscript //nologo " & _
            QUOTE & WScript.ScriptFullName & QUOTE & _
            " " & Str
        WScript.Quit
    End If
End Sub
forceCScriptExecution

On Error Resume Next

Dim objWbemLocator
Set objWbemLocator = CreateObject( "WbemScripting.SWbemLocator" )
Const wbemFlagReturnWhenComplete = &h0
Const wbemFlagForwardOnly = &h20

if Err.Number Then
    WScript.echo "Cannot create scripting object!"
    WScript.echo "Error # " & Hex(Err.Number) & " " & Err.Description
End If

Dim objWMIService
Set objWMIService = objWbemLocator.ConnectServer( "localhost", "Root\CIMV2" )

If Err.Number Then
    WScript.echo "Cannot connect to server!"
    WScript.echo "Error=" & Hex(Err.Number) & " " & Err.Description & " " & Err.Source
End If


'
'Checks if the WMI class exists.
'
Function attemptConnection( wmiClassString )
    On Error Resume Next
    Dim wmiClassNotFound
    Dim wmiQuery
    Dim errorNum
    Dim errorDesc

    WScript.echo "Attempting to connect to " & wmiClassString

    ' Check to see if the query call returns an error
    Set wmiQuery = objWMIService.ExecQuery("Select * from " & wmiClassString, "WQL", wbemFlagReturnWhenComplete + wbemFlagForwardOnly)
    If Err.Number <> 0 then
        Wscript.echo "Connection failed"
        WScript.echo "Error number (hex): " & Hex(Err.Number)
        WScript.echo "Error description: " & Err.Description
    Else
        ' Query returns without errors, this WMI class exists
        WScript.Echo "Connection successful"

    End If
    WScript.Echo ""
    Err.Clear
End Function

WScript.Timeout = 30

WScript.Echo "------------------------------------------------------------------------"
WScript.Echo "Attempting connection to each WMI class."
WScript.Echo ""
WScript.Echo "If the connection is successful, the script will print out 'Connection successful'."
WScript.Echo ""
WScript.Echo "If the connection fails, the script will print out 'Connection failed',"
WScript.Echo "along with the error number and description received from the WMI service."
WScript.Echo ""
WScript.Echo "If the script times out, you may have connection issues to the WMI service itself."
WScript.Echo ""

attemptConnection("Win32_PerfRawData_PerfDisk_LogicalDisk")

attemptConnection("Win32_LogicalDisk")

attemptConnection("Win32_Processor")

attemptConnection("Win32_PerfRawData_PerfOS_Processor")

attemptConnection("Win32_NetworkAdapter")

attemptConnection("Win32_NetworkAdapterConfiguration")

attemptConnection("Win32_PerfRawData_Tcpip_NetworkInterface")

attemptConnection("Win32_Process")

attemptConnection("Win32_PerfRawData_PerfProc_Process")

attemptConnection("Win32_PhysicalMemory")

attemptConnection("Win32_OperatingSystem")

WScript.Echo "Finished attempting to connect to WMI classes."
WScript.Echo "------------------------------------------------------------------------"
WScript.Echo ""