'.SYNOPSIS
'   
'	This PowerShell script can be used in the CONFIGMGR DCM feature to make sure we get the compliance for the machines which are not having the MS17-010 patches.
              
'.DESCRIPTION
'	This PowerShell script can be used in the CONFIGMGR DCM feature to make sure we get the compliance for the machines which are not having the MS17-010 patches. We have created a Non compliance collection and target the machines from the patches directly from the CONFIGMGR console.
'	
'	.PARAMETER DeploymentType: Not applicable
'
'              
'.NOTES
'    Version                           : 1.0
'    Author                            : Umesh Sawant
'    Date                              : 13/07/2019
'    Email                             : Umesh.b.sawant@capgemini.com
'    *********************************************************************************************************


Dim osName, X, Y, S
Dim WshShell      : Set WshShell      = CreateObject ("WScript.Shell")
Dim WshSysEnvP    : Set WshSysEnvP    = WshShell.Environment("Process")
Dim WshSysEnvS    : Set WshSysEnvS    = WshShell.Environment("System")
Dim fso        : Set fso        = CreateObject("Scripting.FileSystemObject")

Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 
for each System in SystemSet 
 osName = System.Caption
BuildNumber= System.BuildNumber
next

X = wshShell.ExpandEnvironmentStrings("%systemroot%") & "\system32\drivers\srv.sys"
Y = wshShell.ExpandEnvironmentStrings("%systemroot%") & "\system32\drivers\srv2.sys"

If fso.FileExists(X) Then
  S = X
ElseIf fso.FileExists(Y) Then
  S = Y
End If

if fso.FileExists(S) Then
              if instr(osName,"Microsoft Windows 10") > 0 or instr(osName,"Microsoft Windows Server 2016") > 0 Then
                 versionInfo = fso.GetFileVersion(S) 
                 if instr(osName,"Microsoft Windows 10") > 0 then                                                   
                                fileVersion  = GetProductVersion(mid(S,1,instrrev(S,"\")-1),mid(S,instrrev(S,"\")+1,len(X)))                          
                 elseif instr(osName,"Microsoft Windows Server 2016") > 0 Then
                                 fileVersion  = GetProductVersion(mid(S,1,instrrev(S,"\")-1),mid(S,instrrev(S,"\")+1,len(X)))
                 else                                 
                      fileVersion =  versionInfo 
                 end if
                 If Err.Number <> 0 then
                                           state = Null
                                           Call EndScript(state)
                end if
              Else
                             fileVersion = fso.GetFileVersion(S) 
              End If
Else
              state = Null
              Call EndScript(state)       
End If

If instr(osName,"Vista") > 0 or instr(osName,"2008") > 0  and not instr(osName,"R2") <> 0 Then
              ArrayfileVersion=split(fileVersion,".")
              If Mid(ArrayfileVersion(3),1,1) = 1 then
                             currentOS = osName & " GDR"
                             expectedVersion = "6.0.6002.19743"
              ElseIf Mid(ArrayfileVersion(3),1,1) = 2 then
                             currentOS = osName & " LDR"
                             expectedVersion = "6.0.6002.24067"
              Else
                             currentOS = osName
                             expectedVersion = "9.9.9999.99999"
              End if
ElseIf instr(osName,"Windows 7")  > 0 or instr(osName,"2008 R2") > 0  Then     
              currentOS = osName & " LDR"
              expectedVersion = "6.1.7601.23689"
ElseIf instr(osName,"Windows 8.1")  > 0 or instr(osName,"2012 R2") > 0  Then     
              currentOS = osName & " LDR"
              expectedVersion = "6.3.9600.18604"
ElseIf instr(osName,"Windows 8")  > 0 or instr(osName,"2012") > 0  Then              
              currentOS = osName & " LDR"
              expectedVersion = "6.2.9200.22099"
ElseIf instr(osName,"Windows 10")  > 0 Then    
              If BuildNumber = "10240" then
                 currentOS =  osName & " TH1"
                 expectedVersion = "10.0.10240.17319"
              ElseIf BuildNumber = "10586" then
                 currentOS =  osName & " TH2"
                 expectedVersion = "10.0.10586.839"
              ElseIf BuildNumber = "14393" then
                 currentOS =  osName & " RS1"
                 expectedVersion = "10.0.14393.953"
    ElseIf BuildNumber = "15063" then
                 currentOS =  osName & " RS2"
                 state = "Patched"
                 Call EndScript(state)                  
              End if
ElseIf instr(osName,"2016")  > 0 Then     
              currentOS = osName 
               expectedVersion = "10.0.14393.953"     
ElseIf instr(osName,"Windows XP")  > 0 Then      
              currentOS = osName 
               expectedVersion = "5.1.2600.7208"        
ElseIf instr(osName,"Server 2003")  > 0 Then       
              currentOS = osName 
               expectedVersion = "5.2.3790.6021"        
Else
              currentOS = osName 
              expectedVersion = "9.9.9999.99999"      
End if


if expectedVersion <> "" then
	arr1=split(fileVersion,".")
	arr2=split(expectedVersion,".")
	for i = 0 to ubound(arr1)
	if CLng(arr1(i)) >= CLng(arr2(i)) then
				  state = "Patched"
	Else
				  state = "NotPatched" & fileVersion
				  Call EndScript(state)       
	end if
next
Else
	state = "Patched"
End if

sub EndScript(Strstate)
    WScript.Echo state
	WScript.Quit
end sub


Function GetProductVersion (sFilePath, sProgram)
Dim FSO,objShell, objFolder, objFolderItem, i 
Set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(sFilePath & "\" & sProgram) Then
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(sFilePath)
    Set objFolderItem = objFolder.ParseName(sProgram)
    Dim arrHeaders(300)
    For i = 0 To 300
        arrHeaders(i) = objFolder.GetDetailsOf(objFolder.Items, i)
        'WScript.Echo i &"- " & arrHeaders(i) & ": " & objFolder.GetDetailsOf(objFolderItem, i)
        If lcase(arrHeaders(i))= "product version" Then
            GetProductVersion= objFolder.GetDetailsOf(objFolderItem, i)
            Exit For
        End If
    Next
End If
End Function
WScript.Echo state
