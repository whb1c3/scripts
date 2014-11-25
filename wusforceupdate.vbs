'Shutdown flags;
 const nLog_Off          =  0 
 const nForced_Log_Off   =  4  '( 0 + 4 ) 
 const nShutdown         =  1
 const nForced_Shutdown  =  5  '( 1 + 4 ) 
 const nReboot           =  2
 const nForced_Reboot    =  6  '( 2 + 4 )
 const nPower_Off        =  8
 const nForced_Power_Off = 12  '( 8 + 4 )

ShutdownOption = nForced_Reboot

'dt = date() : nMonth = Year(dt)*1e2 + Month(dt)
'sLogFile = "C:\Users\e.belokon\Desktop\WUSforceupdate-" & nMonth & ".log"

tempFolder = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%Temp%")
sLogFile = tempFolder & "\WUSforceupdate.log"

Set updateSession = CreateObject("Microsoft.Update.Session")
Set updateSearcher = updateSession.CreateupdateSearcher()

Set searchResult = updateSearcher.Search("IsInstalled=0 and Type='Software'")
Set updatesToInstall = CreateObject("Microsoft.Update.UpdateColl")

Set File = CreateObject("Scripting.FileSystemObject")
Set LogFile = File.OpenTextFile(sLogFile, 8, True)

LogFile.WriteLine("***************************************************************")
LogFile.WriteLine( "START TIME : " & now)
LogFile.WriteLine( "Searching for updates..." & vbCRLF)
LogFile.WriteLine( "List of applicable items on the machine:")

For I = 0 To searchResult.Updates.Count-1
  set update = searchResult.Updates.Item(I)
  If update.IsDownloaded = true Then
    updatesToInstall.Add(update)
    LogFile.WriteLine( I + 1 & "> " & update.Title)
  End if
Next

logFile.WriteLine( "Installing updates...")
Set installer = updateSession.CreateUpdateInstaller()
installer.Updates = updatesToInstall
Set installationResult = installer.Install()

'Output results of install
LogFile.WriteLine( "Installation Result: " & installationResult.ResultCode )
LogFile.WriteLine( "Reboot Required: " & installationResult.RebootRequired & vbCRLF )

LogFile.WriteLine( "Listing of updates installed " _
    & "and individual installation results:" )

For I = 0 to updatesToInstall.Count - 1
  LogFile.WriteLine( I + 1 & "> " & updatesToInstall.Item(i).Title _ 
    & ": " & installationResult.GetUpdateResult(i).ResultCode ) 
Next

If installationResult.RebootRequired = -1 then
  LogFile.WriteLine("RebootRequired")
  ShutDown(ShutdownOption) ' <-- normally now you should call for a R E B O O T.....
End if

'<-- O P T I O N A L

LogFile.WriteLine( "STOP TIME : " & now)
LogFile.WriteLine("***************************************************************")
LogFile.Close

Function ShutDown(sFlag)
 wscript.sleep 600
  Set OScoll = GetObject("winmgmts:{(Shutdown)}").ExecQuery("Select * from Win32_OperatingSystem") 
  For Each osObj in OScoll
    osObj.Win32Shutdown(sFlag)
  Next
End Function
