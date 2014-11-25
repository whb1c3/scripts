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
'slogFile = "C:\Users\e.belokon\Desktop\WUSforceupdate-" & nMonth & ".log"

tempFolder = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%Temp%")
slogFile = tempFolder & "\WUSforceupdate.log"

Set updateSession = CreateObject("Microsoft.Update.Session")
Set updateSearcher = updateSession.CreateupdateSearcher()

Set searchResult = updateSearcher.Search("IsInstalled=0 and Type='Software'")
Set updatesToInstall = CreateObject("Microsoft.Update.UpdateColl")

Set File = CreateObject("Scripting.FileSystemObject")
Set logFile = File.OpenTextFile(slogFile, 8, True)

logFile.WriteLine("***************************************************************")
logFile.WriteLine( "START TIME : " & now)
logFile.WriteLine( "Searching for updates..." & vbCRLF)

If searchResult.Updates.Count Then
	logFile.WriteLine("Найдено неустановленных обновлений: " & searchResult.Updates.Count)
	For I = 0 To searchResult.Updates.Count-1
	  set update = searchResult.Updates.Item(I)
	  If update.IsDownloaded = true Then
		updatesToInstall.Add(update)
		logFile.WriteLine( I + 1 & "> " & update.Title)
	  End if
	Next
	If updatesToInstall.Count Then
		logFile.WriteLine("Из них загружено и будет установлено: " & updatesToInstall.Count)
		
		'установка обновлений
		logFile.WriteLine( "Installing updates...")
		Set installer = updateSession.CreateUpdateInstaller()
		installer.Updates = updatesToInstall
		Set installationResult = installer.Install()
		
		'Output results of install
		logFile.WriteLine("Installation Result: " & installationResult.ResultCode)
		logFile.WriteLine("Reboot Required: " & installationResult.RebootRequired & vbCRLF)
		logFile.WriteLine("Listing of updates installed and individual installation results:")
		For I = 0 to updatesToInstall.Count - 1
		  logFile.WriteLine( I + 1 & "> " & updatesToInstall.Item(i).Title & ": " & installationResult.GetUpdateResult(i).ResultCode ) 
		Next		
	Else
		logFile.WriteLine("Обновления пока не загружены и будут пропущены: " & searchResult.Updates.Count)
	End if
Else
	logFile.WriteLine("Обновлений к установке не найдено")
End if

'проверим требуется ли перезагрузка после установки обновлений
Set systemInfo = CreateObject("Microsoft.Update.SystemInfo")
If systemInfo.RebootRequired Then
	 logFile.WriteLine("RebootRequired")
	 logFile.WriteLine("STOP TIME: " & now)
	 logFile.WriteLine("***************************************************************")
	 logFile.Close
	 ShutDown(ShutdownOption)
Else
	logFile.WriteLine("STOP TIME: " & now)
	logFile.WriteLine("***************************************************************")
	logFile.Close
End if	

Function ShutDown(sFlag)
  wscript.sleep 60
  Set OScoll = GetObject("winmgmts:{(Shutdown)}").ExecQuery("Select * from Win32_OperatingSystem") 
  For Each osObj in OScoll
    osObj.Win32Shutdown(sFlag)
  Next
End Function
