Add-Type –AssemblyName UIAutomationClient
Add-Type –AssemblyName UIAutomationTypes



function FindAndCloseVSTOInstaller () {

    $result = $false

        $timer = [Diagnostics.Stopwatch]::StartNew()
        $maxWaitSeconds = 10

        Do {$installer = get-process VSTOInstaller -ErrorAction SilentlyContinue
            if ($timer.Elapsed.TotalSeconds -ge $maxWaitSeconds) {break}
        }
        while ($installer -eq $null)

        if ($installer -eq $null) {
            $returnresult = writelog "ERROR: Cannot find VSTO Installer"
            return $false
        }

        Try {
            $installerID = $installer.Id
            $root = [Windows.Automation.AutomationElement]::RootElement
            $condition = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ProcessIdProperty, $installerID)
            
            $timer = [Diagnostics.Stopwatch]::StartNew()
            $maxWaitSeconds = 240
            Do {
                if ($timer.Elapsed.TotalSeconds -ge $maxWaitSeconds) {break}
                $installerUI = $root.FindFirst([Windows.Automation.TreeScope]::Children, $condition)
            }
            while ($installerUI -eq $null)



            $timer = [Diagnostics.Stopwatch]::StartNew()
            $maxWaitSeconds = 240
            Do {

                $installerButtons = $installerUI.FindAll([Windows.Automation.TreeScope]::Descendants, `
                        (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, `
                        [System.Windows.Automation.ControlType]::Button)))
                            
                $closeButton = $null

                foreach ($installerButton in $installerButtons) {
                    if (($installerButton.Current.Name -eq "Close") -and ($installerButton.Current.IsEnabled)) {
                        $closeButton = $installerButton
                            
                    }
                }
            } while ($closeButton -eq $null)


            $installerTexts = $installerui.FindAll([Windows.Automation.TreeScope]::Descendants, `
                        (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, `
                        [System.Windows.Automation.ControlType]::Text)))
        
            $result = $true

            foreach ($installerText in $installerTexts) {
                if ($installerText.current.Name -like "*error*") {
                    $result = $false
                }
            }

            $closeResult = $closeButton.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()

        <#
	    $condition1 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ClassNameProperty, "Button")
	    $condition2 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, $name)
	    $condition = New-Object Windows.Automation.AndCondition($condition1, $condition2)
	    $button = $calcUI.FindFirst([Windows.Automation.TreeScope]::Descendants, $condition)
	    $button.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
        #>

        
        }
        Catch {writelog "ERROR Closing PIX Installer: $_"}

    return $result
}


    function WriteLog
    {
        Param (
            [string]$LogString
            )

        Try {
            $LogPath = Split-Path $LogFile -Parent

            #Create the Log file Folder if it doesn't exist
            if ( $(Test-Path $LogPath) -eq $false) {
                New-Item -ItemType Directory -Force -Path $LogPath
            }

            $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
            $LogMessage = "$Stamp`t$LogString"

            Add-content $LogFile -value $LogMessage 
            return $true
        }
        Catch {
            return $false
        }
    }

    Function Create_Shortcut {
        param (
        [string]$originalFilePath,
        [string]$shortcutDestFolder)

        if (!(test-path $originalFilePath)) {
            
            $result = writelog "ERROR: Create_Shortcut Cannot find $originalFilePath"
            return $false
        }

        if (!(test-path $shortcutDestFolder)) {
            
            $result = writelog "ERROR: Create_Shortcut Cannot find $shortcutDestFolder"
            return $false
        }

        Try {

            $shortcutDestination = split-path $originalFilePath -leaf 
            $shortcutDestination = "$shortcutDestFolder\$shortcutDestination.lnk"

            $WshShell = New-Object -comObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut($shortcutDestination)
            $result = $Shortcut.TargetPath = $originalFilePath
            $result = $Shortcut.Save()
            return $true
        }
        Catch {
            $result = writelog "ERROR: Create_Shortcut cannot create shortcut: $_"
            return $false
        }

    }

    function Get_Reg_Value {
        param (
            [string]$RegistryPath,
            [string]$RegistryName)
         return Get-ItemPropertyValue -path $RegistryPath -name $RegistryName 
     }

    function Test_Update_Reg {
        param (
            [string]$RegistryPath,
            [string]$RegistryName,
            [string]$RegistryValue,
            [string]$LogFile = ""
        )
        
        $returnResult = $false            
        $currentAction = "Updating Registry - $RegistryName"

        Try {

            #if registry path doesn't exist, create it.
            if (-not(test-path $RegistryPath)) {
                $newPath = split-path $RegistryPath -Parent
                $newName = split-path $RegistryPath -Leaf
                $result = New-Item -Path $newPath -Name $newName   
            }

            If (Test-Path $RegistryPath) {
                try {$regValue = Get_Reg_Value $RegistryPath $RegistryName}
                catch {$regValue = $null}

                if ($regValue -eq $RegistryValue) {
                    $result = "SUCCESS!! Registry Item: $($RegistryName) already set to $($RegistryValue)"
                    $returnResult = $true
                }
                else {
                    $regValue = New-ItemProperty -Path $RegistryPath -Name $RegistryName -Value $RegistryValue -PropertyType DWORD -Force

                    $regValue = Get_Reg_Value $RegistryPath $RegistryName
                    if ($regValue -eq $RegistryValue) {
                        $result = "SUCCESS!! Registry Item: $($RegistryName) CHANGED to $($RegistryValue)"
                        $returnResult = $true
                    }
                    else
                    {
                        $result = "ERROR!! Registry Item: $($RegistryName) NOT set to $($RegistryValue)"
                    }
                }
            }
            else {
                $result = "Path: $RegistryPath not found and unable to be created."
            }

            if ($LogFile -ne "") {
                $result = writelog "$result ($RegistryPath)" 
            }

            

        } Catch {if ($LogFile -ne "") {$result = writelog "ERROR $currentAction : $_"}}

        return $returnResult

    }

    

    function GBASSetUP {
        
        #Start Log:
            $finalResult = @($TRUE);
            Try {
                $Stamp = (Get-Date).toString("yyyyMMddHHmmss")
            

                #Get Reference to Script Parent Folder, folder with PIX Add-In should be a child.
                #If Calling from exe, psscriptroot will be null
                $installFolder = $PSScriptRoot
                if ($installFolder -eq "") {
                    $installFolder = split-path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -parent 
                }
                $LogFile = "$installFolder\PSLog_$Stamp.txt"
            } Catch {$returnResult = writelog "ERROR $currentAction : $_"; $finalResult += $false }
                        
                                
        #A. Close Any instances of Excel, Word

            #1. Close Excel
                $currentAction = "Closing Excel"
                Try {
                    $excel = Get-Process excel -ErrorAction SilentlyContinue
                    if ($excel -ne $null) {
                        $result = kill -processname excel
                        sleep -Seconds 2
                        $excel = Get-Process excel -ErrorAction SilentlyContinue
                        $result = $excel -eq $null
                        $returnResult = writelog "Killing Excel: $result"
                    }
                    else
                    { $returnResult = writelog "SUCCESS!! Excel not running." }
                } Catch {$returnResult = writelog "ERROR $currentAction : $_"; $finalResult += $false }
            

            #2. Close Word
                $currentAction = "Closing Word"
                Try {
                    $word = Get-Process winword -ErrorAction SilentlyContinue
                    if ($word -ne $null) {
                        $result = kill -processname winword
                        sleep -Seconds 2
                        $word = Get-Process winword -ErrorAction SilentlyContinue
                        $result = $word -eq $null
                        $returnResult = writelog "Killing Word: $result"
                    }
                    else
                    { $returnResult = writelog "SUCCESS!! Word not running." }
                } Catch {$returnResult = writelog "ERROR $currentAction : $_"; $finalResult += $false }
        
        #B. Changing Office Security Settings.
            #1. Change VBA, Protected View Settings:
                #Specific to current version of the application:

                $officeApps = @("Excel", "Word")
                foreach ($officeApp in $officeApps) {

                    #Get Current Version of App:
                    $currentAction = "Getting $officeApp Version"
                    $version = $null
                    $appInstance = $null
                    Try {
                        Switch ($officeApp)
                        {
                            "Excel" {$appInstance = New-Object -ComObject Excel.Application}
                            "Word" {$appInstance = New-Object -ComObject Word.Application}
                        }
                        $version = $appInstance.Version
                        $appInstance.Quit()
                    } Catch {$returnResult = writelog "ERROR $currentAction : $_"; $finalResult += $false }

                    $currentAction = "Changing $officeApp Security Settings"
                    Try {
                        if (($version -ne $null) -and ($officeApp -ne $null)) {                   
                            #a. Enable all Macros:

                                $RegistryPath = "HKCU:\SOFTWARE\Microsoft\Office\$version\$officeApp\Security"
                                $Name         = 'VBAwarnings'
                                $Value        = '1'

                                $finalResult += (Test_Update_Reg $RegistryPath $Name $Value $LogFile)


                            #b. Trust Access to VBA Object Model
                                $Name         = 'AccessVBOM'

                                $finalResult += (Test_Update_Reg $RegistryPath $Name $Value $LogFile)              

                            #c. Un-check Protected View:
                                $RegistryPath = "HKCU:\SOFTWARE\Microsoft\Office\$version\$officeApp\Security\ProtectedView"
                                $Name         = 'DisableInternetFilesInPV'

                                $finalResult += (Test_Update_Reg $RegistryPath $Name $Value $LogFile)  


                                $Name         = 'DisableAttachmentsInPV'
                                $finalResult += (Test_Update_Reg $RegistryPath $Name $Value $LogFile)    

                                $Name         = 'DisableUnsafeLocationsInPV'
                                $finalResult += (Test_Update_Reg $RegistryPath $Name $Value $LogFile) 
                            } 
                        else {
                            $finalResult+= $false; 
                            $returnResult = writelog "ERROR Can't set VBAWarnings, AccessVBOM or ProtectedView without $officeApp Version"
                        }

                    } Catch {$returnResult = writelog "ERROR $currentAction : $_"; $finalResult += $false }
                }

            Try {
            #4. Enable all ActiveX Controls, Disable Safe Mode 
                #(These are set for all Office Applications at once):
                $RegistryPath = 'HKCU:\SOFTWARE\Microsoft\Office\Common\Security'
                $Name         = 'UFIControls'
                
                $finalResult += (Test_Update_Reg $RegistryPath $Name $Value $LogFile)    
                
                #This key may not exist, if it doesn't just ignore.
                $Name         = 'DisableAllActiveX'
                $result = (Test_Update_Reg $RegistryPath $Name '0' $LogFile)    

 
          
            #5. Set Trusted Locations
                $registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey("CurrentUser",[Microsoft.Win32.RegistryView]::Default )
                $registryPath = "SOFTWARE\Microsoft\Office\16.0\Excel\Security\Trusted Locations"
                $registryItem = $registry.OpenSubKey($registryPath)

                $locationPaths = @()
                ForEach ($subKey in $registryItem.GetSubKeyNames())
                    {$locationPaths += $registryItem.OpenSubKey($subKey).GetValue("Path");}

                $addLoc = $locationPaths.count

                $checkPaths =  @("C:\SYLVARApps", "C:\Program Files (x86)\Narrative1")
                $descriptions = @("GBAS VBA Add-Ins", "Narrative1 VBA Add-Ins")

                foreach ($checkPath in $checkPaths) {
                   if ( -not $locationPaths.Contains($checkPath)) {
                      Try {
                          $newName = "$registryPath\Location$addLoc"
                          $index = [array]::IndexOf($checkPaths, $checkPath)

                          $description = $descriptions[$index]

                          $newKey = $registry.CreateSubKey($newName)
                      
                          $newKey.SetValue("Path", $checkPath)
                          $newKey.SetValue("Description", $description)
                          $newKey.SetValue("AllowSubFolders", "1")
                          $addLoc += 1
                      }

                      Catch {
                        $returnResult = writelog "unable to add new key : $checkPath"
                        $registryItem.DeleteSubKey("location$($addLoc)")
                        $addLoc -= 1
                        $finalResult += $false
                      }
                    }
                }
            } Catch {$returnResult = writelog "ERROR $currentAction : $_"; $finalResult += $false}   
           
        #C. Remove Meta Chart Add-in in Excel

            $currentAction = "Removing Meta Chart"
                $addInPath = "$env:APPDATA\Microsoft\AddIns"
                $removePath = "$addInPath\MetaChart"
                $removeFile = "$addInPath\JLL_MetaChart.xlam"
            Try {

                if (test-path $removePath) {
                    Get-ChildItem -Path $removePath -Recurse | Foreach-object {Remove-item -Recurse -path $_.FullName }
                    remove-item $removePath -Force 
                }
                if (test-path $removeFile) {
                    remove-item $removeFile -Force 
                }

                if (((test-path $removePath) -eq $false) -and ((test-path $removeFile) -eq $false)) {
                    $returnResult = writelog "SUCCESS!! Meta-Chart Add-In Removed or didn't exist."
                }
            } Catch {$returnResult = writelog "ERROR $currentAction : $_"; $finalResult += $false }                    

        #D. Install Add-Ins
            #1. Installation Folder

                #$installFolder = "$installFolder\GBAS_SetUp_Files"


            #2. Install PIX Add-In
                $currentAction = "Installing PIX"
                Try {
                    #Determine if PIX Add-In is already installed
                    $RegistryPath = 'HKCU:\SOFTWARE\Microsoft\Office\Excel\Addins\PixXlsAddIn'
                    $pixRegPath = split-path $RegistryPath -NoQualifier
                    $pixRegPath =  $pixRegPath.substring(1)

                    $pixInstall = "$installFolder\PixExcelAddin\"
                    $pixFolder = "$Env:programdata\PiXExcelAddin\"
                    $pixName = "PixXlsAddIn"

                    #Test Registry for PIX AddIn Path
                    if  ((Test-Path $RegistryPath) -eq $false) {
                        $returnResult = writelog "Pix Registry Path not found"

                        $pixAddIns = Get-Package | Where-Object {$_.Name -like 'Pix*'}
                        $pixMachine = $pixAddIns | Where-Object {$_.Name -like '*Machine'}
                        $pixUser = $pixAddIns | Where-Object {$_.Name -notlike '*Machine'}

                            #look for $pixUser app
                            if ($pixUser -ne $null) {
                                #check location, make sure all files exist, if so, point registry there:
                                if ((Test-Path "$pixFolder\$pixName.vsto") -eq $false) {

                                    #copy vsto from install folder
                                    $returnResult = Copy-Item -Path "$pixInstall\$pixName.vsto" -Destination "$pixFolder\$pixName.vsto" -Force
                                    $returnResult = writelog "PIX VSTO Copied: $returnResult"

                                }

                                if (Test-Path "$pixFolder\$pixName.vsto") {
                                      Try {
                                      $returnResult = writelog "PIX Machine Found, adding VSTO File Location to Registry: $pixRegPath"

                                      $newKey = $registry.CreateSubKey($pixRegPath)

                                      $newKey.SetValue("Description", $pixName)
                                      $newKey.SetValue("FriendlyName", $pixName)
                                      $newKey.SetValue("LoadBehavior", "3", [Microsoft.Win32.RegistryValueKind]::DWord)
                                      $newKey.SetValue("Manifest", "file:///$($pixFolder.Replace("\","/"))$pixName.vsto")
                                      }
                                      Catch {
                                        $returnResult = writelog "unable to add new key : $pixRegPath"
                                        $finalResult += $false
                                      }                                
                                }
                                
                            }
                            else {
                                #this doesn't work!
                                #maybe try this!
                                #https://timmyit.com/2016/08/08/sccm-and-powershell-force-installuninstall-of-available-software-in-software-center-through-cimwmi-on-a-remote-client/
                                <#
                                Start-Process -FilePath "$pixFolder\setup.exe"
                                #>

                                #this doesn't seem to change anything
                                <#
                                #run User-PiXExcelAddin.exe
                                $returnResult = writelog "Pix Machine AddIn found, attempting to run User-PixExcelAddin.exe"
                                $pixUserEXE = "$pixFolder\User-PiXExcelAddin.exe"

                                #Pix User
                                if ( $(Test-Path $pixUserEXE)) {
                                    $returnResult = Start-Process $pixUserEXE
                                    $returnResult = writelog "User-PixExcelAddin.exe run"
                                }
                                #>
                            }

                    }

                    else {
                    # if PIX Registry item DOES exist, check where it points to, verify file exists as that location.

                    #if location isn't default location (program data), should we correct or leave it alone?
                    #in that case may need to uninstall & then re-install

                    }

                    #if any of the above failed, just try overwriting everything.
                    $pixAddIns = Get-Package | Where-Object {$_.Name -like 'Pix*'}
                    #$pixMachine = $pixAddIns | Where-Object {$_.Name -like '*Machine'}
                    $pixUser = $pixAddIns | Where-Object {$_.Name -notlike '*Machine'}

                    if (($pixUser -eq $null) -or ((Test-Path $RegistryPath) -eq $false)) {
                        #copy & run setup.exe:
                        #this doesn't work every time.
                        
                        $returnResult = writelog "PIX installation corrupted, attempting install from set up files"
                        <#
                        #a. Copy the PIX into the Excel Add-In 
                            if (test-path $pixInstall) {
                                Copy-Item -Path $pixInstall\* -Destination $pixFolder -Force -Recurse
                            }
                            else {
                                Copy-Item -Path $pixInstall -Destination $pixFolder -Force -Recurse
                            }
                        #>
                        #b. Run the set-up file
                            Start-Process -FilePath "$pixInstall\setup.exe"
                            $finalResult += FindAndCloseVSTOInstaller

                        #c. Check Path:
                            $pixResult = (Test-Path $RegistryPath)
                            if ($pixResult) {$returnResult = writelog "PIX Installed Successfully"}
                            else {$returnResult = writelog "PIX Not Installed Successfully"}
                        
                    }
                } Catch {$returnResult = writelog "ERROR $currentAction : $_"; $finalResult += $false} 
                        
            #3. Install SYLVARapps Add-Ins (GBAS Excel Add-Ins)
                $currentAction = "Installing SYLVARapps Add-Ins"
                
                Try {
                    $sylvarFolder = "C:\SYLVARApps"
                    $sylvarInstall = "$installFolder\SYLVARapps"

                    if (test-path $sylvarFolder) {                       
                        $returnResult = Copy-Item -Path "$sylvarInstall\*" -Destination $sylvarFolder -Recurse -Force
                    }
                    else
                    {                        
                        $returnResult = Copy-Item -Path "$sylvarInstall" -Destination $sylvarFolder -Recurse -Force 
                    }
                    $returnResult = writelog "SUCCESS!! Copied Sylvar Folder $returnResult"

                    #Create shortcut to VASsistant Global in XLSTART:
                    $finalresult += Create_Shortcut "$sylvarFolder\JLL_XL_VASsistantGlobal.xlam" "$env:appdata\Microsoft\Excel\XLSTART"
                    
                } Catch {$returnResult = writelog "ERROR $currentAction : $_"; $finalResult += $false} 
                                
                
            #4. Install LinkManager (GBAS Word Add-Ins)
                $currentAction = "Installing LinkManager"
                Try {
                    $linkFolder = "$Env:appdata\Microsoft\Word\STARTUP"
                    $linkInstall = "$installFolder\STARTUP"

                    if (test-path $linkFolder) {
                        $returnResult = Copy-Item -Path "$linkInstall\*" -Destination $linkFolder -Recurse -Force
                    }
                    else {
                        $returnResult = Copy-Item -Path "$linkInstall" -Destination $linkFolder -Recurse -Force 
                    } 
                    $returnResult = writelog "SUCCESS!! Copied Link Manager $returnResult"
                } Catch {$returnResult = writelog "ERROR $currentAction : $_"; $finalResult += $false} 
                  
                              



        #E. Finalize, Delete LogFile if everything was successful. Otherwise, show logfile, have user contact PIX Support. ###Start LogFile, Delete Logfile
        $finalResult = $(-not $finalResult.Contains($false))

        if ($finalResult) {
            Try {
                #start $LogFile
                #sleep -Seconds 2
                Remove-Item $LogFile -Force 
                return "All Changes Made Successfully!"
            } Catch {}
        }
        else {
            #Try {start $LogFile} Catch {} 
            return "An error has occurred. Please contact PIX or GBAS Support."

        }
    }

  
        Add-Type -assembly System.Windows.Forms
        $main_form = New-Object System.Windows.Forms.Form
        $main_form.Text ='GBAS User Set Up'

        $main_form.Width = 500

        $main_form.Height = 100

        # label
        $objLabel = New-Object System.Windows.Forms.label
        $objLabel.Location = New-Object System.Drawing.Size(7,10)
        $objLabel.Size = New-Object System.Drawing.Size(400,80)
        $objLabel.Text = "Installation in progress. Please Wait... "
        $objLabel.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Bold)

        $main_form.Controls.Add($objLabel)

        $main_form.MaximizeBox = $False
        $main_form.MinimizeBox = $False
        $main_form.ControlBox = $False        

        $main_form.Show()

        $finalResult = GBASSetUP 
        $main_form.Hide()

        $objLabel.Text = $finalResult    
        $main_form.ControlBox = $true
        $main_form.ShowDialog()
        
        
     


