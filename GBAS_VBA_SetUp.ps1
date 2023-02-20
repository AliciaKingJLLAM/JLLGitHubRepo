


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
        
            
        $currentAction = "Updating Registry - $RegistryName"
        Try {
            IF (Test-Path $RegistryPath) {
                $regValue = Get_Reg_Value $RegistryPath $RegistryName 

                if ($regValue -eq $RegistryValue) {
                    $result = "Registry Item: $($RegistryName) already set to $($RegistryValue)"
                    $returnResult = $true
                }
                else {
                    $regValue = New-ItemProperty -Path $RegistryPath -Name $RegistryName -Value $RegistryValue -PropertyType DWORD -Force -ErrorAction Suspend

                    $regValue = Get_Reg_Value $RegistryPath $RegistryName
                    if ($regValue -eq $RegistryValue) {
                        $result = "SUCCESS!! Registry Item: $($RegistryName) CHANGED to $($RegistryValue)"
                        $returnResult = $true
                    }
                    else
                    {
                        $result = "ERROR!! Registry Item: $($RegistryName) NOT set to $($RegistryValue)"
                        $returnResult = $false
                    }
                }
            }
            else {
                $result = "Path: $RegistryPath not found."
                $returnResult = $false
            }

            if ($LogFile -ne "") {
                $returnResult = writelog "$result ($RegistryPath)" 
            }

            return $returnResult

        } Catch {if ($LogFile -ne "") {$returnResult = writelog "ERROR $currentAction : $_"}}
    }


    function GBASSetUP {
        
        ### Close Any instances of Excel:
            $currentAction = "Closing Excel"
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
                    
            $excel = Get-Process excel -ErrorAction SilentlyContinue
            if ($excel -ne $null) {
                $result = kill -processname excel
                $returnResult = writelog "Killing Excel: $result"
            }
            else
            { $returnResult = writelog "Excel not running." }
            
            } Catch {$returnResult = writelog "ERROR $currentAction : $_"; $finalResult += $false }
            
        
        #A. Remove Meta Chart Add-in in Excel

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
                    $returnResult = writelog "Meta-Chart Add-In Removed or didn't exist."
                }
            } Catch {$returnResult = writelog "ERROR $currentAction : $_"; $finalResult += $false }

        #B. Changing Excel Security Settings.
            $currentAction = "Changing Excel Security Settings"
            Try {
            #1. Enable all Macros:
                $RegistryPath = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Excel\Security'
                $Name         = 'VBAwarnings'
                $Value        = '1'

                $finalResult += (Test_Update_Reg $RegistryPath $Name $Value $LogFile)



            #2. Trust Access to VBA Object Model
                $Name         = 'AccessVBOM'

                $finalResult += (Test_Update_Reg $RegistryPath $Name $Value $LogFile)              


            #3. Enable all ActiveX Controls, Disable Safe Mode
                $RegistryPath = 'HKCU:\SOFTWARE\Microsoft\Office\Common\Security'
                $Name         = 'UFIControlsX'
                
                $finalResult += (Test_Update_Reg $RegistryPath $Name $Value $LogFile)    
                

            #4. Un-check Protected View:
                $RegistryPath = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Excel\Security\ProtectedView'
                $Name         = 'DisableInternetFilesInPV'

                $finalResult += (Test_Update_Reg $RegistryPath $Name $Value $LogFile)  


                $Name         = 'DisableAttachmentsInPV'
                $finalResult += (Test_Update_Reg $RegistryPath $Name $Value $LogFile)    

                $Name         = 'DisableUnsafeLocationsInPV'
                $finalResult += (Test_Update_Reg $RegistryPath $Name $Value $LogFile)  
          
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

        #c. Install Add-Ins
            #1. Installation Folder

                $installFolder = "$installFolder\GBAS_SetUp_Files"


            #2. Install PIX Add-In
                $currentAction = "Changing Excel Security Settings"
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

                        #Check if PIX is already installed
                        if ($pixMachine -ne $null) {

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

                    }

                    else {
                    # if PIX Registry item DOES exist, check where it points to, verify file exists as that location.

                    #if location isn't default location (program data), should we correct or leave it alone?
                    #in that case may need to uninstall & then re-install

                    }

                    #if any of the above failed, just try overwriting everything.
                    $pixAddIns = Get-Package | Where-Object {$_.Name -like 'Pix*'}
                    $pixMachine = $pixAddIns | Where-Object {$_.Name -like '*Machine'}
                    $pixUser = $pixAddIns | Where-Object {$_.Name -notlike '*Machine'}

                    if (($pixMachine -eq $null) -or ($pixUser -eq $null) -or ((Test-Path $RegistryPath) -eq $false)) {
                        $returnResult = writelog "Problem installing PIX"
                        $finalResult += $false
                        #copy & run setup.exe:
                        #this doesn't work
                        <#
                        $returnResult = writelog "PIX installation corrupted, installing from set up files"

                        #a. Copy the PIX into the Excel Add-In 
                            if (test-path $pixInstall) {
                                Copy-Item -Path $pixInstall\* -Destination $pixFolder -Force -Recurse
                            }
                            else {
                                Copy-Item -Path $pixInstall -Destination $pixFolder -Force -Recurse
                            }

                        #b. Run the set-up file
                            Start-Process -FilePath "$pixFolder\setup.exe"
                        #c. Check Path:
                            $pixResult = (Test-Path $RegistryPath)
                            if ($pixResult) {$returnResult = writelog "PIX Installed Successfully"}
                            else {$returnResult = writelog "PIX Not Installed Successfully"}
                        #>
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
                    #$returnResult = writelog "Copied Sylvar Folder $returnResult"
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
                    $returnResult = writelog "Copied Link Manager $returnResult"
                } Catch {$returnResult = writelog "ERROR $currentAction : $_"; $finalResult += $false} 
                  
                              
#    FIGURE OUT WHAT HAPPENS WHEN SOMETHING GOES WRONG HERE: , DOES TRY/CATCH WORK?                           
            #5. Look for N1 Excel AddIns, if any don't exist, copy to Excel AddIns Folder as temporary fix:
                $currentAction = "Installing N1 Excel AddIns"
                Try {
                    $n1Folder = "C:\Program Files (x86)\Narrative1"
                    $n1Install = "$installFolder\AddIns"

                    if ((test-path $n1Folder) -eq $false) {
                        #copy the entire folder
                            if (test-path $addInPath) {
                                $returnResult = Copy-Item -Path "$n1Install\*" -Destination $addInPath -Recurse -Force 
                            }
                            else {
                                $returnResult = Copy-Item -Path "$n1Install" -Destination $addInPath -Recurse -Force
                            } 
                            $returnResult = writelog "Copied N1 AddIns $returnResult"
                    }
                    else {
                        #copy file-by-file if necessary:
                        $n1Addins = Get-ChildItem $n1Install -name
                        foreach ($n1AddIn in $n1Addins) {
                            if (-not (test-path $n1Folder\$n1AddIn)) {
                                $returnResult = Copy-Item -Path "$n1Install\$n1Addin" -Destination "$addInPath\$n1AddIn" 
                                $returnResult = writelog "Copied N1 AddIn: $n1AddIn"
                            }
                        }
                    }
                } Catch {$returnResult = writelog "ERROR $currentAction : $_"; $finalResult += $false} 


        #D. Finalize, Delete LogFile if everything was successful. Otherwise, show logfile, have user contact PIX Support. ###Start LogFile, Delete Logfile
        $finalResult = $(-not $finalResult.Contains($false))
        $returnResult = writelog "All Changes Made Successfully?? $finalResult"

        if ($finalResult) {
            Try {
                #start $LogFile
                #sleep -Seconds 2
                Remove-Item $LogFile -Force 
                "All Changes Made Successfully!"
            } Catch {}
        }
        else {
            #Try {start $LogFile} Catch {} 
            "An error has occurred. Please contact PIX or GBAS Support"

        }
    }
"YOU NEED TO FIX THIS TO MATCH CANADA!"

    GBASSetUP 



