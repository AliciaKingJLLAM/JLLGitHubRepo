

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


                              



        #E. Finalize, Delete LogFile if everything was successful. Otherwise, show logfile, have user contact PIX Support. ###Start LogFile, Delete Logfile
        $finalResult = $(-not $finalResult.Contains($false))

        if ($finalResult) {
            Try {
                #start $LogFile
                #sleep -Seconds 2
                Remove-Item $LogFile -Force 
                "All Changes Made Successfully!"
            } Catch {}
        }
        else {
            Try {start $LogFile} Catch {} 
            "An error has occurred. Please contact PIX or GBAS Support"

        }
    }

    GBASSetUP 



