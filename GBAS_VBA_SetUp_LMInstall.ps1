


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

    Function Get-Folder{
        param (
            [string]$initialDirectory = "" 
           ,[string]$description = "Select a folder" 
        )

    #Add-Type -AssemblyName System.Windows.Forms

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    #$foldername.InitialDirectory  = $initialDirectory
    $foldername.Description = $description
    #$foldername.rootfolder = $initialDirectory #"MyComputer"
    $foldername.SelectedPath = $initialDirectory

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}

    function GBASSetUP {
        
        ### Close Any instances of Excel:
            $currentAction = "Closing Word"
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
            
                    
                $excel = Get-Process winword -ErrorAction SilentlyContinue
            
                if ($excel -ne $null) {
                    $result = kill -processname winword
                    $returnResult = writelog "Killing Word: $result"
                }
                else {$returnResult = writelog "Word not running."}
            
            } Catch {$returnResult = writelog "ERROR $currentAction : $_"; $finalResult += $false }
            

        #c. Install Add-Ins
            #1. Installation Folder
                #$installFolder = "$installFolder\GBAS_SetUp_Files"
                $installFolder = "$env:USERPROFILE\Box"

                if (-not(test-path $installFolder)) {
                    $resultFolder = Get-Folder "$env:USERPROFILE\Contacts" "Please select the box drive folder location on your machine"
                    $s = '\Box\'

                    if ($resultFolder -like "*$s*") {
                        $r = $resultFolder.IndexOf($s)
                        $installFolder  = $resultFolder.substring(0, $r+$s.Length-1)
                    }
                    else {$installFolder = "ERROR"}
                }

                if ($installFolder -eq "ERROR") {$finalResult+= $false}
                else {
                    $installFolder = "$installFolder\VAS - Master\Templates\Link Manager\DO NOT USE\UAT"               

                    #4. Install LinkManager (GBAS Word Add-Ins)
                        $currentAction = "Installing LinkManager"
                        Try {
                            $linkFolder = "$Env:appdata\Microsoft\Word\STARTUP"
                            $linkFile = "JLL_Word_Reporting.dotm"
                            $linkInstall = $installFolder

                            $timeString = "Archive_$(Get-Date -format 'yyyyMMddTHHmmssffffZ')"
                            $itemCopied = $false

                            if (-not (test-path $linkFolder)) {$finalresult += new-item -ItemType Directory -Path $linkFolder}
                            if (-not (test-path "$linkFolder\$timeString")) {$finalresult += new-item -ItemType Directory -Path "$linkFolder\$timeString"}
                            if ((test-path "$linkFolder\$linkFile")) {$itemCopied = $true; $finalresult += Move-Item -Path "$linkFolder\$linkFile" -Destination "$linkFolder\$timeString\$linkFile"}
                            if ((test-path "$linkFolder\$linkFile.lnk")) {$itemCopied = $true; $finalresult += Move-Item -Path "$linkFolder\$linkFile.lnk" -Destination "$linkFolder\$timeString\$linkFile.lnk"}
                            if (-not($itemCopied)) {remove-item  "$linkFolder\$timeString" -Force}
                            else {
                                if (-not (test-path "$linkFolder\Archive")) {$finalresult += new-item -itemtype Directory -path "$linkFolder\Archive"}
                                $finalResult+= move-item "$linkFolder\$timeString" -Destination "$linkFolder\Archive\$timeString" -Force}

                            copy-item -path "$installFolder\$linkFile" -Destination "$linkFolder\$linkFile"

                            $returnResult = writelog "Copied Link Manager $returnResult"
                        } Catch {$returnResult = writelog "ERROR $currentAction : $_"; $finalResult += $false} 
                    
                }
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


    GBASSetUP 





