

    function WriteLog
    {
        Param (
            [string]$LogString
            )

        Try {
            $LogPath = Split-Path $LogFile -Parent

            #Create the Log file Folder if it doesn't exist
            if ( $(Test-Path $LogPath) -eq $false) {
                New-Item -ItemType Directory -Force -Path $LogPath -ErrorAction Suspend
            }

            $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
            $LogMessage = "$Stamp`t$LogString"

            Add-content $LogFile -value $LogMessage 
            return $true
        }
        Catch {
            $_
            return $false
        }
    }

    function Get_Reg_Value {
        param (
            [string]$RegistryPath,
            [string]$RegistryName)
         return Get-ItemPropertyValue -path $RegistryPath -name $RegistryName -ErrorAction Suspend
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

            $Stamp = (Get-Date).toString("yyyyMMddHHmmss")
            

            #Get Reference to Script Parent Folder, folder with PIX Add-In should be a child.
            #If Calling from exe, psscriptroot will be null
            $installFolder = $PSScriptRoot
            if ($installFolder -eq "") {
                $installFolder = split-path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -parent 
            }
            $LogFile = "$installFolder\PSLog_$Stamp.txt"

            writeLog "AN ERROR HAS OCCURRED. Please send this log to PIX or Global ASM Support. `n`n"

    }

    GBASSetUP 



