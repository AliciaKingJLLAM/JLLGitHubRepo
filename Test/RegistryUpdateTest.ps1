                    function WriteLog
    {
        Param (
            [string]$LogString
            )

        $LogPath = Split-Path $LogFile -Parent

        #Create the Sylvar Apps Folder if it doesn't exist
        if ( $(Test-Path $LogPath) -eq $false) {
            New-Item -ItemType Directory -Force -Path $LogPath
        }

        $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
        $LogMessage = "$Stamp`t$LogString"

        Add-content $LogFile -value $LogMessage
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
        
            IF (Test-Path $RegistryPath) {
                $regValue = Get_Reg_Value $RegistryPath $RegistryName

                if ($regValue -eq $RegistryValue) {
                    $result = "Registry Item: $($RegistryName) already set to $($RegistryValue)"
                }
                else {
                    $regValue = New-ItemProperty -Path $RegistryPath -Name $RegistryName -Value $RegistryValue -PropertyType DWORD -Force 

                    $regValue = Get_Reg_Value $RegistryPath $RegistryName
                    if ($regValue -eq $RegistryValue) {
                        $result = "SUCCESS!! Registry Item: $($RegistryName) CHANGED to $($RegistryValue)"
                    }
                    else
                    {
                        $result = "ERROR!! Registry Item: $($RegistryName) NOT set to $($RegistryValue)"
                    }
                }
            }
            else {
                $result = "Path: $RegistryPath not found."
            }

            if ($LogFile -ne "") {
                writeLog "$result ($RegistryPath)"
            }

            return $result
    }

    function Delete_Old_Logs {

    #Delete any old logs before creating a new one
    }

            $Stamp = (Get-Date).toString("yyyyMMddHHmmss")
        $LogFile = "C:\SYLVARapps\PSLog_$Stamp.txt"
                
                $RegistryPath = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Excel\Security'
                $Name         = 'VBAwarnings'
                $Value        = '0'

                Test_Update_Reg $RegistryPath $Name $Value $LogFile

                $Value        = '1'
                Test_Update_Reg $RegistryPath $Name $Value $LogFile