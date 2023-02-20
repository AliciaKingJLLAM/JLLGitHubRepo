

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

 

        ### Close Any instances of Excel:
        
        start-process excel
        start-process excel
        start-process excel

        
        #A. Remove Meta Chart Add-in in Excel
                $addInPath = "$env:APPDATA\Microsoft\AddIns"
                $removePath = "$addInPath\MetaChart"
                $removeFile = "$addInPath\JLL_MetaChart.xlam"

            New-Item -ItemType Directory -Force -Path $removePath
            New-Item -ItemType File -Force -Path $removeFile




        #B. Changing Excel Security Settings.
            #1. Enable all Macros:
                $RegistryPath = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Excel\Security'
                $Name         = 'VBAwarnings'
                $Value        = '0'

                $returnResult = Test_Update_Reg $RegistryPath $Name $Value $LogFile


            #2. Trust Access to VBA Object Model
                $Name         = 'AccessVBOM'

                $returnResult = Test_Update_Reg $RegistryPath $Name $Value $LogFile   
               

            #3. Enable all ActiveX Controls, Disable Safe Mode
                $RegistryPath = 'HKCU:\SOFTWARE\Microsoft\Office\Common\Security'
                $Name         = 'UFIControlsX'
                
                $returnResult = Test_Update_Reg $RegistryPath $Name $Value $LogFile    
                

            #4. Un-check Protected View:
                $RegistryPath = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Excel\Security\ProtectedView'
                $Name         = 'DisableInternetFilesInPV'

                $returnResult = Test_Update_Reg $RegistryPath $Name $Value $LogFile    


                $Name         = 'DisableAttachmentsInPV'
                $returnResult = Test_Update_Reg $RegistryPath $Name $Value $LogFile    

                $Name         = 'DisableUnsafeLocationsInPV'
                $returnResult = Test_Update_Reg $RegistryPath $Name $Value $LogFile  
          
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
                   if ($locationPaths.Contains($checkPath)) {
                       "Removing HKCU:\$checkPath"
                       #Remove-Item -Path "HKCU:\$checkPath" 
                       
                    }
                }



                $sylvarFolder = "C:\SYLVARApps"

                if (test-path $sylvarFolder) {
                    Rename-Item $sylvarFolder -NewName $sylvarFolder.Insert($sylvarFolder.Length,"_$(Get-Date -format 'yyyyMMddTHHmmssffffZ')")
                }

                
                
                
            #4. Install LinkManager (GBAS Word Add-Ins)
                $linkFolder = "$Env:appdata\Microsoft\Word\STARTUP"

                if (test-path $linkFolder) {
                    Rename-Item $linkFolder -NewName $linkFolder.Insert($linkFolder.Length, "_$(Get-Date -format 'yyyyMMddTHHmmssffffZ')")
                }


                Invoke-Item (split-path $sylvarFolder -parent)
                Invoke-Item (split-path $linkFolder -parent)
                Invoke-Item $addInPath  
                Invoke-Item "C:\Program Files (x86)\Narrative1"