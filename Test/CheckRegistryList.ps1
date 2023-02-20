

        $registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey("CurrentUser",[Microsoft.Win32.RegistryView]::Default )

        $registryPath = "SOFTWARE\Microsoft\Office\16.0\Excel\Security\Trusted Locations"
        $registryItem =$registry.OpenSubKey($registryPath)

        $locationPaths = @()
        ForEach ($subKey in $registryItem.GetSubKeyNames())
        {
            $locationPaths += $registryItem.OpenSubKey($subKey).GetValue("Path");
        }

        $addLoc = $locationPaths.count

        $checkPaths =  @("C:\SYLVARApps", "C:\Program Files (x86)\Narrative1", "C:\Program Files\Microsoft Office\root\Office16\XLSTART\")

        foreach ($checkPath in $checkPaths) {
           if ( -not $locationPaths.Contains($checkPath)) {
              
              
              Try {
              $newName = "$registryPath\Location$addLoc"

              $newKey = $registry.CreateSubKey($newName)
              $newKey.SetValue("Path", $checkPath)
              $newKey.SetValue("Description", "JLLT Core GASM VBA Add-Ins")
              $newKey.SetValue("AllowSubFolders", "1")
              $addLoc += 1
              }

              Catch {
                
                "unable to add new key : $checkPath"
                $registryItem.DeleteSubKey("location$($addLoc)")
                $addLoc -= 1
              }
              
            }
            else
            {
                "found : $checkpath"
            }
        }

 