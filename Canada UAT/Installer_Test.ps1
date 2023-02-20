Add-Type –AssemblyName UIAutomationClient
Add-Type –AssemblyName UIAutomationTypes

$installFolder = $PSScriptRoot
$installFolder = "$installFolder\GBAS_SetUp_Files"
                    $pixInstall = "$installFolder\PixExcelAddin\"
                    $pixFolder = "$Env:programdata\PiXExcelAddin\"
                    $pixName = "PixXlsAddIn"
Start-Process -FilePath "$pixInstall\setup.exe"

    $result = $false

        $timer = [Diagnostics.Stopwatch]::StartNew()
        $maxWaitSeconds = 10

        Do {$installer = get-process VSTOInstaller -ErrorAction SilentlyContinue
            if ($timer.Elapsed.TotalSeconds -ge $maxWaitSeconds) {break}
        }
        while ($installer -eq $null)


        if ($installer -eq $null) {
             "ERROR: Cannot find VSTO Installer"
            return $false
        }

        #Try {
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

            foreach ($installerButton in $installerButtons) {
                if ($installerButton.current.Name -eq "Cancel") {break}

                if ($installerButton.current.Name -eq "Close") {
                    $closeButton = $installerButton
                    $closeResult = $installerbutton.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
                    $result = $True
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

        


        <#
	    $condition1 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ClassNameProperty, "Button")
	    $condition2 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, $name)
	    $condition = New-Object Windows.Automation.AndCondition($condition1, $condition2)
	    $button = $calcUI.FindFirst([Windows.Automation.TreeScope]::Descendants, $condition)
	    $button.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
        #>

        
        # }         Catch {writelog "ERROR Closing PIX Installer: $_"}


