Add-Type –AssemblyName UIAutomationClient
Add-Type –AssemblyName UIAutomationTypes



function FindAndCloseVSTOInstaller () {

    $result = $false

        $timer = [Diagnostics.Stopwatch]::StartNew()
        $maxWaitSeconds = 1

        Do {$installer = get-process VSTOInstaller
            if ($timer.Elapsed.TotalSeconds -ge $maxWaitSeconds) {break}
        }
        while ($installer -eq $null)

        Try {
            $installerID = $installer.Id
            $root = [Windows.Automation.AutomationElement]::RootElement
            $condition = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ProcessIdProperty, $installerID)
            $installerUI = $root.FindFirst([Windows.Automation.TreeScope]::Children, $condition)

        $installerTexts = $installerui.FindAll([Windows.Automation.TreeScope]::Descendants, `
                    (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, `
                    [System.Windows.Automation.ControlType]::Text)))
        
        $result = $true

        foreach ($installerText in $installerTexts) {
            if ($installerText.current.Name -like "*error*") {
                $result = $false
            }
        }

        $installerButtons = $installerui.FindAll([Windows.Automation.TreeScope]::Descendants, `
                    (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, `
                    [System.Windows.Automation.ControlType]::Button)))

        foreach ($installerButton in $installerButtons) {
            if ($installerButton.current.Name -eq "Close") {
                $closeResult = $installerbutton.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
            }
        }

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
