
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


        function ResetLinkManager {
        
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
                        
                                
        #A. Close Any instances of  Word

            #1. Close Word
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
        
                         

        #D. Reset LinkManager

            #4. Install LinkManager (GBAS Word Add-Ins)
                $currentAction = "Installing LinkManager"
                Try {
                    $linkFolder = "$Env:appdata\Microsoft\Word\STARTUP"
                    $linkMgr = "JLL_Word_Reporting.dotm"
                    $sylvarApps = "C:\SYLVARApps"

                    #Remove any previously created folders:
                    Do {
                        $prevFolder = $null
                        $prevFolder = get-childitem $linkfolder -directory # -ErrorAction SilentlyContinue

                        if ($prevFolder -ne $null) {
                            $prevFolder = "$linkfolder\$($prevFolder[0].Name)"
                            $finalResult += copy-item -Path "$prevFolder\$linkMgr" -Destination $linkFolder -Force
                            $returnResult = copy-item -Path "$prevFolder\$linkMgr" -Destination $sylvarApps -Force
                            $removeFolder = remove-item $prevFolder -Force -Recurse
                            $finalResult +=  writelog "Removed $prevFolder"
                        }

                    } while ($prevFolder -ne $null)

                     
                    #Create a new folder
                    $newPath = "$linkFolder\$((Get-Date).toString('yyyyMMddHHmmss'))"
                    $newFolder = new-item -path $newPath -itemtype Directory

                    #copy link manager to new folder
                    $finalResult +=  copy-item -Path "$linkFolder\$linkMgr" -destination $newPath -Force
                    $finalResult +=  remove-item -Path "$linkFolder\$linkMgr"  -Force
                    $finalResult +=  writelog "SUCCESS!! Copied Link Manager $returnResult"

                    #create shortcut
                    $finalResult +=  Create_Shortcut "$newPath\$linkMgr" $linkFolder
                    
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
            Try {start $LogFile} Catch {} 
            #return "An error has occurred. Please contact PIX or GBAS Support."

        }
    }

    ResetLinkManager


  <#
        Add-Type -assembly System.Windows.Forms
        $main_form = New-Object System.Windows.Forms.Form
        $main_form.Text ='GBAS User Set Up'

        $main_form.Width = 500

        $main_form.Height = 100

        # label
        $objLabel = New-Object System.Windows.Forms.label
        $objLabel.Location = New-Object System.Drawing.Size(7,10)
        $objLabel.Size = New-Object System.Drawing.Size(400,80)
        $objLabel.Text = "Work in progress. Please Wait... "
        $objLabel.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Bold)

        $main_form.Controls.Add($objLabel)

        $main_form.MaximizeBox = $False
        $main_form.MinimizeBox = $False
        $main_form.ControlBox = $False        

        $main_form.Show()

       # $finalResult = ResetLinkManager 
        $main_form.Hide()

        $objLabel.Text = $finalResult    
        $main_form.ControlBox = $true
        $main_form.ShowDialog()
        
        
     


#>