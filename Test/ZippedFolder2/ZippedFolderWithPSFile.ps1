                #Get Reference to Script Parent Folder, folder with PIX Add-In should be a child.
                $installFolder = $PSScriptRoot

                #Unzip Folder:
                $returnResult = Expand-Archive -Path "$installFolder.zip" -DestinationPath $installFolder