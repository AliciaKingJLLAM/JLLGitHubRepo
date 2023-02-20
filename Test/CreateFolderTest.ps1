      
      
              
              $LogFile = "C:\SYLVARappsTEST\PSLog_$Stamp.txt"
              $LogPath = Split-Path $LogFile -Parent


        #Create the Sylvar Apps Folder if it doesn't exist
        if ( $(Test-Path $LogPath) -eq $false) {
            New-Item -ItemType Directory -Force -Path $LogPath
        }


