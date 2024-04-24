Function Get-PatchdataFromMicrosoft {
    <#
        Downloads data from Microsofts websites for Update History. Uses class 'supLeftNavLink' to extract the string we want to process.
            Examples :  Clean up string from "January 17, 2022—KB5010793 (OS Builds 19042.1469, 19043.1469, and 19044.1469) Out-of-band"/ "January 11, 2022—KB5009557 (OS Build 17763.2452)"
                        into something we can use as an object. Replaces known characters into ";" to use as a delimiter.
        
        "Builds" is always treated as an array to make it possible to create objects for each Windows 10-version based on the same core operating system version.
                
            KBDate = January 17, 2022
            KB = KB5010793s
            Builds = 19042.1469, 19043.1469, 19044.1469
            Type = Out-of-band

            KBDate = January 11, 2022
            KB = KB5009557
            Builds = 17763.2452
            Type = Patch Tuesday
                
    Version 1.2
        Fixed output where the original string contained "and" or ",", which resulted in builds being bunched together.
        Should also not output duplicates any more
    Version 1.1
        Fixed freeze from using internal IE-parsing in Invoke-WebRequest. Uses -usebasicparsing and selects "supLeftNavLink" class and removes irrelevant information without using IE-engine.
        Should now be Powershell 7+-compatible.
    Version 1.1.1
        Added Windows Server 23H2 and Windows 11 23H2
    #>
    $OldProgressPreference = $ProgressPreference
    $ProgressPreference = "SilentlyContinue"
    $URIs = @("https://support.microsoft.com/en-us/topic/windows-10-update-history-857b8ccb-71e4-49e5-b3f6-7073197d98fb",          # Windows 10 / Server 2016 / 2019
            "https://support.microsoft.com/en-gb/topic/windows-server-2022-update-history-e1caa597-00c5-4ab9-9f3e-8212fe80b2ee", # Windows Server 2022
            "https://support.microsoft.com/en-gb/topic/windows-server-version-23h2-update-history-68c851ff-825a-4dbc-857b-51c5aa0ab248", # Windows Server 23H2
            "https://support.microsoft.com/en-us/topic/windows-11-update-history-a19cd327-b57f-44b9-84e0-26ced7109ba9")          # Windows 11)
    $MSOutput = @()
    foreach ($URI in $URIs){
        $result = Invoke-WebRequest -Uri $URI -UseBasicParsing
        foreach($r in $result.Links){
            try {
                $tmpOutput = [System.Net.WebUtility]::HtmlDecode( $((($r |Where-Object {$_.class -eq "supLeftNavLink" }).OuterHTML | ForEach-Object {$_.split(">") | Select-Object -last 2 | Where-Object {$_.trim() -ne "" }} ).trim("</a") ) )
                $msOutput += $tmpOutput | Where-Object {$_ -like "*—*"} | Where-Object {$_ -notlike "*Windows 10 Mobile*"}
                    }
            catch {}

        }
        
    }
    $result = @()
    foreach ($string in $MSOutput){
    $builds = (([regex]::Matches($string, '\((.*?)\)').Value).replace("OS Build OS ","").replace("OS Builds ","").replace("OS Build ","").replace("(","").replace(")","").replace(", ",";").replace(" and ",";").replace("and ","")).split(";")

        foreach($build in $builds){
            $ClientOS = $null
            $ServerOS = $null
            switch (([version]$build).major)
            { 
                10240 {$ClientOS = "Windows 10 1507"     }
                10586 {$ClientOS = "Windows 10 1511"     }
                14393 {$ClientOS = "Windows 10 1607"
                    $ServerOS = "Windows Server 2016" }
                15063 {$ClientOS = "Windows 10 1703"     }
                15254 {$ClientOS = "Windows 10 Mobile"   }
                16299 {$ClientOS = "Windows 10 1709"     }
                17134 {$ClientOS = "Windows 10 1803"     }
                17763 {$ServerOS = "Windows Server 2019"
                    $ClientOS = "Windows 10 1809"     } 
                18362 {$ClientOS = "Windows 10 1903"     }
                18363 {$ClientOS = "Windows 10 1909"     }
                19041 {$ClientOS = "Windows 10 2004"     }
                19042 {$ClientOS = "Windows 10 20H2"     }
                19043 {$ClientOS = "Windows 10 21H1"     }
                19044 {$ClientOS = "Windows 10 21H2"     }
                19045 {$ClientOS = "Windows 10 22H2"     }
                20348 {$ServerOS = "Windows Server 2022" } 
                22000 {$ClientOS = "Windows 11 21H2"     } 
                22621 {$ClientOS = "Windows 11 22H2"     }
                22631 {$ClientOS = "Windows 11 23H2"     }
                25398 {$ServerOS = "Windows Server 23H2" }
                }
            $string = ($string.replace($([regex]::Matches($string, '\((.*?)\)').Value),";").replace("—",";").replace(" ; ",";").replace(" ;",";")).split(";")
            if ($string[2]){
                $PatchType = $string[2]
                } else {$PatchType = "Patch Tuesday"}
            $hash = [ordered]@{
                KB        = $string[1].replace(" ","")
                Date      = get-date $string[0]
                ClientOS  = $ClientOS
                ServerOS  = $ServerOS
                Build     = [version]$build
                PatchType = $PatchType

            }
            $result += New-Object PSObject -Property $hash
        }
    }
    $ProgressPreference = $OldProgressPreference
    return $($result | select * | Sort-Object KBDate,KB,Build | Get-Unique -AsString)
}
