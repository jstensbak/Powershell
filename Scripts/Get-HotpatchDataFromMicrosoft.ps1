Function Get-HotpatchDataFromMicrosoft {
    <#
        Downloads data from Microsoft's websites for Update History. Uses class 'supLeftNavLink' to extract the string we want to process.
            Examples:  Clean up string from "January 17, 2022—KB5010793 (OS Builds 19042.1469, 19043.1469, and 19044.1469) Out-of-band"/ "January 11, 2022—KB5009557 (OS Build 17763.2452)"
                        into something we can use as an object.
        
        Each build number gets its own entry in the resulting array, making it possible to create objects for each Windows version
        based on the same core operating system version.
                
        Example output format:
            KB = KB5010793
            Build = 19042.1469
            Date = January 17, 2022
            ClientOS = Windows 10 20H2
            ServerOS = $null
            Type = Out-of-band

            KB = KB5009557
            Build = 17763.2452
            Date = January 11, 2022
            ClientOS = Windows 10 1809
            ServerOS = Windows Server 2019
            Type = Patch Tuesday
    
    Version 1.0
        Initial version, based on Get-PatchdataFromMicrosoft v 2.0.1
        This function is a rewrite of the original Get-PatchDataFromMicrosoft function, which was designed to parse update history from Microsoft's websites.
    #>
    [CmdletBinding()]
    Param()

    Begin {
        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = "SilentlyContinue"
        
        $URIs = @(
            "https://support.microsoft.com/en-US/topic/release-notes-for-hotpatch-on-windows-11-version-24h2-enterprise-clients-c0906ee6-5e62-498f-bd5a-8f4966349f3c" # Windows 11 Hotpatch 24H2
        )
        
        # Regex pattern to match update information
        $regex = [regex]::new('^(?<date>.+?)—(?<patchtype>Hotpatch|Baseline)(?:\s+(?<kb>KB\d+))?\s+\(.*?(?<builds>[\d\.]+).*?\)$')
        
        # Initialize result array
        $result = @()
    }
    
    Process {
        # Collect update information from all Microsoft pages
        $MSOutput = @()
        foreach ($URI in $URIs) {
            try {
                $webResult = Invoke-WebRequest -Uri $URI -UseBasicParsing
                foreach ($link in $webResult.Links) {
                    try {
                        # Extract the text from supLeftNavLink class elements
                        if ($link.class -eq "supLeftNavLink") {
                            $htmlText = $link.OuterHTML
                            $textParts = $htmlText.split(">") | Select-Object -last 2 | Where-Object { $_.trim() -ne "" }
                            $cleanText = ($textParts).trim("</a")
                            $decodedText = [System.Net.WebUtility]::HtmlDecode($cleanText)
                            
                            # Only add entries that contain the em dash and aren't Windows 10 Mobile
                            if ($decodedText -like "*—*" -and $decodedText -notlike "*Windows 10 Mobile*") {
                                $MSOutput += $decodedText
                            }
                        }
                    }
                    catch {
                        # Continue to next link if there's an error with this one
                        continue
                    }
                }
            }
            catch {
                Write-Warning "Failed to retrieve update data from $URI. $_"
                continue
            }
        }
        
        # Process each update string collected from the websites
        foreach ($line in $MSOutput) {
            $match = $regex.Match($line)
            if ($match.Success) {
                $date = $match.Groups['date'].Value
                $kb = $match.Groups['kb'].Value.Trim()
                $buildsText = $match.Groups['builds'].Value
                
                # Set type to any text after the closing parenthesis, or "PatchTuesday" if empty
                $type = if ($match.Groups['patchtype'].Success -and -not [string]::IsNullOrWhiteSpace($match.Groups['patchtype'].Value)) { 
                    $match.Groups['patchtype'].Value.Trim() 
                } else { 
                    "Patch Tuesday"
                    }
                
                # Process build numbers
                # Replace "and" with commas to simplify splitting
                $buildsText = $buildsText -replace ' and ', ', '
                $buildsText = $buildsText -replace ' or ', ', '
                
                # Split by comma and process each build
                foreach ($buildText in ($buildsText -split ',')) {
                    # Extract just the digits, removing any letters
                    $buildNumber = $buildText -replace '[^0-9.]', ''
                    
                    if ($buildNumber -match '(\d+)\.(\d+)') {
                        try {
                            $buildVersion = [version]$buildNumber
                            
                            # Determine OS type based on build number
                            $ClientOS = $null
                            $ServerOS = $null
                            
                            switch ($buildVersion.Major) { 
                                10240 { $ClientOS = "Windows 10 1507" }
                                10586 { $ClientOS = "Windows 10 1511" }
                                14393 { 
                                    $ClientOS = "Windows 10 1607"
                                    $ServerOS = "Windows Server 2016" 
                                }
                                15063 { $ClientOS = "Windows 10 1703" }
                                15254 { $ClientOS = "Windows 10 Mobile" }
                                16299 { $ClientOS = "Windows 10 1709" }
                                17134 { $ClientOS = "Windows 10 1803" }
                                17763 { 
                                    $ServerOS = "Windows Server 2019"
                                    $ClientOS = "Windows 10 1809" 
                                } 
                                18362 { $ClientOS = "Windows 10 1903" }
                                18363 { $ClientOS = "Windows 10 1909" }
                                19041 { $ClientOS = "Windows 10 2004" }
                                19042 { $ClientOS = "Windows 10 20H2" }
                                19043 { $ClientOS = "Windows 10 21H1" }
                                19044 { $ClientOS = "Windows 10 21H2" }
                                19045 { $ClientOS = "Windows 10 22H2" }
                                20348 { $ServerOS = "Windows Server 2022" } 
                                22000 { $ClientOS = "Windows 11 21H2" } 
                                22621 { $ClientOS = "Windows 11 22H2" }
                                22631 { $ClientOS = "Windows 11 23H2" }
                                26100 { $ClientOS = "Windows 11 24H2" }
                                25398 { $ServerOS = "Windows Server 23H2" }
                            }
                            
                            # Create a separate hashtable for each build number
                            $updateInfo = [ordered]@{
                                KB = $kb
                                Build = $buildVersion
                                Date = Get-Date $date -ErrorAction SilentlyContinue
                                ClientOS = $ClientOS
                                ServerOS = $ServerOS
                                PatchType = $type
                            }
                            
                            # Add this individual build entry to the results
                            $result += New-Object PSObject -Property $updateInfo
                        }
                        catch {
                            # Skip this build if we can't parse it as a version
                            continue
                        }
                    }
                }
            }
        }
    }
    
    End {
        $ProgressPreference = $OldProgressPreference
        # Return unique results sorted by date, KB, and build
        return $result | Select-Object * | Sort-Object Date, KB, Build | Get-Unique -AsString
    }
}