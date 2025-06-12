Function Get-HotpatchDataFromMicrosoft {
    <#
    .SYNOPSIS
        Downloads and processes hotpatch data from Microsoft's websites for Update History.
        
    .DESCRIPTION
        Uses class 'supLeftNavLink' to extract update information from Microsoft support pages.
        Parses update strings and creates individual objects for each Windows version/build.
        
        Examples of parsed strings:
        - "January 17, 2022—KB5010793 (OS Builds 19042.1469, 19043.1469, and 19044.1469) Out-of-band"
        - "January 11, 2022—KB5009557 (OS Build 17763.2452)"
        
    .PARAMETER IncludeWindowsMobile
        Include Windows 10 Mobile updates in the results. Default is false.
        
    .PARAMETER MaxRetries
        Maximum number of retries for web requests. Default is 3.
        
    .EXAMPLE
        Get-HotpatchDataFromMicrosoft
        
    .EXAMPLE
        Get-HotpatchDataFromMicrosoft -IncludeWindowsMobile -MaxRetries 5
        
    .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        Array of objects with properties: KB, Build, Date, ClientOS, ServerOS, PatchType
        
    .NOTES
        Version 1.1
        - Improved performance by using ArrayList instead of array concatenation
        - Enhanced error handling and retry logic
        - Added parameter validation
        - Externalized OS mapping for better maintainability
        - Added proper help documentation
    #>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject[]])]
    Param(
        [Parameter()]
        [switch]$IncludeWindowsMobile,
        
        [Parameter()]
        [ValidateRange(1, 10)]
        [int]$MaxRetries = 3
    )    Begin {
        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = "SilentlyContinue"
        
        # URI configuration
        $URIs = @(
            "https://support.microsoft.com/en-US/topic/release-notes-for-hotpatch-on-windows-11-version-24h2-enterprise-clients-c0906ee6-5e62-498f-bd5a-8f4966349f3c" # Windows 11 Hotpatch 24H2
        )
          # Regex pattern to match update information (compiled for better performance)
        # Updated to capture Out-of-band patches that appear after the build information
        $regex = [regex]::new('^(?<date>.+?)—(?<initialtype>Hotpatch|Baseline)?(?:\s+(?<kb>KB\d+))?\s+\(.*?(?<builds>[\d\.]+).*?\)(?:\s+(?<finaltype>Out-of-band|Preview))?', [System.Text.RegularExpressions.RegexOptions]::Compiled)
        
        # Use ArrayList for better performance than array concatenation
        $result = [System.Collections.ArrayList]::new()
        
        # OS mapping hashtable for better maintainability
        $OSMapping = @{
            10240 = @{ ClientOS = "Windows 10 1507"; ServerOS = $null }
            10586 = @{ ClientOS = "Windows 10 1511"; ServerOS = $null }
            14393 = @{ ClientOS = "Windows 10 1607"; ServerOS = "Windows Server 2016" }
            15063 = @{ ClientOS = "Windows 10 1703"; ServerOS = $null }
            15254 = @{ ClientOS = "Windows 10 Mobile"; ServerOS = $null }
            16299 = @{ ClientOS = "Windows 10 1709"; ServerOS = $null }
            17134 = @{ ClientOS = "Windows 10 1803"; ServerOS = $null }
            17763 = @{ ClientOS = "Windows 10 1809"; ServerOS = "Windows Server 2019" }
            18362 = @{ ClientOS = "Windows 10 1903"; ServerOS = $null }
            18363 = @{ ClientOS = "Windows 10 1909"; ServerOS = $null }
            19041 = @{ ClientOS = "Windows 10 2004"; ServerOS = $null }
            19042 = @{ ClientOS = "Windows 10 20H2"; ServerOS = $null }
            19043 = @{ ClientOS = "Windows 10 21H1"; ServerOS = $null }
            19044 = @{ ClientOS = "Windows 10 21H2"; ServerOS = $null }
            19045 = @{ ClientOS = "Windows 10 22H2"; ServerOS = $null }
            20348 = @{ ClientOS = $null; ServerOS = "Windows Server 2022" }
            22000 = @{ ClientOS = "Windows 11 21H2"; ServerOS = $null }
            22621 = @{ ClientOS = "Windows 11 22H2"; ServerOS = $null }
            22631 = @{ ClientOS = "Windows 11 23H2"; ServerOS = $null }
            26100 = @{ ClientOS = "Windows 11 24H2"; ServerOS = $null }
            25398 = @{ ClientOS = $null; ServerOS = "Windows Server 23H2" }
        }
        
        Write-Verbose "Initialized with $($URIs.Count) URI(s) to process"
    }
      Process {
        # Collect update information from all Microsoft pages
        $MSOutput = [System.Collections.ArrayList]::new()
        
        foreach ($URI in $URIs) {
            $retryCount = 0
            $success = $false
            
            while (-not $success -and $retryCount -lt $MaxRetries) {
                try {
                    Write-Verbose "Attempting to retrieve data from: $URI (Attempt $($retryCount + 1)/$MaxRetries)"
                    $webResult = Invoke-WebRequest -Uri $URI -UseBasicParsing -TimeoutSec 30
                    $success = $true
                    
                    foreach ($link in $webResult.Links) {
                        try {                            # Extract the text from supLeftNavLink class elements
                            if ($link.class -eq "supLeftNavLink") {
                                $htmlText = $link.OuterHTML
                                
                                # More robust text extraction from HTML
                                if ($htmlText -match '>([^<]+)</a>') {
                                    $decodedText = [System.Net.WebUtility]::HtmlDecode($matches[1])
                                }
                                else {
                                    # Fallback to original method if regex doesn't match
                                    $textParts = $htmlText.split(">") | Select-Object -Last 2 | Where-Object { $_.Trim() -ne "" }
                                    $cleanText = ($textParts -join "").TrimEnd("</a>")
                                    $decodedText = [System.Net.WebUtility]::HtmlDecode($cleanText)
                                }
                                
                                # Filter conditions
                                $containsEmDash = $decodedText -like "*—*"
                                $isWindowsMobile = $decodedText -like "*Windows 10 Mobile*"
                                
                                if ($containsEmDash -and ($IncludeWindowsMobile -or -not $isWindowsMobile)) {
                                    [void]$MSOutput.Add($decodedText)
                                }
                            }
                        }
                        catch {
                            Write-Debug "Error processing link: $_"
                            continue
                        }
                    }
                    
                    Write-Verbose "Successfully processed $($MSOutput.Count) entries from $URI"
                }
                catch {
                    $retryCount++
                    Write-Warning "Failed to retrieve update data from $URI (Attempt $retryCount/$MaxRetries): $_"
                    
                    if ($retryCount -lt $MaxRetries) {
                        Start-Sleep -Seconds ([math]::Pow(2, $retryCount)) # Exponential backoff
                    }
                }
            }
            
            if (-not $success) {
                Write-Error "Failed to retrieve data from $URI after $MaxRetries attempts"
            }
        }
          # Process each update string collected from the websites
        Write-Verbose "Processing $($MSOutput.Count) update entries"
        
        foreach ($line in $MSOutput) {
            $match = $regex.Match($line)
            if ($match.Success) {
                try {
                    $date = $match.Groups['date'].Value
                    $kb = $match.Groups['kb'].Value.Trim()
                    $buildsText = $match.Groups['builds'].Value
                      # Set type based on patch type groups, prioritizing finaltype (Out-of-band) over initialtype
                    $type = if ($match.Groups['finaltype'].Success -and -not [string]::IsNullOrWhiteSpace($match.Groups['finaltype'].Value)) {
                        # If there's a final type (like "Out-of-band"), use that
                        $match.Groups['finaltype'].Value.Trim()
                    }
                    elseif ($match.Groups['initialtype'].Success -and -not [string]::IsNullOrWhiteSpace($match.Groups['initialtype'].Value)) {
                        # Otherwise use initial type (like "Hotpatch" or "Baseline")
                        $match.Groups['initialtype'].Value.Trim()
                    }
                    else {
                        # Default fallback
                        "Patch Tuesday"
                    }
                    
                    # Parse date with better error handling
                    $parsedDate = try {
                        Get-Date $date -ErrorAction Stop
                    }
                    catch {
                        Write-Warning "Failed to parse date '$date' for KB $kb"
                        $null
                    }
                    
                    # Process build numbers
                    $buildsText = $buildsText -replace ' and ', ', ' -replace ' or ', ', '
                    
                    # Split by comma and process each build
                    foreach ($buildText in ($buildsText -split ',')) {
                        $buildNumber = $buildText -replace '[^0-9.]', ''
                        
                        if ($buildNumber -match '^(\d+)\.(\d+)(?:\.(\d+))?(?:\.(\d+))?$') {
                            try {
                                $buildVersion = [version]$buildNumber
                                
                                # Get OS information from mapping
                                $osInfo = $OSMapping[$buildVersion.Major]
                                if (-not $osInfo) {
                                    Write-Debug "Unknown build major version: $($buildVersion.Major)"
                                    $osInfo = @{ ClientOS = $null; ServerOS = $null }
                                }
                                
                                # Create update information object
                                $updateInfo = [PSCustomObject]@{
                                    PSTypeName = 'Microsoft.Hotpatch.UpdateInfo'
                                    KB = $kb
                                    Build = $buildVersion
                                    Date = $parsedDate
                                    ClientOS = $osInfo.ClientOS
                                    ServerOS = $osInfo.ServerOS
                                    PatchType = $type
                                }
                                
                                # Add to results
                                [void]$result.Add($updateInfo)
                            }
                            catch {
                                Write-Debug "Failed to parse build version '$buildNumber': $_"
                                continue
                            }
                        }
                        else {
                            Write-Debug "Build number '$buildNumber' doesn't match expected pattern"
                        }
                    }
                }
                catch {
                    Write-Warning "Error processing line: $line. Error: $_"
                    continue
                }
            }
            else {
                Write-Debug "Line doesn't match regex pattern: $line"
            }
        }
    }
      End {
        $ProgressPreference = $OldProgressPreference
        
        Write-Verbose "Collected $($result.Count) update entries"
        
        # Return unique results sorted by date, KB, and build
        # Convert ArrayList to array and remove duplicates more efficiently
        $uniqueResults = $result.ToArray() | 
            Sort-Object Date, KB, Build | 
            Group-Object KB, Build | 
            ForEach-Object { $_.Group[0] }
        
        Write-Verbose "Returning $($uniqueResults.Count) unique update entries"
        return $uniqueResults
    }
}