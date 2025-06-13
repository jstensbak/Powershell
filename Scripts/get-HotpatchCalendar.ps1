$Year = (Get-Date).Year

$url = "https://support.microsoft.com/en-us/topic/release-notes-for-hotpatch-on-windows-11-version-24h2-enterprise-clients-c0906ee6-5e62-498f-bd5a-8f4966349f3c"

$htmlContent = Invoke-WebRequest -Uri $url -UseBasicParsing
$html = $htmlContent.Content

# Use regex to find the section id for the specified year
$pattern = "<h2 id=""(ID[0-9A-Z]+)""[^>]*>Hotpatch calendar $Year</h2>"
if ($html -match $pattern) {
    $sectionId = $matches[1]
    Write-Host "Found section ID: $sectionId" -ForegroundColor Green
    $sectionPattern = "(?s)<section aria-labelledby=`"$sectionId`"[^>]*>(.*?)</section>"
    if ($html -match $sectionPattern) {
        $sectionContent = $matches[1]
        
        # Find all tables within the section
        $tablePattern = '(?s)<table[^>]*>.*?</table>'
        $tables = [regex]::Matches($sectionContent, $tablePattern)
        
        $hotpatchData = @()
        
        foreach ($table in $tables) {
            $tableContent = $table.Value
            
            # Extract all table rows
            $rowPattern = '(?s)<tr>(.*?)</tr>'
            $rows = [regex]::Matches($tableContent, $rowPattern)
            
            foreach ($row in $rows) {
                $rowContent = $row.Groups[1].Value
                
                # Skip header rows that contain <b> tags
                if ($rowContent -match '<b class="ocpLegacyBold">') {
                    continue
                }
                
                # Extract data from each cell using a more flexible pattern
                $cellPattern = '(?s)<td[^>]*>\s*<p[^>]*>(.*?)</p>\s*</td>'
                $cells = [regex]::Matches($rowContent, $cellPattern)
                
                if ($cells.Count -ge 3) {  # Changed from 4 to 3 to handle December entry
                    # Clean up the extracted text
                    $month = $cells[0].Groups[1].Value -replace '<[^>]*>', '' -replace '^\s+|\s+$', '' -replace '┬á', ' '
                    $updateType = $cells[1].Groups[1].Value -replace '<[^>]*>', '' -replace '^\s+|\s+$', '' -replace '┬á', ' '
                    $type = $cells[2].Groups[1].Value -replace '<[^>]*>', '' -replace '^\s+|\s+$', '' -replace '┬á', ' '
                      # Extract KB article and URL - handle both with and without links (check if 4th cell exists)
                    $kbArticle = ""
                    $kbUrl = ""
                    if ($cells.Count -ge 4) {
                        $kbCell = $cells[3].Groups[1].Value
                        if ($kbCell -match '<a href="([^"]*)"[^>]*>(KB\d+)</a>') {
                            $kbArticle = $matches[2]
                            $kbUrl = "https://support.microsoft.com" + $matches[1]
                        } elseif ($kbCell -match 'KB\d+') {
                            $kbArticle = $matches[0]
                            # No URL available for non-linked KB articles
                        }
                    }
                    
                    # Only add rows with actual data (skip empty months)
                    if ($month -and $month -ne "" -and $month -notmatch "Month") {                        $hotpatchData += [PSCustomObject]@{
                            Month = $month
                            UpdateType = $updateType
                            Type = $type
                            KBArticle = $kbArticle
                            KBUrl = $kbUrl
                        }
                    }
                }
            }
        }
        
        # Display results
        Write-Host "`nHotpatch Calendar for $Year" -ForegroundColor Yellow
        Write-Host "================================" -ForegroundColor Yellow
        $hotpatchData | Format-Table -AutoSize
        
        # Also return the data for further processing
        return $hotpatchData
    } else {
        Write-Error "Section for year $Year not found."
    }
} else {
    Write-Error "Hotpatch calendar for year $Year not found."
}