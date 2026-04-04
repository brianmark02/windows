<#
.SYNOPSIS
    Syncs files from SharePoint to AWS S3 (Verify & Copy).

.DESCRIPTION
    1. Connects to SharePoint (WebLogin) & AWS S3.
    2. Iterates through 'Documents' library.
    3. Checks if file exists in S3.
    4. If Missing or Size Mismatch -> Copies (Download SP -> Upload S3).
    5. If Match -> Skips.
    6. Generates a consolidated CSV report.

.PARAMETER ConfigFilePath
    Path to JSON config. Defaults to ".\SyncConfig.json"
#>

param(
    [string]$ConfigFilePath = ".\SyncConfig.json"
)

# -----------------------------------------------------------------------------
# 1. SETUP & CONFIGURATION
# -----------------------------------------------------------------------------

if (-not (Test-Path $ConfigFilePath)) {
    Write-Error "Configuration file not found at $ConfigFilePath."
    exit 1
}

try {
    $config = Get-Content -Path $ConfigFilePath -Raw | ConvertFrom-Json

    $S3BucketName = $config.S3BucketName.Trim()
    $TempDownloadPath = if ($config.TempDownloadPath) { $config.TempDownloadPath } else { $env:TEMP }
    $AwsAccessKey = $config.AwsAccessKey.Trim()
    $AwsSecretKey = $config.AwsSecretKey.Trim()
    $AwsRegion = if ($config.AwsRegion) { $config.AwsRegion.Trim() } else { "ap-southeast-1" }

    # Load Site List
    $siteUrls = @()
    if ($config.SiteList -and $config.SiteList.Count -gt 0) {
        $siteUrls = $config.SiteList
    }
    elseif ($config.SiteListPath -and (Test-Path $config.SiteListPath)) {
        Write-Host "Reading site list from $($config.SiteListPath)" -ForegroundColor Cyan
        $siteUrls = Get-Content $config.SiteListPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    }
    elseif ($config.SharePointSiteUrl) {
        $siteUrls = @($config.SharePointSiteUrl)
    }
    else {
        Write-Error "No sites found. Configure 'SiteListPath' or 'SharePointSiteUrl'."
        exit 1
    }

    if ([string]::IsNullOrWhiteSpace($S3BucketName)) {
        Write-Error "S3BucketName is required."
        exit 1
    }
}
catch {
    Write-Error "Config Load Failed: $_"
    exit 1
}

# Ensure Temp Path Exists
if (-not (Test-Path $TempDownloadPath)) {
    New-Item -ItemType Directory -Force -Path $TempDownloadPath | Out-Null
}

# Check Modules
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "PnP.PowerShell module is missing."
    exit 1
}

if (-not (Get-Module -ListAvailable -Name AWSPowerShell) -and -not (Get-Module -ListAvailable -Name AWS.Tools.S3)) {
    Write-Error "AWS PowerShell module is missing."
    exit 1
}

# -----------------------------------------------------------------------------
# 2. MAIN LOGIC
# -----------------------------------------------------------------------------

try {
    # Initialize AWS parameters for the main thread explicitly to avoid changing default config
    Write-Host "Initializing AWS..." -NoNewline
    $awsGlobalParams = @{
        BucketName  = $S3BucketName
        ErrorAction = 'Stop'
    }
    if (-not [string]::IsNullOrWhiteSpace($AwsAccessKey)) {
        $awsGlobalParams.AccessKey = $AwsAccessKey
        $awsGlobalParams.SecretKey = $AwsSecretKey
        $awsGlobalParams.Region    = $AwsRegion
    }

    # Verify S3 Access
    try {
        $verifyParams = $awsGlobalParams.Clone()
        $verifyParams.MaxKeys = 1
        Get-S3Object @verifyParams | Select-Object -First 1 | Out-Null
        Write-Host " OK" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to access S3 Bucket '$S3BucketName'. $_"
        exit 1
    }

    $ReportFilePath = ".\SyncReport.csv"
    $ProcessedLogPath = ".\processed_sites.log"

    # Load Processed Sites into a HashSet for fast, case-insensitive lookup
    $processedSites = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if (Test-Path $ProcessedLogPath) {
        Get-Content $ProcessedLogPath | ForEach-Object {
            if (-not [string]::IsNullOrWhiteSpace($_)) {
                $cleanUrl = $_.Trim().TrimEnd('/')
                [void]$processedSites.Add($cleanUrl)
            }
        }
        Write-Host "Loaded $($processedSites.Count) processed sites." -ForegroundColor Gray
    }

    # Global variables for caching token to prevent interactive login prompts on every site
    $cachedAccessToken = $null
    $tokenExpiry = (Get-Date).AddYears(-1)

    foreach ($url in $siteUrls) {
        $url = $url.Trim()
        $normalizedUrl = $url.TrimEnd('/')

        if ($processedSites.Contains($normalizedUrl)) {
            Write-Host "Skipping Site: $url (Already Completed)" -ForegroundColor DarkGray
            continue
        }

        Write-Host "`n--------------------------------------------------"
        Write-Host "Processing Site: $url" -ForegroundColor Cyan
        Write-Host "--------------------------------------------------"

        try {
            # Authentication Setup
            $clientId = $config.ClientId

            # Connect SharePoint (Hybrid Support) - Restored for PS 5.1 Compatibility
            Write-Host "Connecting to SharePoint..." -NoNewline

            $siteConnection = $null
            if ($PSVersionTable.PSVersion.Major -lt 7) {
                # PowerShell 5.1 (Legacy) - Use WebLogin (Bypasses Modern App Blocks)
                try {
                    $siteConnection = Connect-PnPOnline -Url $url -UseWebLogin -ReturnConnection -ErrorAction Stop
                    Write-Host " Connected (WebLogin - Legacy)" -ForegroundColor Green
                }
                catch {
                    Write-Error "`nPS5.1 WebLogin Failed: $($_.Exception.Message)"
                    throw $_
                }
            }
            else {
                # PowerShell 7+ (Modern)
                # Default to "SharePoint Online Management Shell" App ID (Public Client)
                if ([string]::IsNullOrWhiteSpace($clientId) -or -not ($clientId -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')) {
                    $clientId = "9bc3ab49-b65d-410a-85ad-de819febfddc"
                }

                try {
                    # If we don't have a token or it's expired, do an interactive login
                    if ([string]::IsNullOrWhiteSpace($cachedAccessToken) -or (Get-Date) -gt $tokenExpiry) {
                        Write-Host "`n[AUTH] Starting Interactive Login (Token will be cached)..." -ForegroundColor Yellow

                        $siteConnection = Connect-PnPOnline -Url $url -Interactive -ClientId $clientId -ReturnConnection -ErrorAction Stop
                        Write-Host " Connected (Interactive)" -ForegroundColor Green

                        # Extract the token so we can reuse it for the next 45 minutes
                        try {
                            $cachedAccessToken = Get-PnPAccessToken -Connection $siteConnection -ErrorAction Stop
                            $tokenExpiry = (Get-Date).AddMinutes(45) # Typical Azure AD token lifetime is ~60m, refreshing at 45m is safe
                        } catch {
                            Write-Warning "Could not cache Access Token. You may be prompted to log in again for the next site."
                        }
                    } else {
                        Write-Host " Connected (Using Cached Token)" -ForegroundColor Green
                        $secureToken = ConvertTo-SecureString $cachedAccessToken -AsPlainText -Force
                        $siteConnection = Connect-PnPOnline -Url $url -AccessToken $secureToken -ReturnConnection -ErrorAction Stop
                    }
                }
                catch {
                    Write-Error "`nFAILED: $($_.Exception.Message)"
                    Write-Warning "Authentication Failed."
                    throw $_
                }
            }

            $web = Get-PnPWeb -Connection $siteConnection
            $siteName = $web.Title
            # Sanitize: Remove invalid chars AND trim dots/spaces from start/end
            $siteFolderName = ($siteName -replace '[\\/:*?"<>|]', '').Trim(' .')
            Write-Host "Site Name: $siteName" -ForegroundColor Green
            Write-Host "S3 Folder: $siteFolderName" -ForegroundColor DarkGray

            # Get ALL Document and Page Libraries
            Write-Host "Gathering all valid Document and Page Libraries..." -ForegroundColor Yellow
            $availableLibs = Get-PnPList -Connection $siteConnection | Where-Object {
                $_.BaseType -eq 1 -and
                $_.Hidden -eq $false -and
                $_.Title -notmatch '(?i)^(AppPackages|Apps for SharePoint|Client Side Assets|MicroFeed|ContentTypeSyncLog)$'
            }

            if (-not $availableLibs) {
                Write-Warning "No suitable document libraries found. Skipping."
                continue
            }

            Write-Host "Libraries to Sync:" -ForegroundColor Cyan
            $availableLibs | Select-Object Title, ItemCount, BaseTemplate | Format-Table -AutoSize | Out-String | Write-Host

            # -----------------------------------------------------------------
            # PHASE 1: Bulk S3 Listing (Cache Optimization)
            # -----------------------------------------------------------------
            $s3Cache = @{}
            try {
                Write-Host "Fetching S3 Object List for Cache..." -NoNewline

                # Fetch ALL objects for this site-folder prefix to minimize API calls
                $cacheParams = $awsGlobalParams.Clone()
                $cacheParams.KeyPrefix = "$siteFolderName/"
                $s3Objects = Get-S3Object @cacheParams

                foreach ($obj in $s3Objects) {
                    $s3Cache[$obj.Key] = $obj
                }
                Write-Host " Done. Cached $($s3Cache.Count) files." -ForegroundColor Green
            }
            catch {
                Write-Warning "Failed to list S3 objects. Proceeding without cache (Slow Mode). $_"
            }

            $counters = @{ Unchanged = 0; Copied = 0; Updated = 0; Errors = 0 }

            foreach ($lib in $availableLibs) {
                if ($lib.ItemCount -eq 0) { continue }

                Write-Host "`n--- Scanning Library '$($lib.Title)' ---" -ForegroundColor Yellow

                # CAML Query: RecursiveAll (Traverse Folders)
                $camlQuery = "<View Scope='RecursiveAll'><RowLimit Paged='TRUE'>1000</RowLimit></View>"
                $items = Get-PnPListItem -List $lib -PageSize 1000 -Query $camlQuery -Connection $siteConnection -ErrorAction Stop

                if (-not $items -or $items.Count -eq 0) {
                    Write-Host "No items found in $($lib.Title)." -ForegroundColor DarkGray
                    continue
                }

                # -----------------------------------------------------------------
                # PHASE 2: Sequential Analysis (Scanning & Comparison)
                # -----------------------------------------------------------------
                # Parallel analysis here was an anti-pattern as it serializes the large $s3Cache across runspaces
                # We now do sequential scanning, which is very fast in-memory
                $webServerRelUrl = $web.ServerRelativeUrl
                $syncPlans = @()

            if ($PSVersionTable.PSVersion.Major -ge 7) {
                Write-Host "Analyzing files (Sequential - In Memory)..." -ForegroundColor Yellow

                $syncPlans = $items | ForEach-Object {
                    $item = $_

                    if ($item.FileSystemObjectType -eq 1) {
                        # Folder Logic
                        $fRef = $item.FieldValues.FileRef
                        $folderRelPath = $fRef.Substring($webServerRelUrl.Length)
                        if ($folderRelPath.StartsWith("/")) { $folderRelPath = $folderRelPath.Substring(1) }
                        $safeKeyPath = $folderRelPath -replace '\\', '/'
                        $s3Key = "$siteFolderName/$safeKeyPath/" # Ensure trailing slash

                        if ($s3Cache.ContainsKey($s3Key)) {
                            return [PSCustomObject]@{
                                Action          = "Skip"
                                Status          = "Match"
                                SiteUrl         = $url
                                Site            = $siteName
                                SP_Path         = $fRef
                                SP_Size         = 0
                                SP_DateModified = "N/A"
                                S3_Path         = $s3Key
                                S3_Size         = 0
                                S3_DateCopied   = "N/A"
                                Notes           = "Folder matches"
                                FileName        = $item.FieldValues.FileLeafRef
                            }
                        }
                        else {
                            return [PSCustomObject]@{
                                Action          = "Create Folder"
                                Status          = "Missing"
                                SiteUrl         = $url
                                Site            = $siteName
                                SP_Path         = $fRef
                                SP_Size         = 0
                                SP_DateModified = "N/A"
                                S3_Path         = $s3Key
                                S3_Size         = 0
                                S3_DateCopied   = "N/A"
                                Notes           = "Missing Folder"
                                FileName        = $item.FieldValues.FileLeafRef
                            }
                        }
                    }

                    # File
                    $file = $item.FieldValues
                    $fileName = $file.FileLeafRef
                    $spFileSize = $file.File_x0020_Size
                    $serverRelativeUrl = $file.FileRef
                    $spLastMod = $file.Modified

                    # Handle Shortcuts (.lnk)
                    if ($fileName -like "*.lnk") {
                        return [PSCustomObject]@{
                            Action          = "Skip"
                            SiteUrl         = $url
                            Site            = $siteName
                            SP_Path         = $serverRelativeUrl
                            SP_Size         = [math]::Round($spFileSize / 1MB, 2)
                            SP_DateModified = if ($spLastMod) { $spLastMod.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
                            S3_Path         = "N/A"
                            S3_Size         = 0
                            S3_DateCopied   = "N/A"
                            Status          = "Skipped"
                            Notes           = "Shortcut file (.lnk) skipped"
                        }
                    }

                    # Calculate S3 Key
                    $siteRelativePath = $serverRelativeUrl
                    if ($webServerRelUrl -ne "/") {
                        $siteRelativePath = $serverRelativeUrl.Substring($webServerRelUrl.Length)
                    }
                    if ($siteRelativePath.StartsWith("/")) { $siteRelativePath = $siteRelativePath.Substring(1) }

                    # Sanitize S3 Key to match S3 standards but Allow Slashes
                    # 1. Replace Backslashes with Forward Slashes
                    $safeKeyPath = $siteRelativePath -replace '\\', '/'

                    $s3Key = "$siteFolderName/$safeKeyPath"

                    $status = "Pending"
                    $action = ""
                    $notes = ""
                    $needsDownload = $false
                    $s3Size = 0
                    $s3Date = "N/A"

                    # Check S3 (Using Cache)
                    if ($s3Cache.ContainsKey($s3Key)) {
                        $s3Obj = $s3Cache[$s3Key]
                        $s3Size = $s3Obj.Size
                        $s3Date = $s3Obj.LastModified

                        # Compare Size (Force Long)
                        $sizeMatch = ([long]$s3Size -eq [long]$spFileSize)

                        # Compare Date (DISABLED - S3 Date is Upload Time)
                        # Strategy: Use Size Only to prevent infinite loops.

                        if ($sizeMatch) {
                            $status = "Match"
                            $notes = "File matches S3 (Size)"
                        }
                        else {
                            $status = "Mismatch"
                            $action = "Update"
                            $notes = "Diff: Size(SP:$spFileSize vs S3:$s3Size)"
                            $needsDownload = $true
                        }
                    }
                    else {
                        $status = "Missing"
                        $action = "Copy"
                        $notes = "File missing in S3"
                        $needsDownload = $true
                    }

                    return [PSCustomObject]@{
                        Action            = if ($needsDownload) { $action } else { "Skip" }
                        SiteUrl           = $url
                        Site              = $siteName
                        SP_RawSize        = $spFileSize
                        SP_Path           = $serverRelativeUrl
                        SP_Size           = [math]::Round($spFileSize / 1MB, 2)
                        SP_DateModified   = if ($spLastMod) { $spLastMod.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
                        S3_Path           = $s3Key
                        S3_Size           = [math]::Round($s3Size / 1MB, 2)
                        S3_DateCopied     = if ($s3Date -is [datetime]) { $s3Date.ToString("yyyy-MM-dd HH:mm:ss") } else { $s3Date }
                        Status            = $status
                        Notes             = $notes
                        FileName          = $fileName
                        ServerRelativeUrl = $serverRelativeUrl
                        ItemId            = $item.Id
                    }
                }
            }
            else {
                # -------------------------
                # PS 5.1 SEQUENTIAL (Real-time Reporting + Fail-safe S3 Lookup)
                # -------------------------
                Write-Host "Analyzing & Syncing (Sequential - PS5.1 - RealTime)..." -ForegroundColor Yellow

                $items | ForEach-Object {
                    $item = $_

                    if ($item.FileSystemObjectType -eq 1) {
                        # Folder Logic
                        $fRef = $item.FieldValues.FileRef
                        $folderRelPath = $fRef.Substring($webServerRelUrl.Length)
                        if ($folderRelPath.StartsWith("/")) { $folderRelPath = $folderRelPath.Substring(1) }
                        $safeKeyPath = $folderRelPath -replace '\\', '/'
                        $s3Key = "$siteFolderName/$safeKeyPath/"

                        # Check if Exists
                        if (-not $s3Cache.ContainsKey($s3Key)) {
                            try {
                                $directParams = $awsGlobalParams.Clone()
                                $directParams.KeyPrefix = $s3Key
                                $directParams.ErrorAction = 'SilentlyContinue'
                                $directObj = Get-S3Object @directParams | Where-Object { $_.Key -eq $s3Key } | Select-Object -First 1
                                if ($directObj) { $s3Cache[$s3Key] = $directObj }
                            }
                            catch {}
                        }

                        if (-not $s3Cache.ContainsKey($s3Key)) {
                            # Create Folder
                            Write-Host "Creating Folder: $s3Key" -ForegroundColor Yellow
                            try {
                                $emptyTmp = [System.IO.Path]::GetTempFileName()
                                $folderAwsParams = $awsGlobalParams.Clone()
                                $folderAwsParams.Key = $s3Key
                                $folderAwsParams.File = $emptyTmp
                                Write-S3Object @folderAwsParams
                                Remove-Item -LiteralPath $emptyTmp -Force
                                Write-Host " -> Created" -ForegroundColor Green
                                $counters.Copied++

                                [PSCustomObject]@{
                                    SiteUrl         = $url
                                    Site            = $siteName
                                    SP_Path         = $fRef
                                    SP_Size         = 0
                                    SP_DateModified = "N/A"
                                    S3_Path         = $s3Key
                                    S3_Size         = 0
                                    S3_DateCopied   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                                    Status          = "Created Folder"
                                    Notes           = "New Folder"
                                } | Export-Csv -Path $ReportFilePath -NoTypeInformation -Append -Force
                            }
                            catch {
                                Write-Warning "Folder Create Failed: $_"
                                $counters.Errors++
                            }
                        }
                        return
                    }

                    # File
                    $file = $item.FieldValues
                    $fileName = $file.FileLeafRef
                    $spFileSize = $file.File_x0020_Size
                    $serverRelativeUrl = $file.FileRef
                    $spLastMod = $file.Modified

                    # [v5 Feature] Verbose Logging
                    Write-Host "Processing File: $fileName" -ForegroundColor DarkGray

                    # 1. Handle Shortcuts (.lnk) - Export Immediately
                    if ($fileName -like "*.lnk") {
                        Write-Warning "Skipping Shortcut File: $fileName"
                        [PSCustomObject]@{
                            SiteUrl         = $url
                            Site            = $siteName
                            SP_Path         = $serverRelativeUrl
                            SP_Size         = [math]::Round($spFileSize / 1MB, 2)
                            SP_DateModified = if ($spLastMod) { $spLastMod.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
                            S3_Path         = "N/A"
                            S3_Size         = 0
                            S3_DateCopied   = "N/A"
                            Status          = "Skipped"
                            Notes           = "Shortcut file (.lnk) skipped"
                        } | Export-Csv -Path $ReportFilePath -NoTypeInformation -Append -Force
                        return # Continue loop
                    }

                    # 2. Calculate S3 Key
                    $siteRelativePath = $serverRelativeUrl
                    if ($webServerRelUrl -ne "/") {
                        $siteRelativePath = $serverRelativeUrl.Substring($webServerRelUrl.Length)
                    }
                    if ($siteRelativePath.StartsWith("/")) { $siteRelativePath = $siteRelativePath.Substring(1) }

                    $safeKeyPath = $siteRelativePath -replace '\\', '/'
                    $s3Key = "$siteFolderName/$safeKeyPath"

                    $status = "Pending"
                    $action = ""
                    $notes = ""
                    $needsDownload = $false
                    $s3Size = 0
                    $s3Date = "N/A"

                    # 3. Check S3 (Cache + Fallback)
                    $s3Obj = $null
                    # A. Try Cache
                    if ($s3Cache.ContainsKey($s3Key)) {
                        $s3Obj = $s3Cache[$s3Key]
                    }
                    # B. Fallback (The Missing Function)
                    else {
                        try {
                            # Verify if truly missing
                            $directParams = $awsGlobalParams.Clone()
                            $directParams.KeyPrefix = $s3Key
                            $directParams.ErrorAction = 'SilentlyContinue'
                            $directObj = Get-S3Object @directParams | Where-Object { $_.Key -eq $s3Key } | Select-Object -First 1
                            if ($directObj) {
                                $s3Obj = $directObj
                                # Optionally update cache? $s3Cache[$s3Key] = $directObj
                            }
                        }
                        catch { }
                    }

                    if ($s3Obj) {
                        $s3Size = $s3Obj.Size
                        $s3Date = $s3Obj.LastModified

                        # Compare Size (Force Long)
                        $sizeMatch = ([long]$s3Size -eq [long]$spFileSize)

                        # Size Only Logic
                        if ($sizeMatch) {
                            $status = "Match"
                            $notes = "File matches S3 (Size)"
                            $counters.Unchanged++
                        }
                        else {
                            $status = "Mismatch"
                            $action = "Update"
                            $notes = "Diff: Size(SP:$spFileSize vs S3:$s3Size)"
                            Write-Host "  [Diff] $fileName -> $notes" -ForegroundColor Yellow
                            $needsDownload = $true
                        }
                    }
                    else {
                        $status = "Missing"
                        $action = "Copy"
                        $notes = "File missing in S3"
                        $needsDownload = $true
                    }

                    $finalStatus = $status
                    $finalNotes = $notes
                    $s3SizeReport = if ($s3Size) { [math]::Round($s3Size / 1MB, 2) } else { 0 }
                    $s3DateReport = if ($s3Date -is [datetime]) { $s3Date.ToString("yyyy-MM-dd HH:mm:ss") } else { $s3Date }

                    # 4. ACTION: Download & Upload (Merged for Real-Time)
                    if ($needsDownload) {
                        $safeTempName = [System.Guid]::NewGuid().ToString() + ".tmp"
                        $localTempFile = "$TempDownloadPath\$safeTempName"
                        $maxRetries = 3
                        $retryCount = 0
                        $downloaded = $false

                        try {
                            do {
                                try {
                                    try {
                                        Write-Progress -Activity "Downloading from SharePoint" -Status "File: $fileName ($([math]::Round($spFileSize/1MB, 2)) MB)" -PercentComplete -1
                                        Write-Host "Downloading ($action) [Attempt $($retryCount + 1)]: $fileName" -NoNewline

                                        # Attempt 1: Standard PnP (REST)
                                        if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($fileName)) { throw "Filename contains PowerShell wildcard characters which breaks Get-PnPFile parameter binding (fallback to CSOM)" }
                                        Get-PnPFile -Url $serverRelativeUrl -Path $TempDownloadPath -FileName $safeTempName -AsFile -Force -Connection $siteConnection -ErrorAction Stop

                                        Write-Progress -Activity "Downloading from SharePoint" -Completed
                                        $downloaded = $true
                                        Write-Host " -> Downloaded" -ForegroundColor Green
                                    }
                                    catch {
                                        Write-Warning "`n -> Standard download failed ($($_.Exception.Message)). Switching to CSOM..."
                                        try {
                                            # Attempt 2: CSOM Fallback (ID-Based for robustness)
                                        $ctx = Get-PnPContext -Connection $siteConnection
                                            $list = $ctx.Web.Lists.GetByTitle($lib.Title)
                                            $spItem = $list.GetItemById($item.Id)
                                            $fileObj = $spItem.File
                                            $stream = $fileObj.OpenBinaryStream()
                                            $ctx.ExecuteQuery()

                                            if ($stream.Value) {
                                                try {
                                                    $fs = [System.IO.File]::Create($localTempFile)
                                                    $stream.Value.CopyTo($fs)
                                                    $downloaded = $true
                                                    Write-Host " -> Downloaded (CSOM)" -ForegroundColor Green
                                                }
                                                finally {
                                                    if ($null -ne $fs) { $fs.Dispose() }
                                                    if ($null -ne $stream.Value) { $stream.Value.Dispose() }
                                                }
                                            }
                                            else {
                                                throw "CSOM Stream was null"
                                            }
                                        }
                                        catch {
                                            throw $_
                                        }
                                    }
                                }
                                catch {
                                    $retryCount++
                                    $errMsg = $_.Exception.Message

                                    # Graceful handle for SharePoint Virus Scanner Blocks
                                    if ($errMsg -match "virus scanner discovered an issue" -or $errMsg -match "VBS/Ramnit") {
                                        Write-Warning "`n -> File Blocked by SharePoint Virus Scanner: $errMsg"
                                        $finalStatus = "Blocked (Virus)"
                                        $finalNotes = "Err: Virus Scanner Block"
                                        $counters.Errors++
                                        $retryCount = $maxRetries # Break loop immediately
                                    }
                                    elseif ($retryCount -lt $maxRetries) {
                                        Write-Warning "`n -> Download Failed: $errMsg. Retrying..."
                                        Start-Sleep -Seconds 5
                                    }
                                    else {
                                        Write-Warning "`n -> Download Failed Final: $errMsg"
                                        $finalStatus = "Download Failed"
                                        $finalNotes = "Err: $errMsg"
                                        $counters.Errors++
                                    }
                                }
                            } while (-not $downloaded -and $retryCount -lt $maxRetries)

                            if ($downloaded) {
                                try {
                                    $uploadRetryCount = 0
                                    $uploaded = $false
                                    do {
                                        try {
                                            Write-Host "Uploading to S3..." -NoNewline
                                            $fileAwsParams = $awsGlobalParams.Clone()
                                            $fileAwsParams.Key = $s3Key
                                            $fileAwsParams.File = $localTempFile
                                            Write-S3Object @fileAwsParams
                                            Write-Host " -> Done" -ForegroundColor Green
                                            $uploaded = $true
                                            $finalStatus = "Synced ($action)"
                                            $counters.Copied++

                                            $s3SizeReport = [math]::Round((Get-Item -LiteralPath $localTempFile).Length / 1MB, 2)
                                            $s3DateReport = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                                        }
                                        catch {
                                            $uploadRetryCount++
                                            if ($uploadRetryCount -lt $maxRetries) {
                                                Start-Sleep -Seconds 5
                                            }
                                            else {
                                                $finalStatus = "Upload Failed"
                                                $finalNotes = "UpErr: $($_.Exception.Message)"
                                                $counters.Errors++
                                            }
                                        }
                                    } while (-not $uploaded -and $uploadRetryCount -lt $maxRetries)
                                }
                                catch { throw $_ }
                            }
                        }
                        finally {
                            if (Test-Path -LiteralPath $localTempFile) { Remove-Item -LiteralPath $localTempFile -Force }
                        }
                    }

                    # 5. REPORT: Export Immediately (Real-Time)
                    [PSCustomObject]@{
                        SiteUrl         = $url
                        Site            = $siteName
                        SP_Path         = $serverRelativeUrl
                        SP_Size         = [math]::Round($spFileSize / 1MB, 2)
                        SP_DateModified = if ($spLastMod) { $spLastMod.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
                        S3_Path         = $s3Key
                        S3_Size         = $s3SizeReport
                        S3_DateCopied   = $s3DateReport
                        Status          = $finalStatus
                        Notes           = $finalNotes
                    } | Export-Csv -Path $ReportFilePath -NoTypeInformation -Append -Force
                }

                # IMPORTANT: Nullify $syncPlans to prevent Phase 3 batch processing logic from running again
                $syncPlans = $null
            }
            # -----------------------------------------------------------------
            # PHASE 3: Parallel Logic Completion (Only runs for PS 7)
            # -----------------------------------------------------------------
            if ($syncPlans) {
                # Filter items that need action
                $filesToSync = $syncPlans | Where-Object { $_.Action -eq "Copy" -or $_.Action -eq "Update" -or $_.Action -eq "Create Folder" }
                $unchangedCount = ($syncPlans | Where-Object { $_.Action -eq "Skip" -and $_.Status -eq "Match" }).Count
                $counters.Unchanged = $unchangedCount

                Write-Host "Analysis Complete. Unchanged: $unchangedCount. To Sync: $($filesToSync.Count)" -ForegroundColor Cyan

                # Capture variables for use in the parallel scope
                $currentLibTitle = $lib.Title

                # Extract access token securely from the main thread so we can inject it into runspaces
                $accessToken = $null
                try {
                    # Get-PnPAccessToken is the standard supported cmdlet in modern PnP versions
                    $accessToken = Get-PnPAccessToken -Connection $siteConnection -ErrorAction Stop
                } catch {
                    Write-Warning "Failed to retrieve Access Token. Parallel thread connections may fail."
                }

                $syncedResults = $filesToSync | ForEach-Object -Parallel {
                    $plan = $_
                    $fileName = $plan.FileName
                    $serverRelativeUrl = $plan.SP_Path
                    $s3Key = $plan.S3_Path
                    $action = $plan.Action
                    $spFileSize = $plan.SP_RawSize
                    $libTitle = $using:currentLibTitle
                    $S3BucketName = $using:S3BucketName
                    $TempDownloadPath = $using:TempDownloadPath
                    $url = $using:url
                    $AwsAccessKey = $using:AwsAccessKey
                    $AwsSecretKey = $using:AwsSecretKey
                    $AwsRegion = $using:AwsRegion
                    $accessToken = $using:accessToken

                    # Force module load inside the runspace
                    Import-Module PnP.PowerShell -ErrorAction SilentlyContinue

                    # Setup AWS parameters explicitly (bypasses runspace default config issues)
                    $awsParams = @{
                        BucketName  = $S3BucketName
                        ErrorAction = 'Stop'
                    }
                    if (-not [string]::IsNullOrWhiteSpace($AwsAccessKey)) {
                        $awsParams.AccessKey = $AwsAccessKey
                        $awsParams.SecretKey = $AwsSecretKey
                        $awsParams.Region    = $AwsRegion
                    }

                    # Reconnect locally within the thread using the captured Bearer token
                    # This ensures 100% thread safety without needing the missing Clone-PnPConnection cmdlet
                    $threadConnection = $null
                    if (-not [string]::IsNullOrWhiteSpace($accessToken)) {
                        try {
                            $secureToken = ConvertTo-SecureString $accessToken -AsPlainText -Force
                            $threadConnection = Connect-PnPOnline -Url $url -ReturnConnection -ErrorAction Stop -WarningAction SilentlyContinue -AccessToken $secureToken
                        } catch {
                            Write-Warning "Failed to setup thread connection: $($_.Exception.Message)"
                        }
                    } else {
                        Write-Warning "No Access Token found. Downloading in parallel may fail with Unauthorized."
                    }

                    # --- FOLDER LOGIC ---
                    if ($action -eq "Create Folder") {
                        try {
                            Write-Host "Creating Folder: $s3Key" -NoNewline
                            $emptyTmp = [System.IO.Path]::GetTempFileName()

                            $folderAwsParams = $awsParams.Clone()
                            $folderAwsParams.Key = $s3Key
                            $folderAwsParams.File = $emptyTmp
                            Write-S3Object @folderAwsParams
                            Remove-Item -LiteralPath $emptyTmp -Force
                            Write-Host " -> Created" -ForegroundColor Green
                            $plan.Status = "Created"
                            $plan.S3_DateCopied = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                            return @{ Type = 'Copied'; Plan = $plan }
                        }
                        catch {
                            Write-Warning "Folder Create Failed: $_"
                            $plan.Status = "Folder Error"
                            $plan.Notes = "Err: $_"
                            return @{ Type = 'Error'; Plan = $plan }
                        }
                    }

                    # --- SYNC LOGIC (Files) ---
                    $safeTempName = [System.Guid]::NewGuid().ToString() + ".tmp"
                    $localTempFile = "$TempDownloadPath\$safeTempName"
                    $maxRetries = 3
                    $retryCount = 0
                    $downloaded = $false
                    $finalStatus = $plan.Status
                    $finalNotes = $plan.Notes
                    $returnType = 'Error'

                    try {
                        do {
                            try {
                                try {
                                    # Download
                                    Write-Host "Downloading ($action) [Attempt $($retryCount + 1)]: $fileName" -NoNewline

                                    if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($fileName)) { throw "Filename contains PowerShell wildcard characters which breaks Get-PnPFile parameter binding (fallback to CSOM)" }
                                    Get-PnPFile -Url $serverRelativeUrl -Path $TempDownloadPath -FileName $safeTempName -AsFile -Force -ErrorAction Stop -Connection $threadConnection

                                    $downloaded = $true
                                    Write-Host " -> Downloaded" -ForegroundColor Green
                                }
                                catch {
                                    Write-Warning "`n -> Standard download failed ($($_.Exception.Message)). Switching to CSOM..."
                                    try {
                                        # Attempt 2: CSOM Fallback (ID-Based for robustness)
                                        $ctx = Get-PnPContext -Connection $threadConnection
                                        $list = $ctx.Web.Lists.GetByTitle($libTitle)
                                        $spItem = $list.GetItemById($plan.ItemId)
                                        $fileObj = $spItem.File
                                        $stream = $fileObj.OpenBinaryStream()
                                        $ctx.ExecuteQuery()

                                        if ($stream.Value) {
                                            try {
                                                $fs = [System.IO.File]::Create($localTempFile)
                                                $stream.Value.CopyTo($fs)
                                                $downloaded = $true
                                                Write-Host " -> Downloaded (CSOM)" -ForegroundColor Green
                                            }
                                            finally {
                                                if ($null -ne $fs) { $fs.Dispose() }
                                                if ($null -ne $stream.Value) { $stream.Value.Dispose() }
                                            }
                                        }
                                        else {
                                            throw "CSOM Stream was null"
                                        }
                                    }
                                    catch {
                                        throw $_
                                    }
                                }
                            }
                            catch {
                                $retryCount++
                                $errMsg = $_.Exception.Message

                                # Graceful handle for SharePoint Virus Scanner Blocks
                                if ($errMsg -match "virus scanner discovered an issue" -or $errMsg -match "VBS/Ramnit") {
                                    Write-Warning "`n -> File Blocked by SharePoint Virus Scanner: $errMsg"
                                    $finalStatus = "Blocked (Virus)"
                                    $finalNotes = "Err: Virus Scanner Block"
                                    $retryCount = $maxRetries # Break loop immediately
                                }
                                elseif ($retryCount -lt $maxRetries) {
                                    Write-Warning "`n -> Download failed: $errMsg. Retrying in 5 seconds..."
                                    Start-Sleep -Seconds 5
                                }
                                else {
                                    Write-Warning "`n -> Failed to download after $maxRetries attempts: $errMsg"
                                    $finalStatus = "Download Failed"
                                    $finalNotes = "Failed after $maxRetries attempts: $errMsg"
                                }
                            }
                        } while (-not $downloaded -and $retryCount -lt $maxRetries)

                        if ($downloaded) {
                            try {
                                $uploadRetryCount = 0
                                $uploaded = $false
                                do {
                                    try {
                                        Write-Host "Uploading to S3 (Attempt $($uploadRetryCount + 1))..." -NoNewline
                                        $fileAwsParams = $awsParams.Clone()
                                        $fileAwsParams.Key = $s3Key
                                        $fileAwsParams.File = $localTempFile
                                        Write-S3Object @fileAwsParams
                                        Write-Host " -> Done" -ForegroundColor Green
                                        $uploaded = $true
                                        $finalStatus = "Synced ($action)"
                                        $returnType = 'Copied'
                                        # Update S3 info in plan for report
                                        $plan.S3_Size = [math]::Round((Get-Item -LiteralPath $localTempFile).Length / 1MB, 2)
                                        $plan.S3_DateCopied = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                                    }
                                    catch {
                                        $uploadRetryCount++
                                        $upErrMsg = $_.Exception.Message
                                        if ($uploadRetryCount -lt $maxRetries) {
                                            Write-Warning "`n -> Upload Failed: $upErrMsg. Retrying in 5 seconds..."
                                            Start-Sleep -Seconds 5
                                        }
                                        else {
                                            Write-Warning "`n -> Upload Failed after $maxRetries attempts: $upErrMsg"
                                            $finalStatus = "Upload Failed"
                                            $finalNotes = "Download OK, Upload Failed: $upErrMsg"
                                        }
                                    }
                                } while (-not $uploaded -and $uploadRetryCount -lt $maxRetries)
                            }
                            catch { throw $_ }
                        }
                    }
                    finally {
                        if (Test-Path -LiteralPath $localTempFile) { Remove-Item -LiteralPath $localTempFile -Force }
                    }

                    # Update Plan with Final Status
                    $plan.Status = $finalStatus
                    $plan.Notes = $finalNotes

                    return @{ Type = $returnType; Plan = $plan }
                } -ThrottleLimit 5

                # Tally results from parallel run and update the main syncPlans list
                $finalPlans = @()
                $finalPlans += ($syncPlans | Where-Object { $_.Action -eq "Skip" })

                foreach ($res in @($syncedResults)) {
                    if ($res.Type -eq 'Copied') { $counters.Copied++ }
                    elseif ($res.Type -eq 'Error') { $counters.Errors++ }

                    if ($null -ne $res.Plan) {
                        $finalPlans += $res.Plan
                    }
                }

                # -----------------------------------------------------------------
                # PHASE 2: Batch Logging (I/O Optimization)
                # -----------------------------------------------------------------
                # Export all plans (Unchanged + Synced + Errors) together
                $finalPlans | Select-Object SiteUrl, Site, SP_Path, SP_Size, SP_DateModified, S3_Path, S3_Size, S3_DateCopied, Status, Notes | Export-Csv -Path $ReportFilePath -NoTypeInformation -Append -Force

                # Cleanup Memory
                [System.GC]::Collect()
            }

            }

            Write-Host "Summary for $siteName : Unchanged: $($counters.Unchanged) | Copied/Updated: $($counters.Copied) | Errors: $($counters.Errors)" -ForegroundColor Cyan

            # Mark site as completed
            $url | Out-File -FilePath $ProcessedLogPath -Append -Encoding utf8 -Force
        }
        catch {
            $errMsg = $_.Exception.Message
            if ($errMsg -match "unauthorized operation" -or $errMsg -match "403" -or $errMsg -match "401") {
                Write-Host "Access Denied (Skipped): You do not have permissions to access $url" -ForegroundColor DarkGray
            } else {
                Write-Warning "Site Process Failed ($url): $_"
            }
        }
    }

    # EXPORT REPORT
    Write-Host "`nConsolidated Sync Report saved to: $ReportFilePath" -ForegroundColor Cyan

}
catch {
    Write-Error "Script Failed: $_"
}
