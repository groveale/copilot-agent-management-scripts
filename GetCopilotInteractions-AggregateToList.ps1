#############################################################
# Script to get CopilotInteractions from AuditLogs via Microsoft Graph,
# aggregate by agent into four granularities (daily, weekly, monthly, alltime),
# and upsert summary rows to separate SharePoint lists per granularity.
# Includes rolling cleanup: daily rows kept for 7 days, weekly rows for 13 weeks.
# Script is designed to run in Azure Automation.
# Supports managed identity (default), app registration + secret, or app registration + certificate.
#
# Contact alexgrover@microsoft.com for questions

#############################################################
# Parameters
#############################################################

param (
    [string]$SharePointSiteId = "",   # Graph site ID, e.g. contoso.sharepoint.com,{siteGuid},{webGuid}
    [string]$SharePointListId = "",   # GUID of the source SharePoint list (query queue)
    [string]$DailyListId      = "",   # GUID of the Agent Usage - Daily list
    [string]$WeeklyListId     = "",   # GUID of the Agent Usage - Weekly list
    [string]$MonthlyListId    = "",   # GUID of the Agent Usage - Monthly list
    [string]$AllTimeListId    = "",   # GUID of the Agent Usage - AllTime list

    # Optional: data sources cited in responses
    [string]$DataSourcesUsedListId = "",  # GUID of Agent Data Sources Used list (leave blank to skip)

    # Governance: sharing limit check (optional - leave blank to skip)
    [string]$InventoryListId       = "",  # GUID of Agent Inventory list (enables filtering + owner resolution)
    [string]$NotificationLogListId = "",  # GUID of Notification Log list
    [string]$PolicyRulesListId     = "",  # GUID of Policy Rules list
    [string]$DeletionQueueListId   = "",  # GUID of Deletion Queue list (leave blank to skip)

    # Retention settings
    [int]$DailyRetentionDays   = 7,   # Keep daily rows for this many days
    [int]$WeeklyRetentionWeeks = 13,  # Keep weekly rows for this many weeks

    # App registration auth (optional - leave all blank to use managed identity)
    [string]$TenantId = "",
    [string]$ClientId = "",
    [string]$ClientSecret = "",
    [string]$CertificateThumbprint = ""
)

#############################################################
# Auth Mode Validation
#############################################################

$useSecret      = -not [string]::IsNullOrWhiteSpace($ClientSecret)
$useCert        = -not [string]::IsNullOrWhiteSpace($CertificateThumbprint)
$hasTenantId    = -not [string]::IsNullOrWhiteSpace($TenantId)
$hasClientId    = -not [string]::IsNullOrWhiteSpace($ClientId)
$hasAppRegParam = $useSecret -or $useCert -or $hasTenantId -or $hasClientId

if ($hasAppRegParam) {
    # Validate mutual exclusion
    if ($useSecret -and $useCert) {
        Write-Error "Provide either -ClientSecret OR -CertificateThumbprint, not both."
        exit 1
    }
    # Validate all required app reg params are present
    if (-not $hasTenantId) {
        Write-Error "-TenantId is required when using app registration authentication."
        exit 1
    }
    if (-not $hasClientId) {
        Write-Error "-ClientId is required when using app registration authentication."
        exit 1
    }
    if (-not $useSecret -and -not $useCert) {
        Write-Error "Provide either -ClientSecret or -CertificateThumbprint when using app registration authentication."
        exit 1
    }

    if ($useSecret) {
        $authMode = "AppSecret"
    } else {
        $authMode = "AppCert"
    }
} else {
    $authMode = "ManagedIdentity"
}

Write-Output "Auth mode: $authMode"

#############################################################
# Dependencies
#############################################################

foreach ($moduleName in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Beta.Security')) {
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        try {
            Write-Output "Installing module: $moduleName..."
            Install-Module -Name $moduleName -Force -AllowClobber -Scope CurrentUser
        }
        catch {
            Write-Error "Failed to install module '$moduleName': $_"
            exit 1
        }
    }
    Write-Output "Importing module: $moduleName..."
    Import-Module -Name $moduleName -Force
}

#############################################################
# Functions
#############################################################

# Connect to Microsoft Graph
function ConnectToGraph {
    try {
        switch ($authMode) {
            "ManagedIdentity" {
                Connect-MgGraph -Identity -NoWelcome
            }
            "AppSecret" {
                $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
                $credential   = New-Object System.Management.Automation.PSCredential($ClientId, $secureSecret)
                Connect-MgGraph -ClientSecretCredential $credential -TenantId $TenantId -NoWelcome
            }
            "AppCert" {
                Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome
            }
        }
        Write-Output "Connected to Microsoft Graph."
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph: $_"
        exit 1
    }
}

# Retry logic for Graph API calls (generic for GET, PUT, POST, PATCH)
function Invoke-GraphRequestWithRetry {
    param (
        [Parameter(Mandatory)]
        [string] $Method,
        [Parameter(Mandatory)]
        [string] $Uri,
        [hashtable] $Headers = @{},
        [object] $Body = $null,
        [string] $ContentType = $null,
        [int] $MaxRetries = 8
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $params = @{ Method = $Method; Uri = $Uri; Headers = $Headers }
            if ($null -ne $Body) { $params['Body'] = $Body }
            if ($ContentType) { $params['ContentType'] = $ContentType }

            return Invoke-MgGraphRequest @params
        }
        catch {
            $ex = $_
            $resp = $null
            if ($ex.Exception -and $ex.Exception.Response) { $resp = $ex.Exception.Response }
            if ($resp -and ($resp.StatusCode -eq 429 -or $resp.StatusCode -ge 500)) {
                $retryAfter = [Math]::Min(5 * [Math]::Pow(2, $attempt - 1), 60)
                try { if ($resp.Headers['Retry-After']) { $retryAfter = [int]$resp.Headers['Retry-After'] } } catch {}
                Write-Output "$Method request failed (attempt $attempt/$MaxRetries, HTTP $($resp.StatusCode)), retrying in ${retryAfter}s..."
                Start-Sleep -Seconds $retryAfter
                continue
            }
            throw
        }
    }
    throw "$Method request to $Uri failed after $MaxRetries attempts."
}

# Get status of query
function CheckIfQuerySucceeded {
    param (
        [string]$auditLogQueryId
    )
    try {
        $query = Get-MgBetaSecurityAuditLogQuery -AuditLogQueryId $auditLogQueryId -ErrorAction Stop
        if ($query.status -eq "succeeded") {
            Write-Output "Audit Log Query succeeded."
            return $query
        }
        else {
            Write-Output "Audit Log Query status: $($query.status)"
            Write-Output "Check again later."
            exit 1
        }
    }
    catch {
        Write-Error "Failed to get Audit Log Query: $auditLogQueryId"
        Write-Error "$_"
        exit 1
    }
}

# Script-level variable to hold the list item ID for deletion after processing
$script:ListItemId = $null

# Get the first available query ID from the SharePoint list
function GetAuditQueryIdFromList {
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$SharePointSiteId/lists/$SharePointListId/items?`$expand=fields&`$top=1"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri

        if ($response.value.Count -eq 0) {
            Write-Output "No items found in the SharePoint list."
            exit 1
        }

        $item = $response.value[0]
        $script:ListItemId = $item.id
        $auditLogQueryId   = $item.fields.QueryId

        if ([string]::IsNullOrWhiteSpace($auditLogQueryId)) {
            Write-Error "List item '$($item.id)' has an empty QueryId field."
            exit 1
        }

        Write-Host "Retrieved list item ID: $($script:ListItemId), QueryId: $auditLogQueryId"
        return $auditLogQueryId
    }
    catch {
        Write-Error "Failed to get AuditLogQueryId from list: $_"
        exit 1
    }
}

# Pre-fetch Agent Inventory into a hashtable for O(1) lookups.
# Used to (1) filter audit records to only known/managed agents, and
# (2) resolve owner info for the deletion queue without per-agent API calls.
function Get-InventoryLookup {
    param (
        [Parameter(Mandatory)]
        [string]$siteId,
        [Parameter(Mandatory)]
        [string]$listId
    )

    $lookup = @{}
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listId/items?`$expand=fields(`$select=PackageId,OwnerEmail,OwnerDisplayName)&`$top=200"

    do {
        $response = Invoke-GraphRequestWithRetry -Method GET -Uri $uri
        foreach ($item in $response.value) {
            $packageId = $item.fields.PackageId
            if (-not [string]::IsNullOrWhiteSpace($packageId)) {
                $lookup[$packageId] = @{
                    OwnerEmail       = $item.fields.OwnerEmail
                    OwnerDisplayName = $item.fields.OwnerDisplayName
                }
            }
        }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    return $lookup
}

# Delete the processed list item from the SharePoint list
function DeleteAuditQueryIdFromList {
    try {
        if ([string]::IsNullOrWhiteSpace($script:ListItemId)) {
            Write-Error "No list item ID stored - cannot delete."
            exit 1
        }

        $uri = "https://graph.microsoft.com/v1.0/sites/$SharePointSiteId/lists/$SharePointListId/items/$($script:ListItemId)"
        Invoke-MgGraphRequest -Method DELETE -Uri $uri

        Write-Output "Deleted list item: $($script:ListItemId)"
    }
    catch {
        Write-Error "Failed to delete list item: $_"
        exit 1
    }
}

# Phase 1: Fetch all audit log records and aggregate by agent + period
# Returns a hashtable with four keys: Daily, Weekly, Monthly, AllTime
# Each value is a hashtable keyed by "AgentId|Period"
function GetAggregatedInteractions {
    param (
        [Parameter(Mandatory)]
        [string]$auditLogQueryId,
        [Parameter(Mandatory)]
        [string]$DailyCutoff,
        [Parameter(Mandatory)]
        [string]$WeeklyCutoff,
        [hashtable]$InventoryLookup = $null
    )

    # Four separate hashtables — one per granularity
    $daily   = @{}
    $weekly  = @{}
    $monthly = @{}
    $allTime = @{}
    $dataSources = @{}  # Optional: keyed by "AgentId|Period|SourceUrl"

    $uri = "https://graph.microsoft.com/beta/security/auditLog/queries/$auditLogQueryId/records"
    $rowCount = 0
    $skippedCount = 0
    $skippedAgentIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    # Helper to get or create an aggregation bucket in a given hashtable
    function Get-OrCreateBucket {
        param (
            [hashtable]$table,
            [string]$agentId,
            [string]$agentName,
            [string]$period
        )
        $key = "$agentId|$period"
        if (-not $table.ContainsKey($key)) {
            $table[$key] = @{
                AgentId          = $agentId
                AgentName        = $agentName
                Period           = $period
                InteractionCount = 0
                ActiveUsers      = [System.Collections.Generic.HashSet[string]]::new(
                                       [StringComparer]::OrdinalIgnoreCase)
                LastInteraction  = [DateTime]::MinValue
            }
        }
        return $table[$key]
    }

    # Helper to update a bucket with a new interaction
    function Update-Bucket {
        param (
            [hashtable]$bucket,
            [string]$userId,
            [object]$timestamp
        )
        $bucket.InteractionCount++
        $bucket.ActiveUsers.Add($userId) | Out-Null
        if ($timestamp -is [DateTime] -and $timestamp -gt $bucket.LastInteraction) {
            $bucket.LastInteraction = $timestamp
        }
    }

    do {
        $response = Invoke-GraphRequestWithRetry -Method GET -Uri $uri

        foreach ($item in $response.value) {
            # Extract agent ID and user from the audit record
            $auditData = $item['auditData']
            if ($auditData -is [string]) {
                $auditData = $auditData | ConvertFrom-Json
            }

            # Filter: skip non-agent interactions (no AgentId in auditData)
            if ([string]::IsNullOrWhiteSpace($auditData.AgentId)) { continue }

            # Aggregation key: use TargetPlatformAgentId (matches inventory schema)
            $agentId   = $auditData.CopilotEventData.TargetPlatformAgentId
            if ([string]::IsNullOrWhiteSpace($agentId)) { continue }

            # Filter: skip agents not in the inventory (if inventory lookup is provided)
            if ($InventoryLookup -and -not $InventoryLookup.ContainsKey($agentId)) {
                $skippedCount++
                $skippedAgentIds.Add($agentId) | Out-Null
                continue
            }

            $agentName = $auditData.CopilotEventData.TargetAgentName
            if ([string]::IsNullOrWhiteSpace($agentName)) { $agentName = $auditData.AgentName }

            $userId    = $item['userPrincipalName']
            $timestamp = $item['createdDateTime']

            # Parse timestamp for period calculations
            $ts = if ($timestamp -is [DateTime]) { $timestamp } else { [DateTime]::Parse($timestamp) }

            # Determine period values for each granularity
            $dailyPeriod   = $ts.ToString("yyyy-MM-dd")
            $weekNum        = [System.Globalization.ISOWeek]::GetWeekOfYear($ts)
            $weeklyPeriod  = "{0}-W{1:D2}" -f $ts.Year, $weekNum
            $monthlyPeriod = $ts.ToString("yyyy-MM")

            # Update Daily bucket (only if within retention window)
            if ($dailyPeriod -ge $DailyCutoff) {
                $dailyEntry = Get-OrCreateBucket -table $daily -agentId $agentId -agentName $agentName -period $dailyPeriod
                Update-Bucket -bucket $dailyEntry -userId $userId -timestamp $ts
            }

            # Update Weekly bucket (only if within retention window)
            if ($weeklyPeriod -ge $WeeklyCutoff) {
                $weeklyEntry = Get-OrCreateBucket -table $weekly -agentId $agentId -agentName $agentName -period $weeklyPeriod
                Update-Bucket -bucket $weeklyEntry -userId $userId -timestamp $ts
            }

            # Update Monthly bucket
            $monthlyEntry = Get-OrCreateBucket -table $monthly -agentId $agentId -agentName $agentName -period $monthlyPeriod
            Update-Bucket -bucket $monthlyEntry -userId $userId -timestamp $ts

            # Update AllTime bucket
            $allTimeEntry = Get-OrCreateBucket -table $allTime -agentId $agentId -agentName $agentName -period "AllTime"
            Update-Bucket -bucket $allTimeEntry -userId $userId -timestamp $ts

            # Optional: Extract data sources cited in the response (both Contexts and AccessedResources)
            if ($DataSourcesUsedListId) {
                $allSources = @()
                $contextSources = $auditData.CopilotEventData.Contexts
                if ($contextSources) { $allSources += $contextSources }
                $accessedResources = $auditData.CopilotEventData.AccessedResources
                if ($accessedResources) { $allSources += $accessedResources }

                foreach ($src in $allSources) {
                    $sourceUrl = $src.SiteUrl ?? $src.Id ?? $src.Url ?? $src.Name
                    if ([string]::IsNullOrWhiteSpace($sourceUrl)) { continue }

                    $sourceType = if ($src.Type) { $src.Type } else {
                        switch -Wildcard ($sourceUrl) {
                            "*-my.sharepoint.com*" { "OneDrive" }
                            "*.sharepoint.com*" { "SharePoint" }
                            "http*" { "WebSearch" }
                            default { "Other" }
                        }
                    }

                    $dsKey = "$agentId|$monthlyPeriod|$sourceUrl"
                    if (-not $dataSources.ContainsKey($dsKey)) {
                        $dataSources[$dsKey] = @{
                            AgentId       = $agentId
                            Period        = $monthlyPeriod
                            SourceUrl     = $sourceUrl
                            SourceType    = $sourceType
                            CitationCount = 0
                            LastCited     = [DateTime]::MinValue
                        }
                    }
                    $dataSources[$dsKey].CitationCount++
                    if ($ts -gt $dataSources[$dsKey].LastCited) {
                        $dataSources[$dsKey].LastCited = $ts
                    }
                }
            }

            $rowCount++
        }

        Write-Host "Processed $rowCount records..."
        $uri = $response.'@odata.nextLink'
        if ($uri) { Start-Sleep -Milliseconds 200 }

    } while ($uri)

    Write-Host "Aggregated $rowCount records into $($allTime.Count) agents."
    Write-Host "  Daily:   $($daily.Count) buckets"
    Write-Host "  Weekly:  $($weekly.Count) buckets"
    Write-Host "  Monthly: $($monthly.Count) buckets"
    Write-Host "  AllTime: $($allTime.Count) buckets"
    if ($DataSourcesUsedListId) {
        Write-Host "  DataSources: $($dataSources.Count) unique source citations"
    }
    if ($skippedCount -gt 0) {
        Write-Host "  Skipped $skippedCount records ($($skippedAgentIds.Count) agent(s) not in inventory):"
        foreach ($unknownId in $skippedAgentIds) {
            Write-Host "    - $unknownId"
        }
    }

    return @{
        Daily       = $daily
        Weekly      = $weekly
        Monthly     = $monthly
        AllTime     = $allTime
        DataSources = $dataSources
    }
}

# Phase 2: Batch upsert to target SharePoint list
function WriteAggregationToList {
    param (
        [Parameter(Mandatory)]
        [hashtable]$aggregation,
        [Parameter(Mandatory)]
        [string]$siteId,
        [Parameter(Mandatory)]
        [string]$listId,
        [string]$label = "",
        [int]$BatchSize = 20
    )

    if ($aggregation.Count -eq 0) {
        Write-Output "[$label] No data to write — skipping."
        return
    }

    # 1. Get existing items from the list (keyed by AgentId|Period composite key)
    $existingItems = @{}
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listId/items?`$expand=fields&`$top=200"

    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        foreach ($item in $response.value) {
            $agentId = $item.fields.AgentId
            $period  = $item.fields.Period
            if ($agentId -and $period) {
                $key = "$agentId|$period"
                $existingItems[$key] = @{
                    ItemId           = $item.id
                    InteractionCount = [int]($item.fields.InteractionCount ?? 0)
                    ActiveUsers      = $item.fields.ActiveUsers ?? ""
                }
            }
        }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    Write-Output "[$label] Found $($existingItems.Count) existing items in list."

    # 2. Build individual requests for each agent+period
    $batchRequests = [System.Collections.Generic.List[hashtable]]::new()
    $requestId = 0

    foreach ($compositeKey in $aggregation.Keys) {
        $entry = $aggregation[$compositeKey]
        $requestId++

        if ($existingItems.ContainsKey($compositeKey)) {
            # UPDATE — increment existing counts, merge users
            $existing = $existingItems[$compositeKey]
            $newCount = $existing.InteractionCount + $entry.InteractionCount

            $existingUsers = if ($existing.ActiveUsers) {
                $existing.ActiveUsers -split ";\s*"
            } else { @() }
            $mergedUsers = [System.Collections.Generic.HashSet[string]]::new(
                               $entry.ActiveUsers, [StringComparer]::OrdinalIgnoreCase)
            foreach ($u in $existingUsers) { $mergedUsers.Add($u) | Out-Null }
            $entry.ActiveUsers = $mergedUsers  # Update in-memory with cumulative total
            $mergedCsv = ($mergedUsers | Sort-Object) -join "; "

            $batchRequests.Add(@{
                id      = "$requestId"
                method  = "PATCH"
                url     = "/sites/$siteId/lists/$listId/items/$($existing.ItemId)/fields"
                headers = @{ "Content-Type" = "application/json" }
                body    = @{
                    AgentName        = $entry.AgentName
                    InteractionCount = $newCount
                    ActiveUsers      = $mergedCsv
                    ActiveUserCount  = $mergedUsers.Count
                    LastInteraction  = $entry.LastInteraction.ToString("o")
                }
            })
        }
        else {
            # CREATE
            $usersCsv = ($entry.ActiveUsers | Sort-Object) -join "; "

            $batchRequests.Add(@{
                id      = "$requestId"
                method  = "POST"
                url     = "/sites/$siteId/lists/$listId/items"
                headers = @{ "Content-Type" = "application/json" }
                body    = @{
                    fields = @{
                        Title            = "$($entry.AgentName) — $($entry.Period)"
                        AgentId          = $entry.AgentId
                        AgentName        = $entry.AgentName
                        Period           = $entry.Period
                        InteractionCount = $entry.InteractionCount
                        ActiveUsers      = $usersCsv
                        ActiveUserCount  = $entry.ActiveUsers.Count
                        LastInteraction  = $entry.LastInteraction.ToString("o")
                    }
                }
            })
        }
    }

    # 3. Send in batches of 20
    $totalBatches = [Math]::Ceiling($batchRequests.Count / $BatchSize)
    $batchNum = 0

    for ($i = 0; $i -lt $batchRequests.Count; $i += $BatchSize) {
        $batchNum++
        $chunk = $batchRequests[$i..([Math]::Min($i + $BatchSize - 1, $batchRequests.Count - 1))]

        $batchBody = @{ requests = @($chunk) } | ConvertTo-Json -Depth 10

        $batchResponse = Invoke-GraphRequestWithRetry -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/`$batch" `
            -Body $batchBody `
            -ContentType "application/json"

        # 4. Check individual responses for errors
        $failures = @()
        foreach ($resp in $batchResponse.responses) {
            if ($resp.status -ge 400) {
                $failures += "Request $($resp.id): HTTP $($resp.status) - $($resp.body.error.message)"
            }
        }

        if ($failures.Count -gt 0) {
            Write-Warning "[$label] Batch $batchNum/$totalBatches had $($failures.Count) failures:"
            $failures | ForEach-Object { Write-Warning "  $_" }
        }
        else {
            Write-Output "[$label] Batch $batchNum/${totalBatches}: $($chunk.Count) requests succeeded."
        }

        # Throttle between batches
        if ($i + $BatchSize -lt $batchRequests.Count) {
            Start-Sleep -Milliseconds 200
        }
    }

    Write-Output "[$label] Finished writing $($aggregation.Count) entries in $totalBatches batch(es)."
}

# Phase 3: Delete rows older than a cutoff from a usage list
# Used for rolling retention on Daily and Weekly lists
function CleanupExpiredRows {
    param (
        [Parameter(Mandatory)]
        [string]$siteId,
        [Parameter(Mandatory)]
        [string]$listId,
        [Parameter(Mandatory)]
        [string]$cutoffPeriod,       # Rows with Period < this value are deleted
        [string]$label = "",
        [int]$BatchSize = 20
    )

    # Fetch items with Period < cutoff
    $itemsToDelete = [System.Collections.Generic.List[string]]::new()
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listId/items?`$expand=fields(`$select=Period)&`$top=200"

    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        foreach ($item in $response.value) {
            $period = $item.fields.Period
            if ($period -and $period -lt $cutoffPeriod) {
                $itemsToDelete.Add($item.id)
            }
        }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    if ($itemsToDelete.Count -eq 0) {
        Write-Output "[$label] No expired rows to clean up (cutoff: $cutoffPeriod)."
        return
    }

    Write-Output "[$label] Deleting $($itemsToDelete.Count) expired rows (Period < $cutoffPeriod)..."

    # Batch delete
    $requestId = 0
    $totalBatches = [Math]::Ceiling($itemsToDelete.Count / $BatchSize)
    $batchNum = 0

    for ($i = 0; $i -lt $itemsToDelete.Count; $i += $BatchSize) {
        $batchNum++
        $chunk = $itemsToDelete[$i..([Math]::Min($i + $BatchSize - 1, $itemsToDelete.Count - 1))]

        $requests = @($chunk | ForEach-Object {
            $requestId++
            @{
                id     = "$requestId"
                method = "DELETE"
                url    = "/sites/$siteId/lists/$listId/items/$_"
            }
        })

        $batchBody = @{ requests = $requests } | ConvertTo-Json -Depth 10

        $batchResponse = Invoke-GraphRequestWithRetry -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/`$batch" `
            -Body $batchBody `
            -ContentType "application/json"

        $failures = @($batchResponse.responses | Where-Object { $_.status -ge 400 })
        if ($failures.Count -gt 0) {
            Write-Warning "[$label] Cleanup batch $batchNum/$totalBatches had $($failures.Count) failures."
        }
        else {
            Write-Output "[$label] Cleanup batch $batchNum/${totalBatches}: $($chunk.Count) deletes succeeded."
        }

        if ($i + $BatchSize -lt $itemsToDelete.Count) {
            Start-Sleep -Milliseconds 200
        }
    }

    Write-Output "[$label] Cleanup complete — deleted $($itemsToDelete.Count) expired rows."
}

# Phase 4 (optional): Batch upsert data source citations to the Agent Data Sources Used list
function WriteDataSourcesToList {
    param (
        [Parameter(Mandatory)]
        [hashtable]$sources,
        [Parameter(Mandatory)]
        [string]$siteId,
        [Parameter(Mandatory)]
        [string]$listId,
        [int]$BatchSize = 20
    )

    if ($sources.Count -eq 0) {
        Write-Output "[DataSources] No data source citations to write — skipping."
        return
    }

    # Fetch existing items keyed by AgentId|Period|SourceUrl
    $existingItems = @{}
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listId/items?`$expand=fields&`$top=200"

    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        foreach ($item in $response.value) {
            $f = $item.fields
            $key = "$($f.AgentId)|$($f.Period)|$($f.SourceUrl)"
            $existingItems[$key] = @{
                ItemId        = $item.id
                CitationCount = [int]($f.CitationCount ?? 0)
            }
        }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    Write-Output "[DataSources] Found $($existingItems.Count) existing source rows."

    $batchRequests = [System.Collections.Generic.List[hashtable]]::new()
    $requestId = 0

    foreach ($key in $sources.Keys) {
        $entry = $sources[$key]
        $requestId++

        if ($existingItems.ContainsKey($key)) {
            $existing = $existingItems[$key]
            $batchRequests.Add(@{
                id      = "$requestId"
                method  = "PATCH"
                url     = "/sites/$siteId/lists/$listId/items/$($existing.ItemId)/fields"
                headers = @{ "Content-Type" = "application/json" }
                body    = @{
                    CitationCount = $existing.CitationCount + $entry.CitationCount
                    LastCited     = $entry.LastCited.ToString("o")
                }
            })
        }
        else {
            $batchRequests.Add(@{
                id      = "$requestId"
                method  = "POST"
                url     = "/sites/$siteId/lists/$listId/items"
                headers = @{ "Content-Type" = "application/json" }
                body    = @{
                    fields = @{
                        Title         = "$($entry.AgentId) — $($entry.SourceType)"
                        AgentId       = $entry.AgentId
                        Period        = $entry.Period
                        SourceUrl     = $entry.SourceUrl
                        SourceType    = $entry.SourceType
                        CitationCount = $entry.CitationCount
                        LastCited     = $entry.LastCited.ToString("o")
                    }
                }
            })
        }
    }

    # Send in batches
    $totalBatches = [Math]::Ceiling($batchRequests.Count / $BatchSize)
    $batchNum = 0

    for ($i = 0; $i -lt $batchRequests.Count; $i += $BatchSize) {
        $batchNum++
        $chunk = $batchRequests[$i..([Math]::Min($i + $BatchSize - 1, $batchRequests.Count - 1))]
        $batchBody = @{ requests = @($chunk) } | ConvertTo-Json -Depth 10

        $batchResponse = Invoke-GraphRequestWithRetry -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/`$batch" `
            -Body $batchBody `
            -ContentType "application/json"

        $failures = @($batchResponse.responses | Where-Object { $_.status -ge 400 })
        if ($failures.Count -gt 0) {
            Write-Warning "[DataSources] Batch $batchNum/$totalBatches had $($failures.Count) failures."
        }
        else {
            Write-Output "[DataSources] Batch $batchNum/${totalBatches}: $($chunk.Count) requests succeeded."
        }

        if ($i + $BatchSize -lt $batchRequests.Count) {
            Start-Sleep -Milliseconds 200
        }
    }

    Write-Output "[DataSources] Finished writing $($sources.Count) source citations in $totalBatches batch(es)."
}

# Phase 5: Check sharing limit and write to Notification Log for agents above threshold
# Only writes a notification if one doesn't already exist for that agent (dedup).
# Uses in-memory AllTime data — ActiveUsers is updated to cumulative merged set during Phase 2 upsert.
# Flow 5 handles the Inventory lookup and card send.
function CheckSharingLimitAndNotify {
    param (
        [Parameter(Mandatory)]
        [string]$siteId,
        [Parameter(Mandatory)]
        [hashtable]$allTimeData,
        [Parameter(Mandatory)]
        [string]$notificationLogListId,
        [Parameter(Mandatory)]
        [string]$policyRulesListId,
        [int]$BatchSize = 20
    )

    # 1. Read SharingLimit threshold from Policy Rules
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$policyRulesListId/items?`$expand=fields&`$filter=fields/RuleType eq 'SharingLimit'"
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri

    if ($response.value.Count -eq 0) {
        Write-Output "[Governance] No 'SharingLimit' policy rule found — skipping sharing limit check."
        return
    }

    $sharingLimit = [int]$response.value[0].fields.ThresholdValue
    Write-Output "[Governance] SharingLimit threshold: $sharingLimit"

    # 2. Find agents above the sharing limit from in-memory AllTime data (cumulative after Phase 2 merge)
    $violators = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($entry in $allTimeData.Values) {
        if ($entry.ActiveUsers.Count -gt $sharingLimit) {
            $violators.Add(@{
                AgentId         = $entry.AgentId
                AgentName       = $entry.AgentName
                ActiveUserCount = $entry.ActiveUsers.Count
            })
        }
    }

    if ($violators.Count -eq 0) {
        Write-Output "[Governance] No agents above sharing limit ($sharingLimit) — no notifications needed."
        return
    }

    Write-Output "[Governance] Found $($violators.Count) agent(s) above sharing limit."

    # 3. Fetch existing SharingLimit notifications for dedup
    $existingNotifications = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$notificationLogListId/items?`$expand=fields(`$select=PackageId,NotificationType)&`$filter=fields/NotificationType eq 'SharingLimit'&`$top=200"

    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        foreach ($item in $response.value) {
            $existingNotifications.Add($item.fields.PackageId) | Out-Null
        }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    Write-Output "[Governance] Found $($existingNotifications.Count) existing SharingLimit notification(s) — will skip these."

    # 4. Filter to only new violations
    $newViolators = $violators | Where-Object { -not $existingNotifications.Contains($_.AgentId) }

    if ($newViolators.Count -eq 0) {
        Write-Output "[Governance] All violating agents already have notifications — nothing to do."
        return
    }

    Write-Output "[Governance] Writing $($newViolators.Count) new SharingLimit notification(s) to Notification Log..."

    # 5. Batch-create Notification Log entries
    $batchRequests = [System.Collections.Generic.List[hashtable]]::new()
    $requestId = 0

    foreach ($agent in $newViolators) {
        $requestId++
        $batchRequests.Add(@{
            id      = "$requestId"
            method  = "POST"
            url     = "/sites/$siteId/lists/$notificationLogListId/items"
            headers = @{ "Content-Type" = "application/json" }
            body    = @{
                fields = @{
                    Title            = "Sharing Limit: $($agent.AgentName)"
                    PackageId        = $agent.AgentId
                    NotificationType = "SharingLimit"
                    SentDate         = (Get-Date).ToUniversalTime().ToString("o")
                    ResponseStatus   = "N/A"
                    FlowRunId        = "Script-AggregateToList"
                    Details          = "$($agent.ActiveUserCount) active users"
                }
            }
        })
    }

    # Send in batches
    $totalBatches = [Math]::Ceiling($batchRequests.Count / $BatchSize)
    $batchNum = 0

    for ($i = 0; $i -lt $batchRequests.Count; $i += $BatchSize) {
        $batchNum++
        $chunk = $batchRequests[$i..([Math]::Min($i + $BatchSize - 1, $batchRequests.Count - 1))]
        $batchBody = @{ requests = @($chunk) } | ConvertTo-Json -Depth 10

        $batchResponse = Invoke-GraphRequestWithRetry -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/`$batch" `
            -Body $batchBody `
            -ContentType "application/json"

        $failures = @($batchResponse.responses | Where-Object { $_.status -ge 400 })
        if ($failures.Count -gt 0) {
            Write-Warning "[Governance] Batch $batchNum/$totalBatches had $($failures.Count) failures:"
            foreach ($f in $failures) {
                Write-Warning "  Request $($f.id): HTTP $($f.status) - $($f.body.error.message)"
            }
        }
        else {
            Write-Output "[Governance] Batch $batchNum/${totalBatches}: $($chunk.Count) notifications created."
        }

        if ($i + $BatchSize -lt $batchRequests.Count) {
            Start-Sleep -Milliseconds 200
        }
    }

    Write-Output "[Governance] Sharing limit check complete — $($newViolators.Count) new notification(s) written."
}

# Phase 6: Check agent inactivity and write InactivityWarning / InactivityDeletion to Notification Log.
# Agents past the deletion threshold get InactivityDeletion only (no redundant warning).
# Flow 5 handles the Inventory lookup and card send.
# If DeletionQueue params are provided, deletion-tier agents are also written to the Deletion Queue
# with owner info from Inventory for admin action.
# NOTE: Inactive agents won't appear in the current audit batch, so we must query the AllTime list.
function CheckInactivityAndNotify {
    param (
        [Parameter(Mandatory)]
        [string]$siteId,
        [Parameter(Mandatory)]
        [string]$allTimeListId,
        [Parameter(Mandatory)]
        [string]$notificationLogListId,
        [Parameter(Mandatory)]
        [string]$policyRulesListId,
        [string]$deletionQueueListId = "",
        [hashtable]$inventoryLookup = $null,
        [int]$BatchSize = 20
    )

    # 1. Read InactivityWarningDays and InactivityDeletionDays from Policy Rules
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$policyRulesListId/items?`$expand=fields&`$filter=fields/RuleType eq 'InactivityWarningDays' or fields/RuleType eq 'InactivityDeletionDays'"
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri

    $warningDays  = $null
    $deletionDays = $null
    foreach ($item in $response.value) {
        switch ($item.fields.RuleType) {
            "InactivityWarningDays"  { $warningDays  = [int]$item.fields.ThresholdValue }
            "InactivityDeletionDays" { $deletionDays = [int]$item.fields.ThresholdValue }
        }
    }

    if (-not $warningDays -and -not $deletionDays) {
        Write-Output "[Governance] No inactivity policy rules found — skipping inactivity check."
        return
    }

    # Use warning threshold as the broadest filter (catches both warning and deletion agents)
    $effectiveDays = if ($warningDays) { $warningDays } else { $deletionDays }
    $warningCutoff  = (Get-Date).AddDays(-$effectiveDays).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $deletionCutoff = if ($deletionDays) { (Get-Date).AddDays(-$deletionDays).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") } else { $null }

    Write-Output "[Governance] Inactivity thresholds — Warning: $warningDays days, Deletion: $deletionDays days"
    Write-Output "[Governance] Cutoffs — Warning: $warningCutoff, Deletion: $deletionCutoff"

    # 2. Query AllTime list for agents with LastInteraction at or before the warning cutoff
    #    (inactive agents by definition won't be in the current audit batch's in-memory data)
    $inactiveAgents = [System.Collections.Generic.List[hashtable]]::new()
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$allTimeListId/items?`$expand=fields&`$filter=fields/LastInteraction le '$warningCutoff'&`$top=200"

    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        foreach ($item in $response.value) {
            $inactiveAgents.Add(@{
                AgentId         = $item.fields.AgentId
                AgentName       = $item.fields.AgentName
                LastInteraction = $item.fields.LastInteraction
            })
        }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    if ($inactiveAgents.Count -eq 0) {
        Write-Output "[Governance] No inactive agents found — no notifications needed."
        return
    }

    Write-Output "[Governance] Found $($inactiveAgents.Count) inactive agent(s)."

    # 3. Split into deletion vs warning (deletion takes priority — no double-notify)
    $deletionAgents = [System.Collections.Generic.List[hashtable]]::new()
    $warningAgents  = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($agent in $inactiveAgents) {
        if ($deletionCutoff -and $agent.LastInteraction -le $deletionCutoff) {
            $deletionAgents.Add($agent)
        }
        elseif ($warningDays) {
            $warningAgents.Add($agent)
        }
    }

    Write-Output "[Governance] Breakdown — Deletion: $($deletionAgents.Count), Warning: $($warningAgents.Count)"

    # 4. Fetch existing inactivity notifications for dedup
    $existingWarnings  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $existingDeletions = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    foreach ($notifType in @("InactivityWarning", "InactivityDeletion")) {
        $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$notificationLogListId/items?`$expand=fields(`$select=PackageId,NotificationType)&`$filter=fields/NotificationType eq '$notifType'&`$top=200"
        do {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri
            foreach ($item in $response.value) {
                if ($notifType -eq "InactivityWarning") {
                    $existingWarnings.Add($item.fields.PackageId) | Out-Null
                } else {
                    $existingDeletions.Add($item.fields.PackageId) | Out-Null
                }
            }
            $uri = $response.'@odata.nextLink'
        } while ($uri)
    }

    Write-Output "[Governance] Existing notifications — Warnings: $($existingWarnings.Count), Deletions: $($existingDeletions.Count)"

    # 5. Filter to only new notifications
    $newWarnings  = $warningAgents  | Where-Object { -not $existingWarnings.Contains($_.AgentId) }
    $newDeletions = $deletionAgents | Where-Object { -not $existingDeletions.Contains($_.AgentId) }

    $totalNew = @($newWarnings).Count + @($newDeletions).Count
    if ($totalNew -eq 0) {
        Write-Output "[Governance] All inactive agents already have notifications — nothing to do."
        return
    }

    Write-Output "[Governance] Writing $(@($newWarnings).Count) warning(s) and $(@($newDeletions).Count) deletion(s) to Notification Log..."

    # 6. Build batch requests
    $batchRequests = [System.Collections.Generic.List[hashtable]]::new()
    $requestId = 0
    $now = (Get-Date).ToUniversalTime().ToString("o")

    foreach ($agent in $newWarnings) {
        $requestId++
        $lastAct = if ($agent.LastInteraction) {
            try { ([DateTime]$agent.LastInteraction).ToString("yyyy-MM-dd") } catch { $agent.LastInteraction }
        } else { "Unknown" }
        $batchRequests.Add(@{
            id      = "$requestId"
            method  = "POST"
            url     = "/sites/$siteId/lists/$notificationLogListId/items"
            headers = @{ "Content-Type" = "application/json" }
            body    = @{
                fields = @{
                    Title            = "Inactivity Warning: $($agent.AgentName)"
                    PackageId        = $agent.AgentId
                    NotificationType = "InactivityWarning"
                    SentDate         = $now
                    ResponseStatus   = "N/A"
                    FlowRunId        = "Script-AggregateToList"
                    Details          = $lastAct
                }
            }
        })
    }

    foreach ($agent in $newDeletions) {
        $requestId++
        $lastAct = if ($agent.LastInteraction) {
            try { ([DateTime]$agent.LastInteraction).ToString("yyyy-MM-dd") } catch { $agent.LastInteraction }
        } else { "Unknown" }
        $batchRequests.Add(@{
            id      = "$requestId"
            method  = "POST"
            url     = "/sites/$siteId/lists/$notificationLogListId/items"
            headers = @{ "Content-Type" = "application/json" }
            body    = @{
                fields = @{
                    Title            = "Inactivity Deletion: $($agent.AgentName)"
                    PackageId        = $agent.AgentId
                    NotificationType = "InactivityDeletion"
                    SentDate         = $now
                    ResponseStatus   = "N/A"
                    FlowRunId        = "Script-AggregateToList"
                    Details          = $lastAct
                }
            }
        })
    }

    # 7. Send in batches
    $totalBatches = [Math]::Ceiling($batchRequests.Count / $BatchSize)
    $batchNum = 0

    for ($i = 0; $i -lt $batchRequests.Count; $i += $BatchSize) {
        $batchNum++
        $chunk = $batchRequests[$i..([Math]::Min($i + $BatchSize - 1, $batchRequests.Count - 1))]
        $batchBody = @{ requests = @($chunk) } | ConvertTo-Json -Depth 10

        $batchResponse = Invoke-GraphRequestWithRetry -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/`$batch" `
            -Body $batchBody `
            -ContentType "application/json"

        $failures = @($batchResponse.responses | Where-Object { $_.status -ge 400 })
        if ($failures.Count -gt 0) {
            Write-Warning "[Governance] Batch $batchNum/$totalBatches had $($failures.Count) failures:"
            foreach ($f in $failures) {
                Write-Warning "  Request $($f.id): HTTP $($f.status) - $($f.body.error.message)"
            }
        }
        else {
            Write-Output "[Governance] Batch $batchNum/${totalBatches}: $($chunk.Count) notifications created."
        }

        if ($i + $BatchSize -lt $batchRequests.Count) {
            Start-Sleep -Milliseconds 200
        }
    }

    Write-Output "[Governance] Inactivity check complete — $(@($newWarnings).Count) warning(s), $(@($newDeletions).Count) deletion(s) written."

    # 8. Write deletion-tier agents to Deletion Queue (if configured)
    if (-not $deletionQueueListId -or -not $inventoryLookup -or @($newDeletions).Count -eq 0) {
        if (-not $deletionQueueListId -and @($newDeletions).Count -gt 0) {
            Write-Output "[Governance] DeletionQueueListId not provided — skipping deletion queue write."
        }
        return
    }

    Write-Output "[Governance] Resolving owners from inventory lookup for $(@($newDeletions).Count) deletion agent(s)..."

    # 8a. Resolve owner info from pre-fetched inventory lookup (no additional API calls)
    $ownerLookup = @{}  # keyed by PackageId
    foreach ($agent in $newDeletions) {
        if ($inventoryLookup.ContainsKey($agent.AgentId)) {
            $ownerLookup[$agent.AgentId] = $inventoryLookup[$agent.AgentId]
        }
    }

    Write-Output "[Governance] Resolved owners for $($ownerLookup.Count) of $(@($newDeletions).Count) agent(s)."

    # 8b. Dedup against existing Deletion Queue entries
    $existingQueueItems = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$deletionQueueListId/items?`$expand=fields(`$select=PackageId)&`$filter=fields/Status eq 'Pending'&`$top=200"

    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        foreach ($item in $response.value) {
            $existingQueueItems.Add($item.fields.PackageId) | Out-Null
        }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    $newQueueAgents = @($newDeletions | Where-Object { -not $existingQueueItems.Contains($_.AgentId) })

    if ($newQueueAgents.Count -eq 0) {
        Write-Output "[Governance] All deletion agents already in queue — nothing to add."
        return
    }

    # 8c. Batch-create Deletion Queue entries
    $queueRequests = [System.Collections.Generic.List[hashtable]]::new()
    $requestId = 0
    $now = (Get-Date).ToUniversalTime().ToString("o")

    foreach ($agent in $newQueueAgents) {
        $requestId++
        $owner = $ownerLookup[$agent.AgentId]
        $queueRequests.Add(@{
            id      = "$requestId"
            method  = "POST"
            url     = "/sites/$siteId/lists/$deletionQueueListId/items"
            headers = @{ "Content-Type" = "application/json" }
            body    = @{
                fields = @{
                    Title            = $agent.AgentName
                    PackageId        = $agent.AgentId
                    AgentName        = $agent.AgentName
                    OwnerEmail       = if ($owner) { $owner.OwnerEmail } else { "" }
                    OwnerDisplayName = if ($owner) { $owner.OwnerDisplayName } else { "" }
                    LastInteraction  = $agent.LastInteraction
                    QueuedDate       = $now
                    Status           = "Pending"
                }
            }
        })
    }

    $totalBatches = [Math]::Ceiling($queueRequests.Count / $BatchSize)
    $batchNum = 0

    for ($i = 0; $i -lt $queueRequests.Count; $i += $BatchSize) {
        $batchNum++
        $chunk = $queueRequests[$i..([Math]::Min($i + $BatchSize - 1, $queueRequests.Count - 1))]
        $batchBody = @{ requests = @($chunk) } | ConvertTo-Json -Depth 10

        $batchResponse = Invoke-GraphRequestWithRetry -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/`$batch" `
            -Body $batchBody `
            -ContentType "application/json"

        $failures = @($batchResponse.responses | Where-Object { $_.status -ge 400 })
        if ($failures.Count -gt 0) {
            Write-Warning "[Governance] Deletion Queue batch $batchNum/$totalBatches had $($failures.Count) failures:"
            foreach ($f in $failures) {
                Write-Warning "  Request $($f.id): HTTP $($f.status) - $($f.body.error.message)"
            }
        }
        else {
            Write-Output "[Governance] Deletion Queue batch $batchNum/${totalBatches}: $($chunk.Count) items queued."
        }

        if ($i + $BatchSize -lt $queueRequests.Count) {
            Start-Sleep -Milliseconds 200
        }
    }

    Write-Output "[Governance] Deletion queue complete — $($newQueueAgents.Count) agent(s) queued for admin action."
}

#############################################################
# Main Script Execution
#############################################################

# Connect to Microsoft Graph
ConnectToGraph

# Get query ID from the SharePoint list
$AuditLogQueryId = GetAuditQueryIdFromList
Write-Output "Retrieved AuditLogQueryId from list: $SharePointListId"

# Check if ready to process / download (Exits if not ready)
$query = CheckIfQuerySucceeded -auditLogQueryId $AuditLogQueryId

# Compute retention cutoffs (used for both aggregation filtering and cleanup)
$dailyCutoff  = (Get-Date).AddDays(-$DailyRetentionDays).ToString("yyyy-MM-dd")
$weeklyCutoff = "{0}-W{1:D2}" -f (Get-Date).AddDays(-($WeeklyRetentionWeeks * 7)).Year,
    [System.Globalization.ISOWeek]::GetWeekOfYear((Get-Date).AddDays(-($WeeklyRetentionWeeks * 7)))

Write-Output "Retention cutoffs — Daily: $dailyCutoff, Weekly: $weeklyCutoff"

# Pre-fetch Agent Inventory for filtering and owner resolution (if InventoryListId provided)
$inventoryLookup = $null
if (-not [string]::IsNullOrWhiteSpace($InventoryListId)) {
    $inventoryLookup = Get-InventoryLookup -siteId $SharePointSiteId -listId $InventoryListId
    Write-Output "[Inventory] Loaded $($inventoryLookup.Count) agents from inventory for filtering."
} else {
    Write-Output "[Inventory] No InventoryListId provided — all agents will be recorded (no filtering)."
}

# Phase 1: Aggregate interactions by agent into 4 granularities
$aggregation = GetAggregatedInteractions -auditLogQueryId $AuditLogQueryId -DailyCutoff $dailyCutoff -WeeklyCutoff $weeklyCutoff -InventoryLookup $inventoryLookup

# Phase 2: Batch upsert each granularity to its target list
WriteAggregationToList -aggregation $aggregation.Daily   -siteId $SharePointSiteId -listId $DailyListId   -label "Daily"
WriteAggregationToList -aggregation $aggregation.Weekly  -siteId $SharePointSiteId -listId $WeeklyListId  -label "Weekly"
WriteAggregationToList -aggregation $aggregation.Monthly -siteId $SharePointSiteId -listId $MonthlyListId -label "Monthly"
WriteAggregationToList -aggregation $aggregation.AllTime -siteId $SharePointSiteId -listId $AllTimeListId -label "AllTime"

# Phase 3: Clean up expired daily and weekly rows (handles stale rows from prior runs)

CleanupExpiredRows -siteId $SharePointSiteId -listId $DailyListId  -cutoffPeriod $dailyCutoff  -label "Daily"
CleanupExpiredRows -siteId $SharePointSiteId -listId $WeeklyListId -cutoffPeriod $weeklyCutoff -label "Weekly"

# Phase 4 (optional): Write data source citations
if ($DataSourcesUsedListId) {
    WriteDataSourcesToList -sources $aggregation.DataSources -siteId $SharePointSiteId -listId $DataSourcesUsedListId
}

# Phase 5 (optional): Check sharing limit and notify via Notification Log
if ($NotificationLogListId -and $PolicyRulesListId) {
    CheckSharingLimitAndNotify `
        -siteId $SharePointSiteId `
        -allTimeData $aggregation.AllTime `
        -notificationLogListId $NotificationLogListId `
        -policyRulesListId $PolicyRulesListId
}

# Phase 6 (optional): Check inactivity and notify via Notification Log
if ($NotificationLogListId -and $PolicyRulesListId) {
    CheckInactivityAndNotify `
        -siteId $SharePointSiteId `
        -allTimeListId $AllTimeListId `
        -notificationLogListId $NotificationLogListId `
        -policyRulesListId $PolicyRulesListId `
        -inventoryLookup $inventoryLookup `
        -deletionQueueListId $DeletionQueueListId
}

# Remove item from source list after processing
DeleteAuditQueryIdFromList

Write-Output "Agent interaction data written to all usage lists."
