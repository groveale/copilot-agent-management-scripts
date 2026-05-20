#############################################################
# Script to enrich an Agent Inventory item with full package
# details from the Graph Agent Registry beta API.
# Designed to run as an Azure Automation runbook, called from
# Power Automate Flow 2 via webhook or Create Job action.
#
# Responsibilities:
#   1. Parse package response JSON (passed from Power Automate)
#   2. Resolve OwnerId → OwnerEmail + OwnerDisplayName
#   3. Parse definition JSON → extract capabilities
#   4. Delete existing capability rows for this agent
#   5. Create new capability rows in Agent Capabilities list
#   6. Update Agent Inventory with detail fields + NeedsEnrichment = false
#
# Supports managed identity (default), app registration + secret,
# or app registration + certificate.
#
# Permissions needed:
#   - Microsoft Graph: User.Read.All (Application), Sites.Selected or Sites.FullControl.All
#   - Note: CopilotPackages.Read.All is NOT needed — the package response is passed in from the flow
#
# Contact alexgrover@microsoft.com for questions
#############################################################

#############################################################
# Parameters
#############################################################

param (
    [Parameter(Mandatory)]
    [object]$PackageResponseJson,     # Full JSON response from /beta/copilot/admin/catalog/packages/{packageId} — passed from Power Automate (accepts string or deserialized object)

    [Parameter(Mandatory)]
    [int]$ItemId,                     # SharePoint list item ID for the Agent Inventory row

    [string]$SharePointSiteId = "",        # Graph site ID, e.g. contoso.sharepoint.com,{siteGuid},{webGuid}

    [string]$InventoryListId = "",         # GUID of Agent Inventory list

    [string]$CapabilitiesListId = "",      # GUID of Agent Capabilities list

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
    if ($useSecret -and $useCert) {
        Write-Error "Provide either -ClientSecret OR -CertificateThumbprint, not both."
        exit 1
    }
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

foreach ($moduleName in @('Microsoft.Graph.Authentication')) {
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
# Classes
#############################################################

class CapabilityRow {
    [string]$Title
    [string]$ParentPackageId
    [string]$AgentName
    [string]$CapabilityType
    [string]$SourceType
    [string]$SourceValue
    [bool]$IncludeRelatedContent
}

#############################################################
# Functions
#############################################################

function ConnectToGraph {
    try {
        switch ($script:authMode) {
            "ManagedIdentity" {
                Connect-MgGraph -Identity -NoWelcome
            }
            "AppSecret" {
                $secureSecret = ConvertTo-SecureString $script:ClientSecret -AsPlainText -Force
                $credential   = New-Object System.Management.Automation.PSCredential($script:ClientId, $secureSecret)
                Connect-MgGraph -ClientSecretCredential $credential -TenantId $script:TenantId -NoWelcome
            }
            "AppCert" {
                Connect-MgGraph -ClientId $script:ClientId -TenantId $script:TenantId -CertificateThumbprint $script:CertificateThumbprint -NoWelcome
            }
        }
        Write-Output "Connected to Microsoft Graph."
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph: $_"
        exit 1
    }
}

function Resolve-OwnerIdentity {
    param ([string]$OwnerId)
    if ([string]::IsNullOrWhiteSpace($OwnerId)) {
        Write-Output "No OwnerId provided — skipping owner resolution."
        return $null
    }
    try {
        $uri = "https://graph.microsoft.com/v1.0/users/$OwnerId"
        $user = Invoke-MgGraphRequest -Method GET -Uri $uri
        Write-Output "Resolved owner: $($user.userPrincipalName)"
        return @{
            Email       = $user.userPrincipalName
            DisplayName = $user.displayName
        }
    }
    catch {
        Write-Warning "Failed to resolve OwnerId '$OwnerId': $_. Continuing without owner details."
        return $null
    }
}

function ConvertTo-CapabilityRows {
    param (
        [object[]]$Capabilities,
        [string]$PackageId,
        [string]$AgentName
    )

    $rows = [System.Collections.Generic.List[CapabilityRow]]::new()

    foreach ($cap in $Capabilities) {
        $capName = $cap.name
        $hasRows = $false

        switch ($capName) {
            "OneDriveAndSharePoint" {
                if ($cap.items_by_url) {
                    foreach ($item in $cap.items_by_url) {
                        $rows.Add([CapabilityRow]@{
                            Title                 = $capName
                            ParentPackageId       = $PackageId
                            AgentName             = $AgentName
                            CapabilityType        = $capName
                            SourceType            = "SharePointUrl"
                            SourceValue           = $item.url
                            IncludeRelatedContent = $false
                        })
                        $hasRows = $true
                    }
                }
                if ($cap.items_by_sharepoint_ids) {
                    foreach ($item in $cap.items_by_sharepoint_ids) {
                        if ($item.'x-is_embedded' -eq $true) {
                            # Embedded file surfaced via ODSP reference — route to EmbeddedKnowledge
                            $rows.Add([CapabilityRow]@{
                                Title                 = "EmbeddedKnowledge"
                                ParentPackageId       = $PackageId
                                AgentName             = $AgentName
                                CapabilityType        = "EmbeddedKnowledge"
                                SourceType            = "EmbeddedSharePointItem"
                                SourceValue           = "$($item.site_id)/$($item.list_id)/$($item.unique_id)"
                                IncludeRelatedContent = $false
                            })
                        } else {
                            $rows.Add([CapabilityRow]@{
                                Title                 = $capName
                                ParentPackageId       = $PackageId
                                AgentName             = $AgentName
                                CapabilityType        = $capName
                                SourceType            = "SharePointItem"
                                SourceValue           = "$($item.site_id)/$($item.list_id)/$($item.unique_id)"
                                IncludeRelatedContent = $false
                            })
                        }
                        $hasRows = $true
                    }
                }
            }
            "WebSearch" {
                if ($cap.sites) {
                    foreach ($item in $cap.sites) {
                        $rows.Add([CapabilityRow]@{
                            Title                 = $capName
                            ParentPackageId       = $PackageId
                            AgentName             = $AgentName
                            CapabilityType        = $capName
                            SourceType            = "WebSearchSite"
                            SourceValue           = $item.url
                            IncludeRelatedContent = $false
                        })
                        $hasRows = $true
                    }
                }
            }
            "GraphConnectors" {
                if ($cap.connections) {
                    foreach ($item in $cap.connections) {
                        $rows.Add([CapabilityRow]@{
                            Title                 = $capName
                            ParentPackageId       = $PackageId
                            AgentName             = $AgentName
                            CapabilityType        = $capName
                            SourceType            = "GraphConnector"
                            SourceValue           = $item.connection_id
                            IncludeRelatedContent = $false
                        })
                        $hasRows = $true
                    }
                }
            }
            "Email" {
                if ($cap.shared_mailbox) {
                    $rows.Add([CapabilityRow]@{
                        Title                 = $capName
                        ParentPackageId       = $PackageId
                        AgentName             = $AgentName
                        CapabilityType        = $capName
                        SourceType            = "EmailAddress"
                        SourceValue           = $cap.shared_mailbox
                        IncludeRelatedContent = $false
                    })
                    $hasRows = $true
                }
                if ($cap.group_mailboxes) {
                    foreach ($mailbox in $cap.group_mailboxes) {
                        $rows.Add([CapabilityRow]@{
                            Title                 = $capName
                            ParentPackageId       = $PackageId
                            AgentName             = $AgentName
                            CapabilityType        = $capName
                            SourceType            = "GroupMailbox"
                            SourceValue           = $mailbox
                            IncludeRelatedContent = $false
                        })
                        $hasRows = $true
                    }
                }
            }
            "TeamsMessages" {
                if ($cap.urls) {
                    foreach ($item in $cap.urls) {
                        $rows.Add([CapabilityRow]@{
                            Title                 = $capName
                            ParentPackageId       = $PackageId
                            AgentName             = $AgentName
                            CapabilityType        = $capName
                            SourceType            = "TeamsChannel"
                            SourceValue           = $item.url
                            IncludeRelatedContent = $false
                        })
                        $hasRows = $true
                    }
                }
            }
            "Dataverse" {
                if ($cap.knowledge_sources) {
                    foreach ($item in $cap.knowledge_sources) {
                        $rows.Add([CapabilityRow]@{
                            Title                 = $capName
                            ParentPackageId       = $PackageId
                            AgentName             = $AgentName
                            CapabilityType        = $capName
                            SourceType            = "Dataverse"
                            SourceValue           = $item.host_name
                            IncludeRelatedContent = $false
                        })
                        $hasRows = $true
                    }
                }
            }
            "EmbeddedKnowledge" {
                if ($cap.files) {
                    foreach ($item in $cap.files) {
                        $rows.Add([CapabilityRow]@{
                            Title                 = $capName
                            ParentPackageId       = $PackageId
                            AgentName             = $AgentName
                            CapabilityType        = $capName
                            SourceType            = "EmbeddedFile"
                            SourceValue           = $item.file
                            IncludeRelatedContent = $false
                        })
                        $hasRows = $true
                    }
                }
            }
            "Meetings" {
                if ($cap.items_by_id) {
                    foreach ($item in $cap.items_by_id) {
                        $rows.Add([CapabilityRow]@{
                            Title                 = $capName
                            ParentPackageId       = $PackageId
                            AgentName             = $AgentName
                            CapabilityType        = $capName
                            SourceType            = "MeetingItem"
                            SourceValue           = $item.id
                            IncludeRelatedContent = $false
                        })
                        $hasRows = $true
                    }
                }
            }
            "People" {
                $rows.Add([CapabilityRow]@{
                    Title                 = $capName
                    ParentPackageId       = $PackageId
                    AgentName             = $AgentName
                    CapabilityType        = $capName
                    SourceType            = $null
                    SourceValue           = $null
                    IncludeRelatedContent = [bool]($cap.include_related_content)
                })
                $hasRows = $true
            }
            "ScenarioModels" {
                if ($cap.models) {
                    foreach ($item in $cap.models) {
                        $rows.Add([CapabilityRow]@{
                            Title                 = $capName
                            ParentPackageId       = $PackageId
                            AgentName             = $AgentName
                            CapabilityType        = $capName
                            SourceType            = "ScenarioModel"
                            SourceValue           = $item.id
                            IncludeRelatedContent = $false
                        })
                        $hasRows = $true
                    }
                }
            }
            default {
                # Future capability types (GraphicArt, CodeInterpreter, etc.)
                $rows.Add([CapabilityRow]@{
                    Title                 = $capName
                    ParentPackageId       = $PackageId
                    AgentName             = $AgentName
                    CapabilityType        = $capName
                    SourceType            = $null
                    SourceValue           = $null
                    IncludeRelatedContent = $false
                })
                $hasRows = $true
            }
        }

        # If a known capability had no sources (defensive)
        if (-not $hasRows) {
            $rows.Add([CapabilityRow]@{
                Title                 = $capName
                ParentPackageId       = $PackageId
                AgentName             = $AgentName
                CapabilityType        = $capName
                SourceType            = $null
                SourceValue           = $null
                IncludeRelatedContent = $false
            })
        }
    }

    return $rows
}

function Remove-ExistingCapabilities {
    param (
        [string]$SiteId,
        [string]$ListId,
        [string]$PackageId
    )
    try {
        $filter = "fields/ParentPackageId eq '$PackageId'"
        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$ListId/items?`$filter=$filter&`$select=id"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        $items = $response.value

        if ($items.Count -eq 0) {
            Write-Output "No existing capability rows to delete for PackageId: $PackageId"
            return
        }

        Write-Output "Deleting $($items.Count) existing capability rows..."
        foreach ($item in $items) {
            $deleteUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$ListId/items/$($item.id)"
            Invoke-MgGraphRequest -Method DELETE -Uri $deleteUri | Out-Null
        }
        Write-Output "Deleted all existing capability rows."
    }
    catch {
        Write-Error "Failed to delete existing capabilities: $_"
        exit 1
    }
}

function Write-CapabilityRows {
    param (
        [string]$SiteId,
        [string]$ListId,
        [CapabilityRow[]]$Rows
    )
    $now = (Get-Date).ToUniversalTime().ToString("o")
    $created = 0

    foreach ($row in $Rows) {
        $fields = @{
            Title                 = $row.Title
            ParentPackageId       = $row.ParentPackageId
            AgentName             = $row.AgentName
            CapabilityType        = $row.CapabilityType
            IncludeRelatedContent = $row.IncludeRelatedContent
            LastSynced            = $now
        }
        if ($row.SourceType)  { $fields["SourceType"]  = $row.SourceType }
        if ($row.SourceValue) { $fields["SourceValue"] = $row.SourceValue }

        $body = @{ fields = $fields } | ConvertTo-Json -Depth 10
        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$ListId/items"
        Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType "application/json" | Out-Null
        $created++
    }

    Write-Output "Created $created capability rows."
}

function Update-InventoryItem {
    param (
        [string]$SiteId,
        [string]$ListId,
        [int]$ItemId,
        [hashtable]$Fields
    )
    $Fields["LastSynced"] = (Get-Date).ToUniversalTime().ToString("o")
    $Fields["NeedsEnrichment"] = $false

    # Remove null values — Graph API for SharePoint rejects PATCH when Note fields receive null
    $cleanFields = @{}
    foreach ($key in $Fields.Keys) {
        if ($null -ne $Fields[$key]) {
            $cleanFields[$key] = $Fields[$key]
        }
    }

    $body = $cleanFields | ConvertTo-Json -Depth 10
    $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$ListId/items/$ItemId/fields"
    Write-Output "DEBUG — PATCH URI: $uri"
    Write-Output "DEBUG — PATCH body: $body"
    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $body -ContentType "application/json" | Out-Null
        Write-Output "Updated inventory item $ItemId — NeedsEnrichment set to false."
    }
    catch {
        Write-Error "PATCH failed for item $ItemId`: $_"
        throw
    }
}

#############################################################
# Main Script Execution
#############################################################

Write-Output "Script started."
Write-Output "ItemId: $ItemId"

# Step 1: Parse the package response JSON (passed from Power Automate)
# With [object] parameter type, Azure Automation typically receives a deserialized
# PSCustomObject directly. If it arrives as a JSON string, we parse it.
Write-Output "PackageResponseJson type: $($PackageResponseJson.GetType().FullName)"
try {
    if ($PackageResponseJson -is [string]) {
        $package = $PackageResponseJson | ConvertFrom-Json
    } else {
        $package = $PackageResponseJson
    }
    Write-Output "Parsed package response for: $($package.id) ($($package.displayName))"
}
catch {
    Write-Error "Failed to parse PackageResponseJson: $_"
    exit 1
}

$packageId = $package.id

# Step 2: Extract and parse the definition field + element ID
# The definition is a JSON string nested inside elementDetails[].elements[].definition
# The element ID is used to construct deep links to the agent in M365
$definitionJson = $null
if ($package.elementDetails) {
    foreach ($detail in $package.elementDetails) {
        if ($detail.elements) {
            foreach ($element in $detail.elements) {
                if ($element.definition) {
                    $definitionJson = $element.definition
                    break
                }
            }
        }
        if ($definitionJson) { break }
    }
}

if (-not $definitionJson) {
    Write-Warning "No definition found in elementDetails — skipping capability parsing."
    $definition = $null
} elseif ($definitionJson -is [string]) {
    $definition = $definitionJson | ConvertFrom-Json
} else {
    $definition = $definitionJson
}

# Connect to Microsoft Graph (needed for owner resolution + SharePoint writes)
ConnectToGraph

# Step 3: Resolve owner identity
$ownerId = $package.ownerId
$owner = Resolve-OwnerIdentity -OwnerId $ownerId

# Step 4: Build capability rows
$agentName = $package.displayName
$capabilities = if ($definition) { $definition.capabilities } else { $null }
$capabilityRows = @()

if ($capabilities) {
    $capabilityRows = ConvertTo-CapabilityRows -Capabilities $capabilities -PackageId $packageId -AgentName $agentName
    Write-Output "Parsed $($capabilityRows.Count) capability rows."
} else {
    Write-Output "No capabilities found in definition."
}

# Step 5: Delete existing capability rows
Remove-ExistingCapabilities -SiteId $SharePointSiteId -ListId $CapabilitiesListId -PackageId $packageId

# Step 6: Write new capability rows
if ($capabilityRows.Count -gt 0) {
    Write-CapabilityRows -SiteId $SharePointSiteId -ListId $CapabilitiesListId -Rows $capabilityRows
}

# Step 7: Build inventory update fields
$inventoryFields = @{}

# Owner fields
if ($owner) {
    $inventoryFields["OwnerEmail"]       = $owner.Email
    $inventoryFields["OwnerDisplayName"] = $owner.DisplayName
}

# Package metadata fields (from top-level package response)
if ($package.shortDescription)    { $inventoryFields["ShortDescription"]  = $package.shortDescription }
if ($package.longDescription)     { $inventoryFields["LongDescription"]   = $package.longDescription }
if ($package.publisher)           { $inventoryFields["Publisher"]          = $package.publisher }
if ($package.version)             { $inventoryFields["AgentVersion"]      = $package.version }
if ($package.manifestVersion)     { $inventoryFields["ManifestVersion"]   = $package.manifestVersion }
if ($package.platform)            { $inventoryFields["Platform"]          = $package.platform }
if ($package.elementTypes)        { $inventoryFields["AgentType"]         = ($package.elementTypes -join ", ") }
if ($null -ne $package.isBlocked) { $inventoryFields["IsBlocked"]         = $package.isBlocked }
if ($package.availableTo)         { $inventoryFields["AvailableTo"]       = $package.availableTo }
if ($package.deployedTo)          { $inventoryFields["DeployedTo"]        = $package.deployedTo }

# Fields from parsed definition (if available)
if ($definition) {
    if ($definition.sensitivity_label) { $inventoryFields["SensitivityLabel"] = $definition.sensitivity_label.id }
}

# Element ID (for constructing deep links downstream)
if ($definition -and $definition.id) {
    $inventoryFields["ElementId"] = $definition.id
    Write-Output "Extracted element ID: $($definition.id)"
} else {
    Write-Warning "No element ID found in definition."
}

# Governance flags
$hasActions      = if ($definition) { ($null -ne $definition.actions -and $definition.actions.Count -gt 0) } else { $false }
$hasWorkerAgents = if ($definition) { ($null -ne $definition.worker_agents -and $definition.worker_agents.Count -gt 0) } else { $false }
$containsEmbeddedFiles = ($capabilityRows | Where-Object { $_.CapabilityType -eq "EmbeddedKnowledge" }).Count -gt 0
$inventoryFields["HasActions"]           = $hasActions
$inventoryFields["HasWorkerAgents"]      = $hasWorkerAgents
$inventoryFields["ContainsEmbeddedFiles"] = $containsEmbeddedFiles

# Capability summary (comma-separated list of capability names)
if ($capabilities) {
    $inventoryFields["CapabilitySummary"] = ($capabilities | ForEach-Object { $_.name }) -join ", "
}

# Instructions (system prompt) + instruction hash for change detection
$instructionsText = if ($definition -and $definition.instructions) { $definition.instructions } else { $null }
$inventoryFields["Instructions"] = $instructionsText

if ($instructionsText) {
    $instructionBytes = [System.Text.Encoding]::UTF8.GetBytes($instructionsText)
    $hashBytes = [System.Security.Cryptography.SHA256]::Create().ComputeHash($instructionBytes)
    $newHash = [BitConverter]::ToString($hashBytes).Replace("-", "").ToLower()
} else {
    $newHash = $null
}
$inventoryFields["InstructionHash"] = $newHash

# Get current hash from inventory item to detect changes
$currentItemUri = "https://graph.microsoft.com/v1.0/sites/$SharePointSiteId/lists/$InventoryListId/items/$ItemId`?`$select=id&`$expand=fields(`$select=InstructionHash)"
try {
    $currentItem = Invoke-MgGraphRequest -Method GET -Uri $currentItemUri
    $previousHash = $currentItem.fields.InstructionHash
} catch {
    $previousHash = $null
}

# If hash changed, flag for prompt scan (compliance flow will pick this up)
if ($previousHash -ne $newHash -and $instructionsText) {
    Write-Output "Instruction hash changed — setting NeedsPromptScan = true."
    $inventoryFields["NeedsPromptScan"] = $true
} elseif ($previousHash -and -not $instructionsText) {
    Write-Output "Instructions removed from agent — clearing hash and prompt scan flag."
    $inventoryFields["NeedsPromptScan"] = $false
} else {
    Write-Output "Instruction hash unchanged — no prompt scan needed."
}

# Flag for compliance flow to pick up (covers both description + instruction checks)
$inventoryFields["NeedsComplianceCheck"] = $true

# Step 8: Update inventory item
Update-InventoryItem -SiteId $SharePointSiteId -ListId $InventoryListId -ItemId $ItemId -Fields $inventoryFields

Write-Output "Enrichment complete for agent '$agentName' (PackageId: $packageId)."
