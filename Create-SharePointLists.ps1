#############################################################
# PnP PowerShell script to create all SharePoint lists for
# the Agent Registry Sync and Agent Usage Sync solutions.
#
# Idempotent — safe to re-run. Skips lists/fields that already exist.
#
# Lists created:
#   1. Agent Inventory          — parent list for agent registry data
#   2. Agent Capabilities       — child list for agent capability details
#   3. Query Queue              — queue for Purview audit log query IDs
#   4. Agent Usage - Daily      — rolling 7-day daily interaction data
#   5. Agent Usage - Weekly     — rolling 13-week weekly interaction data
#   6. Agent Usage - Monthly    — indefinite monthly interaction data
#   7. Agent Usage - AllTime    — lifetime totals per agent (one row per agent)
#   8. Agent Data Sources Used  — (optional) sources cited in responses, from audit log context
#   9. Risky Prompt Criteria    — keyword list for system prompt/description scanning
#  10. Policy Rules             — governance comms engine thresholds (seeded with defaults)
#  11. Notification Log         — governance comms engine audit trail
#  12. Data Owner Mapping       — site URL to data owner mapping (for row-level access)
#  13. Notification Templates   — Adaptive Card JSON templates keyed by notification type
#  14. Pilot Users              — UPN and ID of users in the pilot program
#  15. Deletion Queue           — agents flagged for deletion, actioned by an admin
#
# Note: Instruction compliance is handled via the NeedsPromptScan flag on Agent Inventory
#       (no separate queue list needed).
#
# Prerequisites:
#   Install-Module PnP.PowerShell -Scope CurrentUser
#
# Usage:
#   .\Create-SharePointLists.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/YourSite"
#   .\Create-SharePointLists.ps1 -SiteUrl "..." -IncludeLists "AgentInventory","AgentCapabilities"
#   .\Create-SharePointLists.ps1 -SiteUrl "..." -ExcludeLists "DataSourcesUsed","DataOwnerMapping"
#
# Contact alexgrover@microsoft.com for questions
#############################################################

param (
    [string]$SiteUrl = "https://m365cpi77517573.sharepoint.com/sites/AgentGuard",
    [string]$ClientId = "7491c129-aafb-4c37-b26c-7386ddce3b4c",

    # Control which lists to create. If IncludeLists is specified, ONLY those are created.
    # If ExcludeLists is specified, all EXCEPT those are created. Cannot use both.
    [ValidateSet(
        "AgentInventory",
        "AgentCapabilities",
        "QueryQueue",
        "UsageDaily",
        "UsageWeekly",
        "UsageMonthly",
        "UsageAllTime",
        "DataSourcesUsed",
        "RiskyPromptCriteria",
        "PolicyRules",
        "NotificationLog",
        "DataOwnerMapping",
        "NotificationTemplates",
        "PilotUsers",
        "DeletionQueue"
    )]
    [string[]]$IncludeLists = @(),

    [ValidateSet(
        "AgentInventory",
        "AgentCapabilities",
        "QueryQueue",
        "UsageDaily",
        "UsageWeekly",
        "UsageMonthly",
        "UsageAllTime",
        "DataSourcesUsed",
        "RiskyPromptCriteria",
        "PolicyRules",
        "NotificationLog",
        "DataOwnerMapping",
        "NotificationTemplates",
        "PilotUsers",
        "DeletionQueue"
    )]
    [string[]]$ExcludeLists = @()
)

#############################################################
# Validate params
#############################################################

if ($IncludeLists.Count -gt 0 -and $ExcludeLists.Count -gt 0) {
    Write-Error "Specify either -IncludeLists or -ExcludeLists, not both."
    exit 1
}

$allListKeys = @(
    "AgentInventory",
    "AgentCapabilities",
    "QueryQueue",
    "UsageDaily",
    "UsageWeekly",
    "UsageMonthly",
    "UsageAllTime",
    "DataSourcesUsed",
    "RiskyPromptCriteria",
    "PolicyRules",
    "NotificationLog",
    "DataOwnerMapping",
    "NotificationTemplates",
    "PilotUsers",
    "DeletionQueue"
)

if ($IncludeLists.Count -gt 0) {
    $enabledLists = $IncludeLists
} elseif ($ExcludeLists.Count -gt 0) {
    $enabledLists = $allListKeys | Where-Object { $_ -notin $ExcludeLists }
} else {
    $enabledLists = $allListKeys
}

function ShouldCreate([string]$key) { return $key -in $enabledLists }

#############################################################
# Connect
#############################################################

Write-Output "Connecting to $SiteUrl..."
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive

#############################################################
# Idempotent helpers
#############################################################

function Ensure-List {
    param ([string]$Title)
    $existing = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Output "List '$Title' already exists — skipping creation."
    } else {
        New-PnPList -Title $Title -Template GenericList -ErrorAction Stop
        Write-Output "Created list: $Title"
    }
}

function Ensure-Field {
    param (
        [string]$List,
        [string]$DisplayName,
        [string]$InternalName,
        [string]$Type,
        [switch]$Required,
        [string[]]$Choices
    )
    $existing = Get-PnPField -List $List -Identity $InternalName -ErrorAction SilentlyContinue
    if ($existing) { return }

    $params = @{
        List         = $List
        DisplayName  = $DisplayName
        InternalName = $InternalName
        Type         = $Type
    }
    if ($Required) { $params["Required"] = $true }
    if ($Choices)  { $params["Choices"]  = $Choices }

    Add-PnPField @params | Out-Null
}

function Ensure-Index {
    param ([string]$List, [string]$Field)
    try {
        Set-PnPField -List $List -Identity $Field -Values @{Indexed=$true} -ErrorAction SilentlyContinue
    } catch { }
}

#############################################################
# 1. Agent Inventory
#############################################################

if (ShouldCreate "AgentInventory") {
    $list = "Agent Inventory"
    Write-Output "`n--- $list ---"
    Ensure-List -Title $list

    # Core fields
    Ensure-Field -List $list -DisplayName "PackageId" -InternalName "PackageId" -Type Text -Required
    Ensure-Field -List $list -DisplayName "ShortDescription" -InternalName "ShortDescription" -Type Text
    Ensure-Field -List $list -DisplayName "LongDescription" -InternalName "LongDescription" -Type Note
    Ensure-Field -List $list -DisplayName "Publisher" -InternalName "Publisher" -Type Text
    Ensure-Field -List $list -DisplayName "Version" -InternalName "AgentVersion" -Type Text
    Ensure-Field -List $list -DisplayName "ManifestVersion" -InternalName "ManifestVersion" -Type Text
    Ensure-Field -List $list -DisplayName "Platform" -InternalName "Platform" -Type Text
    Ensure-Field -List $list -DisplayName "AgentType" -InternalName "AgentType" -Type Text

    # Governance fields
    Ensure-Field -List $list -DisplayName "IsBlocked" -InternalName "IsBlocked" -Type Boolean
    Ensure-Field -List $list -DisplayName "AvailableTo" -InternalName "AvailableTo" -Type Text
    Ensure-Field -List $list -DisplayName "DeployedTo" -InternalName "DeployedTo" -Type Text
    Ensure-Field -List $list -DisplayName "LastModified" -InternalName "AgentLastModified" -Type DateTime

    # Definition fields
    Ensure-Field -List $list -DisplayName "SensitivityLabel" -InternalName "SensitivityLabel" -Type Text
    Ensure-Field -List $list -DisplayName "HasActions" -InternalName "HasActions" -Type Boolean
    Ensure-Field -List $list -DisplayName "HasWorkerAgents" -InternalName "HasWorkerAgents" -Type Boolean
    Ensure-Field -List $list -DisplayName "ContainsEmbeddedFiles" -InternalName "ContainsEmbeddedFiles" -Type Boolean
    Ensure-Field -List $list -DisplayName "CapabilitySummary" -InternalName "CapabilitySummary" -Type Text

    # Flow control
    Ensure-Field -List $list -DisplayName "NeedsEnrichment" -InternalName "NeedsEnrichment" -Type Boolean

    # Owner fields (for governance comms engine + row-level access)
    Ensure-Field -List $list -DisplayName "OwnerId" -InternalName "OwnerId" -Type Text
    Ensure-Field -List $list -DisplayName "OwnerEmail" -InternalName "OwnerEmail" -Type Text
    Ensure-Field -List $list -DisplayName "OwnerDisplayName" -InternalName "OwnerDisplayName" -Type Text
    Ensure-Field -List $list -DisplayName "LastAttestationDate" -InternalName "LastAttestationDate" -Type DateTime

    # Description compliance (true = non-compliant / fail)
    Ensure-Field -List $list -DisplayName "IsDescriptionNonCompliant" -InternalName "IsDescriptionNonCompliant" -Type Boolean

    # Compliance check flag (set by enrichment when instructions change)
    Ensure-Field -List $list -DisplayName "NeedsComplianceCheck" -InternalName "NeedsComplianceCheck" -Type Boolean

    # Instruction hash (for change detection → compliance queue)
    Ensure-Field -List $list -DisplayName "InstructionHash" -InternalName "InstructionHash" -Type Text

    # Risky prompt scanning fields
    Ensure-Field -List $list -DisplayName "Instructions" -InternalName "Instructions" -Type Note
    Ensure-Field -List $list -DisplayName "IsRiskyPrompt" -InternalName "IsRiskyPrompt" -Type Boolean
    Ensure-Field -List $list -DisplayName "RiskyPromptMatches" -InternalName "RiskyPromptMatches" -Type Note
    Ensure-Field -List $list -DisplayName "NeedsPromptScan" -InternalName "NeedsPromptScan" -Type Boolean

    # Deep link building block
    Ensure-Field -List $list -DisplayName "ElementId" -InternalName "ElementId" -Type Text

    # Metadata
    Ensure-Field -List $list -DisplayName "LastSynced" -InternalName "LastSynced" -Type DateTime

    # Index for upsert performance
    Ensure-Index -List $list -Field "PackageId"

    Write-Output "Done: $list"
}

#############################################################
# 2. Agent Capabilities
#############################################################

if (ShouldCreate "AgentCapabilities") {
    $child = "Agent Capabilities"
    Write-Output "`n--- $child ---"
    Ensure-List -Title $child

    Ensure-Field -List $child -DisplayName "ParentPackageId" -InternalName "ParentPackageId" -Type Text -Required
    Ensure-Field -List $child -DisplayName "AgentName" -InternalName "AgentName" -Type Text
    Ensure-Field -List $child -DisplayName "CapabilityType" -InternalName "CapabilityType" -Type Choice -Choices @(
        "WebSearch",
        "OneDriveAndSharePoint",
        "GraphConnectors",
        "GraphicArt",
        "CodeInterpreter",
        "EmbeddedKnowledge",
        "Email",
        "People",
        "Meetings",
        "TeamsMessages",
        "Dataverse",
        "ScenarioModels"
    )
    Ensure-Field -List $child -DisplayName "SourceType" -InternalName "SourceType" -Type Choice -Choices @(
        "SharePointUrl",
        "SharePointItem",
        "OneDriveItem",
        "EmailAddress",
        "GroupMailbox",
        "WebSearchSite",
        "MeetingItem",
        "GraphConnector",
        "EmbeddedFile",
        "EmbeddedSharePointItem",
        "TeamsChannel",
        "Dataverse",
        "ScenarioModel"
    )
    Ensure-Field -List $child -DisplayName "SourceValue" -InternalName "SourceValue" -Type Text
    Ensure-Field -List $child -DisplayName "IncludeRelatedContent" -InternalName "IncludeRelatedContent" -Type Boolean
    Ensure-Field -List $child -DisplayName "LastSynced" -InternalName "LastSynced" -Type DateTime

    # Indexes for query performance
    Ensure-Index -List $child -Field "ParentPackageId"
    Ensure-Index -List $child -Field "SourceType"
    Ensure-Index -List $child -Field "SourceValue"

    Write-Output "Done: $child"
}

#############################################################
# 3. Query Queue
#############################################################

if (ShouldCreate "QueryQueue") {
    $queue = "Query Queue"
    Write-Output "`n--- $queue ---"
    Ensure-List -Title $queue

    Ensure-Field -List $queue -DisplayName "QueryId" -InternalName "QueryId" -Type Text -Required

    Write-Output "Done: $queue"
}

#############################################################
# 4. Agent Usage — 4 separate lists by granularity
#############################################################

$usageMap = @{
    "UsageDaily"   = "Agent Usage - Daily"
    "UsageWeekly"  = "Agent Usage - Weekly"
    "UsageMonthly" = "Agent Usage - Monthly"
    "UsageAllTime" = "Agent Usage - AllTime"
}

foreach ($key in $usageMap.Keys) {
    if (-not (ShouldCreate $key)) { continue }

    $usage = $usageMap[$key]
    Write-Output "`n--- $usage ---"
    Ensure-List -Title $usage

    Ensure-Field -List $usage -DisplayName "AgentId" -InternalName "AgentId" -Type Text -Required
    Ensure-Field -List $usage -DisplayName "AgentName" -InternalName "AgentName" -Type Text
    Ensure-Field -List $usage -DisplayName "Period" -InternalName "Period" -Type Text -Required
    Ensure-Field -List $usage -DisplayName "PeriodSort" -InternalName "PeriodSort" -Type Text
    Ensure-Field -List $usage -DisplayName "InteractionCount" -InternalName "InteractionCount" -Type Number
    Ensure-Field -List $usage -DisplayName "ActiveUserCount" -InternalName "ActiveUserCount" -Type Number
    Ensure-Field -List $usage -DisplayName "ActiveUsers" -InternalName "ActiveUsers" -Type Note
    Ensure-Field -List $usage -DisplayName "LastInteraction" -InternalName "LastInteraction" -Type DateTime

    Ensure-Index -List $usage -Field "AgentId"
    Ensure-Index -List $usage -Field "Period"
    Ensure-Index -List $usage -Field "LastInteraction"

    Write-Output "Done: $usage"
}

#############################################################
# 5. Agent Data Sources Used (optional — audit log context)
#############################################################

if (ShouldCreate "DataSourcesUsed") {
    $dsUsed = "Agent Data Sources Used"
    Write-Output "`n--- $dsUsed ---"
    Ensure-List -Title $dsUsed

    Ensure-Field -List $dsUsed -DisplayName "AgentId" -InternalName "AgentId" -Type Text -Required
    Ensure-Field -List $dsUsed -DisplayName "Period" -InternalName "Period" -Type Text -Required
    Ensure-Field -List $dsUsed -DisplayName "SourceUrl" -InternalName "SourceUrl" -Type Note
    Ensure-Field -List $dsUsed -DisplayName "SourceType" -InternalName "SourceType" -Type Choice -Choices @(
        "SharePoint",
        "OneDrive",
        "GraphConnector",
        "WebSearch",
        "EmbeddedKnowledge",
        "Email",
        "Other"
    )
    Ensure-Field -List $dsUsed -DisplayName "CitationCount" -InternalName "CitationCount" -Type Number
    Ensure-Field -List $dsUsed -DisplayName "LastCited" -InternalName "LastCited" -Type DateTime

    Ensure-Index -List $dsUsed -Field "AgentId"
    Ensure-Index -List $dsUsed -Field "Period"

    Write-Output "Done: $dsUsed"
}

#############################################################
# 6. Risky Prompt Criteria
#############################################################

if (ShouldCreate "RiskyPromptCriteria") {
    $criteria = "Risky Prompt Criteria"
    Write-Output "`n--- $criteria ---"
    Ensure-List -Title $criteria

    Ensure-Field -List $criteria -DisplayName "Keyword" -InternalName "Keyword" -Type Text -Required
    Ensure-Field -List $criteria -DisplayName "Category" -InternalName "Category" -Type Text
    Ensure-Field -List $criteria -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean
    Ensure-Field -List $criteria -DisplayName "Notes" -InternalName "Notes" -Type Note

    # Seed with BA's confirmed criteria (only if list was just created / empty)
    $existingItems = Get-PnPListItem -List $criteria -PageSize 1 -ErrorAction SilentlyContinue
    if (-not $existingItems -or $existingItems.Count -eq 0) {
        $seedKeywords = @(
            @{ Title = "Performance Management";  Keyword = "Performance Management";  Category = "HR"; IsActive = $true },
            @{ Title = "Performance Measurement";  Keyword = "Performance Measurement";  Category = "HR"; IsActive = $true },
            @{ Title = "Financials";               Keyword = "Financials";               Category = "Finance"; IsActive = $true },
            @{ Title = "ICFR";                     Keyword = "ICFR";                     Category = "Finance"; IsActive = $true },
            @{ Title = "Safety";                   Keyword = "Safety";                   Category = "Operations"; IsActive = $true; Notes = "Broad term — consider qualifying" },
            @{ Title = "Security";                 Keyword = "Security";                 Category = "Cyber"; IsActive = $true; Notes = "Broad term — may catch legitimate agents" },
            @{ Title = "Cyber";                    Keyword = "Cyber";                    Category = "Cyber"; IsActive = $true },
            @{ Title = "Health";                   Keyword = "Health";                   Category = "Medical"; IsActive = $true },
            @{ Title = "Healthcare";               Keyword = "Healthcare";               Category = "Medical"; IsActive = $true },
            @{ Title = "Medical";                  Keyword = "Medical";                  Category = "Medical"; IsActive = $true },
            @{ Title = "Revenue Management";       Keyword = "Revenue Management";       Category = "Finance"; IsActive = $true },
            @{ Title = "IAG AI Policy";            Keyword = "IAG AI Policy";            Category = "Policy"; IsActive = $true }
        )
        foreach ($kw in $seedKeywords) {
            Add-PnPListItem -List $criteria -Values $kw | Out-Null
        }
        Write-Output "Seeded $($seedKeywords.Count) keywords."
    } else {
        Write-Output "Seed data already exists — skipping."
    }

    Write-Output "Done: $criteria"
}

#############################################################
# 7. Policy Rules
#############################################################

if (ShouldCreate "PolicyRules") {
    $rules = "Policy Rules"
    Write-Output "`n--- $rules ---"
    Ensure-List -Title $rules

    Ensure-Field -List $rules -DisplayName "RuleType" -InternalName "RuleType" -Type Choice -Choices @(
        "InactivityWarningDays",
        "InactivityDeletionDays",
        "SharingLimit"
    ) -Required
    Ensure-Field -List $rules -DisplayName "ThresholdValue" -InternalName "ThresholdValue" -Type Number -Required
    Ensure-Field -List $rules -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean
    Ensure-Field -List $rules -DisplayName "NotificationTemplate" -InternalName "NotificationTemplate" -Type Note

    Ensure-Index -List $rules -Field "RuleType"

    # Seed default rules (only if list is empty)
    $existingItems = Get-PnPListItem -List $rules -PageSize 1 -ErrorAction SilentlyContinue
    if (-not $existingItems -or $existingItems.Count -eq 0) {
        Add-PnPListItem -List $rules -Values @{
            Title          = "Inactivity warning — 90 days"
            RuleType       = "InactivityWarningDays"
            ThresholdValue = 90
            IsActive       = $true
        } | Out-Null
        Add-PnPListItem -List $rules -Values @{
            Title          = "Inactivity deletion — 120 days"
            RuleType       = "InactivityDeletionDays"
            ThresholdValue = 120
            IsActive       = $true
        } | Out-Null
        Add-PnPListItem -List $rules -Values @{
            Title          = "Sharing limit — 50 users"
            RuleType       = "SharingLimit"
            ThresholdValue = 50
            IsActive       = $true
        } | Out-Null
        Write-Output "Seeded 3 default rules."
    } else {
        Write-Output "Seed data already exists — skipping."
    }

    Write-Output "Done: $rules"
}

#############################################################
# 8. Notification Log
#############################################################

if (ShouldCreate "NotificationLog") {
    $log = "Notification Log"
    Write-Output "`n--- $log ---"
    Ensure-List -Title $log

    Ensure-Field -List $log -DisplayName "PackageId" -InternalName "PackageId" -Type Text -Required
    Ensure-Field -List $log -DisplayName "NotificationType" -InternalName "NotificationType" -Type Choice -Choices @(
        "Blocked",
        "Unblocked",
        "InactivityWarning",
        "Attestation",
        "UsageThreshold",
        "SharingLimit",
        "RiskyPromptFlag",
        "DescriptionNonCompliant",
        "EmbeddedFileBlock"
    ) -Required
    Ensure-Field -List $log -DisplayName "RecipientEmail" -InternalName "RecipientEmail" -Type Text
    Ensure-Field -List $log -DisplayName "SentDate" -InternalName "SentDate" -Type DateTime
    Ensure-Field -List $log -DisplayName "ResponseStatus" -InternalName "ResponseStatus" -Type Choice -Choices @(
        "Pending",
        "Acknowledged",
        "Rejected",
        "Expired",
        "N/A"
    )
    Ensure-Field -List $log -DisplayName "ResponseDate" -InternalName "ResponseDate" -Type DateTime
    Ensure-Field -List $log -DisplayName "FlowRunId" -InternalName "FlowRunId" -Type Text
    Ensure-Field -List $log -DisplayName "Details" -InternalName "Details" -Type Text

    Ensure-Index -List $log -Field "PackageId"
    Ensure-Index -List $log -Field "NotificationType"

    Write-Output "Done: $log"
}

#############################################################
# 9. Data Owner Mapping
#############################################################

if (ShouldCreate "DataOwnerMapping") {
    $ownerMap = "Data Owner Mapping"
    Write-Output "`n--- $ownerMap ---"
    Ensure-List -Title $ownerMap

    Ensure-Field -List $ownerMap -DisplayName "SiteUrl" -InternalName "SiteUrl" -Type Text -Required
    Ensure-Field -List $ownerMap -DisplayName "OwnerEmail" -InternalName "OwnerEmail" -Type Text -Required
    Ensure-Field -List $ownerMap -DisplayName "OwnerDisplayName" -InternalName "OwnerDisplayName" -Type Text
    Ensure-Field -List $ownerMap -DisplayName "LastVerified" -InternalName "LastVerified" -Type DateTime

    Ensure-Index -List $ownerMap -Field "SiteUrl"
    Ensure-Index -List $ownerMap -Field "OwnerEmail"

    Write-Output "Done: $ownerMap"
}

#############################################################
# 10. Notification Templates
#############################################################

if (ShouldCreate "NotificationTemplates") {
    $templates = "Notification Templates"
    Write-Output "`n--- $templates ---"
    Ensure-List -Title $templates

    Ensure-Field -List $templates -DisplayName "NotificationType" -InternalName "NotificationType" -Type Choice -Choices @(
        "Blocked",
        "Unblocked",
        "InactivityWarning",
        "Attestation",
        "UsageThreshold",
        "SharingLimit",
        "RiskyPromptFlag",
        "DescriptionNonCompliant",
        "EmbeddedFileBlock"
    ) -Required
    Ensure-Field -List $templates -DisplayName "Description" -InternalName "Description" -Type Note
    Ensure-Field -List $templates -DisplayName "ImageUrl" -InternalName "ImageUrl" -Type Text
    Ensure-Field -List $templates -DisplayName "AdaptiveCardJson" -InternalName "AdaptiveCardJson" -Type Note

    Ensure-Index -List $templates -Field "NotificationType"

    # Seed compliance notification templates (only if list is empty)
    $existingItems = Get-PnPListItem -List $templates -PageSize 1 -ErrorAction SilentlyContinue
    if (-not $existingItems -or $existingItems.Count -eq 0) {

        # --- DescriptionNonCompliant ---
        $descNonCompliantCard = @'
{
  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "type": "AdaptiveCard",
  "version": "1.4",
  "body": [
    {
      "type": "TextBlock",
      "text": "⚠️ Compliance Issue — Invalid Description",
      "weight": "Bolder",
      "size": "Large",
      "color": "Warning"
    },
    {
      "type": "TextBlock",
      "text": "Hi ${OwnerDisplayName}, your agent's description does not meet compliance requirements and has been flagged for review.",
      "wrap": true
    },
    {
      "type": "FactSet",
      "facts": [
        { "title": "Agent", "value": "${AgentName}" },
        { "title": "Package ID", "value": "${PackageId}" },
        { "title": "Platform", "value": "${Platform}" },
        { "title": "Flagged on", "value": "${FlaggedDate}" }
      ]
    },
    {
      "type": "TextBlock",
      "text": "Your agent's description must meet all of the following rules:",
      "wrap": true,
      "spacing": "Medium",
      "weight": "Bolder"
    },
    {
      "type": "TextBlock",
      "text": "• Minimum 10 characters\n• At least 2 words\n• No placeholder text (e.g. test, tbd, n/a, placeholder, asdf, none, abc)",
      "wrap": true
    },
    {
      "type": "TextBlock",
      "text": "Your agent has been blocked. Update your agent's description to a meaningful summary of what it does — the agent will be automatically re-checked and unblocked if compliant.",
      "wrap": true,
      "spacing": "Medium"
    }
  ],
  "actions": [
    {
      "type": "Action.OpenUrl",
      "title": "Edit Agent",
      "url": "https://m365.cloud.microsoft/chat/agent/edit/${PackageId}.${ElementId}"
    }
  ]
}
'@
        Add-PnPListItem -List $templates -Values @{
            Title            = "Description Non-Compliant"
            NotificationType = "DescriptionNonCompliant"
            Description      = "Sent when an agent's description is too short, a single word, or contains placeholder text (e.g. test, tbd, n/a). Owner should update to a meaningful description."
            ImageUrl         = ""
            AdaptiveCardJson = $descNonCompliantCard
        } | Out-Null

        # --- RiskyPromptFlag ---
        $riskyPromptCard = @'
{
  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "type": "AdaptiveCard",
  "version": "1.4",
  "body": [
    {
      "type": "TextBlock",
      "text": "🚨 Compliance Issue — Risky Prompt Detected",
      "weight": "Bolder",
      "size": "Large",
      "color": "Attention"
    },
    {
      "type": "TextBlock",
      "text": "Hi ${OwnerDisplayName}, your agent's instructions or description contain flagged keywords and has been referred for governance review.",
      "wrap": true
    },
    {
      "type": "FactSet",
      "facts": [
        { "title": "Agent", "value": "${AgentName}" },
        { "title": "Package ID", "value": "${PackageId}" },
        { "title": "Platform", "value": "${Platform}" },
        { "title": "Matched keywords", "value": "${RiskyPromptMatches}" },
        { "title": "Flagged on", "value": "${FlaggedDate}" }
      ]
    },
    {
      "type": "TextBlock",
      "text": "Your agent has been blocked. Contact the governance team to discuss your agent's use case, then update your agent — it will be automatically re-checked and unblocked if compliant.",
      "wrap": true,
      "spacing": "Medium"
    }
  ],
  "actions": [
    {
      "type": "Action.OpenUrl",
      "title": "Edit Agent",
      "url": "https://m365.cloud.microsoft/chat/agent/edit/${PackageId}.${ElementId}"
    }
  ]
}
'@
        Add-PnPListItem -List $templates -Values @{
            Title            = "Risky Prompt Detected"
            NotificationType = "RiskyPromptFlag"
            Description      = "Sent when an agent's instructions or description match keywords from the Risky Prompt Criteria list (e.g. ICFR, Security, Performance Management). Flagged for governance review."
            ImageUrl         = ""
            AdaptiveCardJson = $riskyPromptCard
        } | Out-Null

        # --- EmbeddedFileBlock ---
        $embeddedFileCard = @'
{
  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "type": "AdaptiveCard",
  "version": "1.4",
  "body": [
    {
      "type": "TextBlock",
      "text": "📁 Compliance Issue — Embedded Files Detected",
      "weight": "Bolder",
      "size": "Large",
      "color": "Warning"
    },
    {
      "type": "TextBlock",
      "text": "Hi ${OwnerDisplayName}, your agent contains embedded knowledge files. Under current governance policy, agents with uploaded files should not be shared broadly.",
      "wrap": true
    },
    {
      "type": "FactSet",
      "facts": [
        { "title": "Agent", "value": "${AgentName}" },
        { "title": "Package ID", "value": "${PackageId}" },
        { "title": "Platform", "value": "${Platform}" },
        { "title": "Flagged on", "value": "${FlaggedDate}" }
      ]
    },
    {
      "type": "TextBlock",
      "text": "Your agent has been blocked. Remove the embedded files and use SharePoint or OneDrive as the knowledge source instead — the agent will be automatically re-checked and unblocked if compliant. Contact the governance team if you need guidance.",
      "wrap": true,
      "spacing": "Medium"
    }
  ],
  "actions": [
    {
      "type": "Action.OpenUrl",
      "title": "Edit Agent",
      "url": "https://m365.cloud.microsoft/chat/agent/edit/${PackageId}.${ElementId}"
    }
  ]
}
'@
        Add-PnPListItem -List $templates -Values @{
            Title            = "Embedded Files Detected"
            NotificationType = "EmbeddedFileBlock"
            Description      = "Sent when an agent has EmbeddedKnowledge capability rows. Policy requires knowledge sources to use SharePoint/OneDrive rather than uploaded files."
            ImageUrl         = ""
            AdaptiveCardJson = $embeddedFileCard
        } | Out-Null

        # --- Unblocked ---
        $unblockedCard = @'
{
  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "type": "AdaptiveCard",
  "version": "1.4",
  "body": [
    {
      "type": "TextBlock",
      "text": "✅ Agent Unblocked",
      "weight": "Bolder",
      "size": "Large",
      "color": "Good"
    },
    {
      "type": "TextBlock",
      "text": "Hi ${OwnerDisplayName}, thanks for updating your agent — it has passed compliance checks and has been unblocked.",
      "wrap": true
    },
    {
      "type": "FactSet",
      "facts": [
        { "title": "Agent", "value": "${AgentName}" },
        { "title": "Package ID", "value": "${PackageId}" },
        { "title": "Platform", "value": "${Platform}" },
        { "title": "Unblocked on", "value": "${FlaggedDate}" }
      ]
    },
    {
      "type": "TextBlock",
      "text": "No further action is required. Your agent is now available to users again.",
      "wrap": true,
      "spacing": "Medium"
    }
  ],
  "actions": [
    {
      "type": "Action.OpenUrl",
      "title": "Open Agent",
      "url": "https://m365.cloud.microsoft/chat/agent/${PackageId}.${ElementId}"
    }
  ]
}
'@
        Add-PnPListItem -List $templates -Values @{
            Title            = "Agent Unblocked"
            NotificationType = "Unblocked"
            Description      = "Sent when a previously blocked agent passes compliance re-check and is automatically unblocked."
            ImageUrl         = ""
            AdaptiveCardJson = $unblockedCard
        } | Out-Null

        # --- SharingLimit ---
        $sharingLimitCard = @'
{
  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "type": "AdaptiveCard",
  "version": "1.4",
  "body": [
    {
      "type": "TextBlock",
      "text": "👥 Usage Notice — Sharing Limit Exceeded",
      "weight": "Bolder",
      "size": "Large",
      "color": "Warning"
    },
    {
      "type": "TextBlock",
      "text": "Hi ${OwnerDisplayName}, your agent has exceeded the active user sharing limit. This is not a compliance block — your agent will remain active.",
      "wrap": true
    },
    {
      "type": "FactSet",
      "facts": [
        { "title": "Agent", "value": "${AgentName}" },
        { "title": "Package ID", "value": "${PackageId}" },
        { "title": "Platform", "value": "${Platform}" },
        { "title": "Active users", "value": "${Details}" },
        { "title": "Flagged on", "value": "${FlaggedDate}" }
      ]
    },
    {
      "type": "TextBlock",
      "text": "Your agent will not be blocked, however IT will be in contact with you to discuss onboarding details for broad-use agents. This ensures appropriate support, monitoring, and governance is in place.",
      "wrap": true,
      "spacing": "Medium"
    }
  ],
  "actions": [
    {
      "type": "Action.OpenUrl",
      "title": "Open Agent",
      "url": "https://m365.cloud.microsoft/chat/agent/${PackageId}.${ElementId}"
    }
  ]
}
'@
        Add-PnPListItem -List $templates -Values @{
            Title            = "Sharing Limit Exceeded"
            NotificationType = "SharingLimit"
            Description      = "Sent when an agent's active user count exceeds the sharing limit threshold. Informational — agent remains active. IT will follow up for onboarding."
            ImageUrl         = ""
            AdaptiveCardJson = $sharingLimitCard
        } | Out-Null

        # --- InactivityWarning ---
        $inactivityWarningCard = @'
{
  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "type": "AdaptiveCard",
  "version": "1.4",
  "body": [
    {
      "type": "TextBlock",
      "text": "💤 Inactivity Notice — Agent Unused for 90 Days",
      "weight": "Bolder",
      "size": "Large",
      "color": "Warning"
    },
    {
      "type": "TextBlock",
      "text": "Hi ${OwnerDisplayName}, your agent has had no interactions for 90 days and is scheduled for cleanup.",
      "wrap": true
    },
    {
      "type": "FactSet",
      "facts": [
        { "title": "Agent", "value": "${AgentName}" },
        { "title": "Package ID", "value": "${PackageId}" },
        { "title": "Platform", "value": "${Platform}" },
        { "title": "Last activity", "value": "${Details}" },
        { "title": "Flagged on", "value": "${FlaggedDate}" }
      ]
    },
    {
      "type": "TextBlock",
      "text": "If inactivity continues, this agent will be automatically deleted in 30 days. To keep your agent, simply use it or contact the governance team to request an exemption.",
      "wrap": true,
      "spacing": "Medium"
    }
  ],
  "actions": [
    {
      "type": "Action.OpenUrl",
      "title": "Open Agent",
      "url": "https://m365.cloud.microsoft/chat/agent/${PackageId}.${ElementId}"
    }
  ]
}
'@
        Add-PnPListItem -List $templates -Values @{
            Title            = "Inactivity Warning (90 Days)"
            NotificationType = "InactivityWarning"
            Description      = "Sent when an agent has had no interactions for 90 days. Warns owner that the agent will be deleted in 30 days if inactivity continues."
            ImageUrl         = ""
            AdaptiveCardJson = $inactivityWarningCard
        } | Out-Null

        # --- InactivityDeletion ---
        $inactivityDeletionCard = @'
{
  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "type": "AdaptiveCard",
  "version": "1.4",
  "body": [
    {
      "type": "TextBlock",
      "text": "🗑️ Agent Flagged for Deletion — 120 Days of Inactivity",
      "weight": "Bolder",
      "size": "Large",
      "color": "Attention"
    },
    {
      "type": "TextBlock",
      "text": "Hi ${OwnerDisplayName}, your agent has been flagged for deletion due to 120 days of inactivity and will be removed shortly.",
      "wrap": true
    },
    {
      "type": "FactSet",
      "facts": [
        { "title": "Agent", "value": "${AgentName}" },
        { "title": "Package ID", "value": "${PackageId}" },
        { "title": "Platform", "value": "${Platform}" },
        { "title": "Last activity", "value": "${Details}" },
        { "title": "Flagged on", "value": "${FlaggedDate}" }
      ]
    },
    {
      "type": "TextBlock",
      "text": "This agent has had no interactions for over 120 days and has been flagged for deletion as part of the tenant lifecycle policy. If you believe this is in error, please contact IT before the agent is removed.",
      "wrap": true,
      "spacing": "Medium"
    }
  ],
  "actions": [
    {
      "type": "Action.OpenUrl",
      "title": "Open Agent",
      "url": "https://m365.cloud.microsoft/chat/agent/${PackageId}.${ElementId}"
    }
  ]
}
'@
        Add-PnPListItem -List $templates -Values @{
            Title            = "Agent Flagged for Deletion (Inactivity)"
            NotificationType = "InactivityDeletion"
            Description      = "Sent when an agent is flagged for deletion after 120 days of inactivity. Owner is notified before removal."
            ImageUrl         = ""
            AdaptiveCardJson = $inactivityDeletionCard
        } | Out-Null

        Write-Output "Seeded 7 notification templates."
    } else {
        Write-Output "Seed data already exists — skipping."
    }

    Write-Output "Done: $templates"
}

#############################################################
# 14. Pilot Users
#############################################################

if (ShouldCreate "PilotUsers") {
    $pilot = "Pilot Users"
    Write-Output "`n--- $pilot ---"
    Ensure-List -Title $pilot

    Ensure-Field -List $pilot -DisplayName "UPN" -InternalName "UPN" -Type Text -Required
    Ensure-Field -List $pilot -DisplayName "UserId" -InternalName "UserId" -Type Text -Required

    Ensure-Index -List $pilot -Field "UPN"
    Ensure-Index -List $pilot -Field "UserId"

    Write-Output "Done: $pilot"
}

#############################################################
# 15. Deletion Queue
#############################################################

if (ShouldCreate "DeletionQueue") {
    $delQueue = "Deletion Queue"
    Write-Output "`n--- $delQueue ---"
    Ensure-List -Title $delQueue

    Ensure-Field -List $delQueue -DisplayName "PackageId" -InternalName "PackageId" -Type Text -Required
    Ensure-Field -List $delQueue -DisplayName "AgentName" -InternalName "AgentName" -Type Text
    Ensure-Field -List $delQueue -DisplayName "OwnerEmail" -InternalName "OwnerEmail" -Type Text
    Ensure-Field -List $delQueue -DisplayName "OwnerDisplayName" -InternalName "OwnerDisplayName" -Type Text
    Ensure-Field -List $delQueue -DisplayName "LastInteraction" -InternalName "LastInteraction" -Type DateTime
    Ensure-Field -List $delQueue -DisplayName "QueuedDate" -InternalName "QueuedDate" -Type DateTime
    Ensure-Field -List $delQueue -DisplayName "Status" -InternalName "Status" -Type Choice -Choices @(
        "Pending",
        "Completed",
        "Cancelled"
    )
    Ensure-Field -List $delQueue -DisplayName "ActionedBy" -InternalName "ActionedBy" -Type User
    Ensure-Field -List $delQueue -DisplayName "ActionedDate" -InternalName "ActionedDate" -Type DateTime

    Ensure-Index -List $delQueue -Field "PackageId"
    Ensure-Index -List $delQueue -Field "Status"

    Write-Output "Done: $delQueue"
}

#############################################################
# Summary
#############################################################

Write-Output ""
Write-Output "=========================================="
Write-Output "Provisioning complete on: $SiteUrl"
Write-Output "=========================================="
Write-Output "Enabled lists: $($enabledLists -join ', ')"
Write-Output ""
Write-Output "Next steps:"
Write-Output "  - To retrieve list IDs, run:"
Write-Output "    Get-PnPList | Select-Object Title, Id | Format-Table"
Write-Output "  - Populate 'Data Owner Mapping' with site URL → owner entries"
Write-Output "  - Schedule Set-AgentInventoryPermissions.ps1 weekly"
