<#
.SYNOPSIS
    Provisions the ConferenceEvents list with site columns, content types, and views.

.DESCRIPTION
    Creates a single Modern custom list to track Conference Events with four
    logical event types implemented via three SharePoint content types:
      - Conference
      - Call Milestone (Open or Close, for Speakers or Sponsors)
      - Session

    Includes a self-referencing lookup so child events point back to their parent
    Conference row.

.NOTES
    Requires PnP.PowerShell module.
    Run Connect-PnPOnline -Url <site> -Interactive before executing.

    Tested with PnP.PowerShell 2.x. Group prefix "SF" used throughout for site
    columns and content types so they're easy to find in the gallery.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ListTitle = "ConferenceEvents",

    [Parameter(Mandatory = $false)]
    [string]$ListUrl = "Lists/ConferenceEvents",

    [Parameter(Mandatory = $false)]
    [string]$SiteColumnGroup = "Solution Foundry - Conference Events",

    [Parameter(Mandatory = $false)]
    [string]$ContentTypeGroup = "Solution Foundry - Conference Events"
)

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# 0. Sanity check connection
# ---------------------------------------------------------------------------
try {
    $ctx = Get-PnPContext
    Write-Host "Connected to: $($ctx.Url)" -ForegroundColor Cyan
}
catch {
    throw "No active PnP connection. Run Connect-PnPOnline first."
}

# ---------------------------------------------------------------------------
# 1. Create the list (Generic List, will get a Calendar view added later)
# ---------------------------------------------------------------------------
Write-Host "`n[1/6] Creating list '$ListTitle'..." -ForegroundColor Yellow

$list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
if (-not $list) {
    $list = New-PnPList -Title $ListTitle -Url $ListUrl -Template GenericList -OnQuickLaunch
    Write-Host "  Created list: $ListTitle"
}
else {
    Write-Host "  List already exists, continuing..."
}

# Enable management of content types on the list
Set-PnPList -Identity $ListTitle -EnableContentTypes $true
Write-Host "  Content type management enabled."

# ---------------------------------------------------------------------------
# 2. Create site columns
# ---------------------------------------------------------------------------
Write-Host "`n[2/6] Creating site columns..." -ForegroundColor Yellow

# Helper to skip if already exists
function New-PnPFieldIfMissing {
    param(
        [string]$InternalName,
        [scriptblock]$CreateScript
    )
    $existing = Get-PnPField -Identity $InternalName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  - $InternalName already exists, skipping."
        return $existing
    }
    $field = & $CreateScript
    Write-Host "  + $InternalName created."
    return $field
}

# 2a. Event start/end - we use the Event content type's built-in EventDate/EndDate.
# Adding "EventDate" and "EndDate" comes for free when we attach the Event CT.

# 2b. sfWebsite - Hyperlink
New-PnPFieldIfMissing -InternalName "sfWebsite" -CreateScript {
    Add-PnPField -DisplayName "Website" -InternalName "sfWebsite" `
        -Type URL -Group $SiteColumnGroup
}

# 2c. sfConferenceImage - Image (Modern image column, type 34)
# PnP doesn't expose "Image" as a Type enum; create via XML.
New-PnPFieldIfMissing -InternalName "sfConferenceImage" -CreateScript {
    $imageFieldXml = @"
<Field
    Type="Thumbnail"
    DisplayName="Conference Image"
    StaticName="sfConferenceImage"
    Name="sfConferenceImage"
    Group="$SiteColumnGroup"
    Required="FALSE" />
"@
    Add-PnPFieldFromXml -FieldXml $imageFieldXml
}

# 2d. sfCallType - Choice: Speakers, Sponsors
New-PnPFieldIfMissing -InternalName "sfCallType" -CreateScript {
    Add-PnPField -DisplayName "Call Type" -InternalName "sfCallType" `
        -Type Choice -Choices "Speakers", "Sponsors" -Group $SiteColumnGroup
}

# 2e. sfCallPhase - Choice: Open, Close
New-PnPFieldIfMissing -InternalName "sfCallPhase" -CreateScript {
    Add-PnPField -DisplayName "Call Phase" -InternalName "sfCallPhase" `
        -Type Choice -Choices "Open", "Close" -Group $SiteColumnGroup
}

# 2f. sfSubmissionStatus - Choice: Planned, Submitted
New-PnPFieldIfMissing -InternalName "sfSubmissionStatus" -CreateScript {
    Add-PnPField -DisplayName "Submission Status" -InternalName "sfSubmissionStatus" `
        -Type Choice -Choices "Planned", "Submitted" -Group $SiteColumnGroup
}

# 2g. sfSubmittedDate - Date only
New-PnPFieldIfMissing -InternalName "sfSubmittedDate" -CreateScript {
    $dateFieldXml = @"
<Field
    Type="DateTime"
    DisplayName="Submitted Date"
    StaticName="sfSubmittedDate"
    Name="sfSubmittedDate"
    Format="DateOnly"
    Group="$SiteColumnGroup"
    Required="FALSE">
    <Default></Default>
</Field>
"@
    Add-PnPFieldFromXml -FieldXml $dateFieldXml
}

# 2h. sfSessionType - Choice: Workshop, Session
New-PnPFieldIfMissing -InternalName "sfSessionType" -CreateScript {
    Add-PnPField -DisplayName "Session Type" -InternalName "sfSessionType" `
        -Type Choice -Choices "Workshop", "Session" -Group $SiteColumnGroup
}

# 2i. sfConference - Self-referencing lookup
# We have to create the list first (done above), then create the lookup
# pointing at it. Using XML so we can set ShowField and List ID precisely.
$listId = (Get-PnPList -Identity $ListTitle).Id.ToString("B").ToUpper()

New-PnPFieldIfMissing -InternalName "sfConference" -CreateScript {
    $lookupXml = @"
<Field
    Type="Lookup"
    DisplayName="Conference"
    StaticName="sfConference"
    Name="sfConference"
    List="$listId"
    ShowField="Title"
    Group="$SiteColumnGroup"
    Required="FALSE" />
"@
    Add-PnPFieldFromXml -FieldXml $lookupXml
}

# ---------------------------------------------------------------------------
# 3. Create content types
# ---------------------------------------------------------------------------
Write-Host "`n[3/6] Creating content types..." -ForegroundColor Yellow

# Parent content type IDs:
#   0x0102        = Event (gives us EventDate, EndDate, fAllDayEvent, fRecurrence)
#   0x01          = Item (use for non-event-shaped types if desired)
# We'll inherit from Event for all three so the calendar view honors the
# date range fields automatically.

function New-PnPContentTypeIfMissing {
    param(
        [string]$Name,
        [string]$ParentId,
        [string]$Description
    )
    $ct = Get-PnPContentType -Identity $Name -ErrorAction SilentlyContinue
    if ($ct) {
        Write-Host "  - Content type '$Name' already exists, skipping creation."
        return $ct
    }
    # Add by parent ID via XML for precision
    $parent = Get-PnPContentType | Where-Object { $_.StringId -eq $ParentId }
    if (-not $parent) {
        throw "Parent content type ID $ParentId not found at site level."
    }
    $ct = Add-PnPContentType -Name $Name -Description $Description `
        -Group $ContentTypeGroup -ParentContentType $parent
    Write-Host "  + Content type '$Name' created."
    return $ct
}

$ctConference = New-PnPContentTypeIfMissing `
    -Name "SF Conference" `
    -ParentId "0x0102" `
    -Description "A conference spanning one or more days. Parent of Call Milestones and Sessions."

$ctCallMilestone = New-PnPContentTypeIfMissing `
    -Name "SF Call Milestone" `
    -ParentId "0x0102" `
    -Description "Open or Close milestone for Call for Speakers or Call for Sponsors."

$ctSession = New-PnPContentTypeIfMissing `
    -Name "SF Session" `
    -ParentId "0x0102" `
    -Description "An individual conference session or workshop."

# ---------------------------------------------------------------------------
# 4. Add site columns to content types
# ---------------------------------------------------------------------------
Write-Host "`n[4/6] Adding site columns to content types..." -ForegroundColor Yellow

function Add-FieldToCT {
    param(
        [string]$ContentTypeName,
        [string]$FieldInternalName,
        [bool]$Required = $false
    )
    Add-PnPFieldToContentType -ContentType $ContentTypeName `
        -Field $FieldInternalName -Required:$Required -ErrorAction SilentlyContinue
    Write-Host "  $ContentTypeName <- $FieldInternalName"
}

# Conference: Website + Image (no self-lookup; it IS the parent)
Add-FieldToCT -ContentTypeName "SF Conference"     -FieldInternalName "sfWebsite"
Add-FieldToCT -ContentTypeName "SF Conference"     -FieldInternalName "sfConferenceImage"

# Call Milestone: Conference lookup, CallType, CallPhase, SubmissionStatus, SubmittedDate
Add-FieldToCT -ContentTypeName "SF Call Milestone" -FieldInternalName "sfConference"        -Required $true
Add-FieldToCT -ContentTypeName "SF Call Milestone" -FieldInternalName "sfCallType"          -Required $true
Add-FieldToCT -ContentTypeName "SF Call Milestone" -FieldInternalName "sfCallPhase"         -Required $true
Add-FieldToCT -ContentTypeName "SF Call Milestone" -FieldInternalName "sfSubmissionStatus"
Add-FieldToCT -ContentTypeName "SF Call Milestone" -FieldInternalName "sfSubmittedDate"

# Session: Conference lookup, SessionType
Add-FieldToCT -ContentTypeName "SF Session"        -FieldInternalName "sfConference"        -Required $true
Add-FieldToCT -ContentTypeName "SF Session"        -FieldInternalName "sfSessionType"       -Required $true

# ---------------------------------------------------------------------------
# 5. Attach content types to the list and remove default Item CT
# ---------------------------------------------------------------------------
Write-Host "`n[5/6] Attaching content types to list..." -ForegroundColor Yellow

Add-PnPContentTypeToList -List $ListTitle -ContentType "SF Conference"     -DefaultContentType
Add-PnPContentTypeToList -List $ListTitle -ContentType "SF Call Milestone"
Add-PnPContentTypeToList -List $ListTitle -ContentType "SF Session"

# Remove the default Item / Event content type from the list (cleanup)
$defaultsToRemove = @("Item", "Event")
foreach ($ctName in $defaultsToRemove) {
    try {
        Remove-PnPContentTypeFromList -List $ListTitle -ContentType $ctName -ErrorAction SilentlyContinue
        Write-Host "  Removed default '$ctName' from list."
    }
    catch {
        # Not present, fine
    }
}

# Index the lookup column for performance on a self-referencing list
Write-Host "  Adding index on sfConference..."
Set-PnPField -List $ListTitle -Identity "sfConference" -Values @{ Indexed = $true } -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# 6. Create views
# ---------------------------------------------------------------------------
Write-Host "`n[6/6] Creating views..." -ForegroundColor Yellow

# 6a. Calendar view
# PnP creates Calendar views via Add-PnPView -ViewType Calendar
$calendarView = Get-PnPView -List $ListTitle -Identity "Calendar" -ErrorAction SilentlyContinue
if (-not $calendarView) {
    Add-PnPView -List $ListTitle -Title "Calendar" -ViewType Calendar `
        -Fields "Title", "EventDate", "EndDate" -SetAsDefault | Out-Null
    Write-Host "  + Calendar view created."
}
else {
    Write-Host "  - Calendar view already exists."
}

# 6b. By Conference (grouped)
$byConference = Get-PnPView -List $ListTitle -Identity "By Conference" -ErrorAction SilentlyContinue
if (-not $byConference) {
    $byConfQuery = "<GroupBy Collapse=`"TRUE`" GroupLimit=`"30`"><FieldRef Name=`"sfConference`" /></GroupBy><OrderBy><FieldRef Name=`"EventDate`" /></OrderBy>"
    Add-PnPView -List $ListTitle -Title "By Conference" `
        -Fields "Title", "EventDate", "EndDate", "ContentType", "sfConference", "sfSessionType" `
        -Query $byConfQuery | Out-Null
    Write-Host "  + 'By Conference' view created."
}

# 6c. Submission Pipeline (Call Close events grouped by status)
$pipeline = Get-PnPView -List $ListTitle -Identity "Submission Pipeline" -ErrorAction SilentlyContinue
if (-not $pipeline) {
    $pipelineQuery = @"
<Where>
  <And>
    <Eq><FieldRef Name="ContentType" /><Value Type="Text">SF Call Milestone</Value></Eq>
    <Eq><FieldRef Name="sfCallPhase" /><Value Type="Text">Close</Value></Eq>
  </And>
</Where>
<GroupBy Collapse="FALSE"><FieldRef Name="sfSubmissionStatus" /></GroupBy>
<OrderBy><FieldRef Name="EventDate" /></OrderBy>
"@
    Add-PnPView -List $ListTitle -Title "Submission Pipeline" `
        -Fields "Title", "sfConference", "sfCallType", "EventDate", "sfSubmissionStatus", "sfSubmittedDate" `
        -Query $pipelineQuery | Out-Null
    Write-Host "  + 'Submission Pipeline' view created."
}

# 6d. Upcoming Sessions
$upcoming = Get-PnPView -List $ListTitle -Identity "Upcoming Sessions" -ErrorAction SilentlyContinue
if (-not $upcoming) {
    $upcomingQuery = @"
<Where>
  <And>
    <Eq><FieldRef Name="ContentType" /><Value Type="Text">SF Session</Value></Eq>
    <Geq><FieldRef Name="EventDate" /><Value Type="DateTime"><Today /></Value></Geq>
  </And>
</Where>
<OrderBy><FieldRef Name="EventDate" Ascending="TRUE" /></OrderBy>
"@
    Add-PnPView -List $ListTitle -Title "Upcoming Sessions" `
        -Fields "Title", "sfConference", "EventDate", "EndDate", "sfSessionType" `
        -Query $upcomingQuery | Out-Null
    Write-Host "  + 'Upcoming Sessions' view created."
}

Write-Host "`nDone." -ForegroundColor Green
Write-Host "Next: Apply the calendar view formatting JSON to the Calendar view." -ForegroundColor Cyan