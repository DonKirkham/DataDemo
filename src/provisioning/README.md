# ConferenceEvents List Provisioning

## Files

- **Provision-ConferenceEventsList.ps1** — PnP PowerShell script that creates the list, site columns, content types, field-to-CT mappings, and four views.
- **CalendarView-Formatting.json** — Calendar view formatter. Color-codes events by content type and flags overdue Planned submissions in red.
- **ColumnFormatting-SubmissionFields.json** — Hides SubmissionStatus and SubmittedDate fields on the form unless Call Phase = Close.

## Run order

```powershell
# 1. Connect (use your tenant)
Connect-PnPOnline -Url https://yourtenant.sharepoint.com/sites/yoursite -Interactive

# 2. Run the provisioning script
.\Provision-ConferenceEventsList.ps1

# 3. Apply calendar view formatting
#    Open the Calendar view in the browser > View dropdown > Format current view
#    > Advanced mode > paste contents of CalendarView-Formatting.json > Save

# 4. Apply column formatting to sfSubmissionStatus and sfSubmittedDate
#    List Settings > each column > Column formatting > paste ColumnFormatting-SubmissionFields.json
```

## Color legend (calendar view)

| Event | Color |
|---|---|
| Conference (multi-day all-day block) | Cornflower blue |
| Call for Speakers/Sponsors — Open | Mint green |
| Call for Speakers/Sponsors — Close | Gold |
| Call Close, Planned status, past due | Dust rose (red flag) |
| Session — Session | Lavender |
| Session — Workshop | Peach |

## Authoring sequence

1. Add the Conference item first. Title, EventDate (start), EndDate (end), Website, Image.
2. For each Call milestone, add a Call Milestone item with the Conference lookup, Call Type (Speakers/Sponsors), Call Phase (Open/Close), and date.
3. Once a CFS or CFSp is submitted, edit the Close milestone, set Submission Status = Submitted and the Submitted Date.
4. Add Session items as the agenda firms up, lookup'd to the parent Conference.

## Self-referencing lookup notes

- The Conference picker on Call Milestone and Session forms shows EVERY item in the list, not just Conferences. Discipline-only filtering for now.
- Cascade delete is not supported. Delete child events before deleting a Conference, or build a Power Automate cleanup flow.
- The lookup column is indexed in the script for list view threshold protection.

## Schema summary

| Internal Name | Type | Used By |
|---|---|---|
| Title | Built-in | All |
| EventDate | Built-in (Event CT) | All |
| EndDate | Built-in (Event CT) | Conference, Session |
| fAllDayEvent | Built-in (Event CT) | Conference (Yes), Calls (Yes), Session (No) |
| sfConference | Lookup → self (Title) | Call Milestone, Session |
| sfWebsite | URL | Conference |
| sfConferenceImage | Thumbnail (Image) | Conference |
| sfCallType | Choice (Speakers, Sponsors) | Call Milestone |
| sfCallPhase | Choice (Open, Close) | Call Milestone |
| sfSubmissionStatus | Choice (Planned, Submitted) | Call Milestone (Close only) |
| sfSubmittedDate | Date only | Call Milestone (Close only) |
| sfSessionType | Choice (Workshop, Session) | Session |
