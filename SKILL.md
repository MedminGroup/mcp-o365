---
name: teams-meeting-analyser
description: Fetches Microsoft Teams meeting transcripts via the mcp-o365 MCP server and analyses communication patterns, speaking ratios, filler words, conflict avoidance, facilitation style, and key decisions for any named participant. Use when asked to analyse a Teams meeting, a meeting transcript, or a person's communication style in a recorded meeting.
---

# Teams Meeting Analyser

Fetches live transcripts from Microsoft Teams via the Microsoft Graph API and produces
deep communication pattern analysis for one or more named participants.

## When to Use This Skill

- "Analyse [person]'s contribution to [meeting name]"
- "What was decided in last Thursday's [meeting]?"
- "How did [person] communicate in the [meeting] meeting?"
- "Pull the transcript from [meeting name] and analyse it"
- "Analyse this week's / last week's [recurring meeting name]"

## Prerequisites

- Microsoft 365 account authenticated via the mcp-o365 MCP server
  (run `accounts_add` in Claude if not yet signed in)
- Azure app registration must include these delegated permissions (admin-consented):
    - OnlineMeetings.Read
    - OnlineMeetingTranscript.Read.All
- Meeting must have had transcription enabled (Teams: ... → Start transcription)

## What This Skill Does

1. Resolves the meeting from the user's calendar using `calendar_list_events`
2. Retrieves the Teams online meeting object via `meetings_get_by_join_url`
3. Lists available transcripts via `meetings_list_transcripts`
4. Downloads the VTT transcript via `meetings_get_transcript`
5. Analyses communication patterns for named participants
6. Produces a structured Meeting Insights Summary

## Step-by-Step Workflow

### 0. Confirm signed-in account

Call `accounts_list` to get the user's account identifier. Use that value as the
`account` parameter in all subsequent calls.

### 1. Find the calendar event

Use `calendar_list_events` to locate the meeting. Ask the user for the date if not
provided. Search a ±1 day window to allow for timezone differences.

```
calendar_list_events(
  account="<from accounts_list>",
  start="<date>T00:00:00",
  end="<date>T23:59:59"
)
```

Note the event's `joinWebUrl` from the response. If the event has no `joinWebUrl`
it was not a Teams meeting and has no transcript.

### 2. Get the online meeting ID

```
meetings_get_by_join_url(
  join_url="<joinWebUrl from step 1>",
  account="<account>"
)
```

Note the `id` field from the response.

### 3. List transcripts

```
meetings_list_transcripts(
  meeting_id="<id from step 2>",
  account="<account>"
)
```

For recurring meetings, multiple transcripts will be listed. Match by
`createdDateTime` to find the correct occurrence. Note the target transcript's `id`.

### 4. Download the VTT transcript

```
meetings_get_transcript(
  meeting_id="<id from step 2>",
  transcript_id="<id from step 3>",
  account="<account>"
)
```

The response is raw VTT text with speaker labels (`<v Speaker Name>`) and timestamps.

### 5. Validate the transcript

Before analysing, extract basic stats from the VTT:

- All unique speaker names (from `<v ...>` tags)
- Turn count per speaker
- Approximate duration (last timestamp minus first)

Confirm with the user who to focus on if not already specified.

### 6. Analyse the transcript

For each participant to be analysed, produce a Meeting Insights Summary covering:

**Speaking Statistics**
- Turn count and percentage per speaker
- Word count and percentage per speaker
- Average words per turn
- Filler word counts: "um", "uh", "like", "you know", "I think", "sort of", "kind of"
- Question vs statement ratio

**Communication Patterns**
- Directness — are statements assertive or hedged?
- Conflict avoidance — hedging language, subject changes, indirect phrasing
- Active listening — building on others' points, clarifying questions, paraphrasing
- Leadership/facilitation — agenda control, drawing out quieter voices, handling disagreement

**Specific Examples**
For each pattern, include:
- Timestamp
- Verbatim quote
- Why it matters
- Better approach (for growth areas)

**Key Decisions and Action Items**
Extract all decisions and action items with owners.

**Strengths and Growth Opportunities**
Minimum 3 strengths, minimum 4 growth opportunities, all with timestamps.

## Output Format

```markdown
# Meeting Insights Summary — [Name]
**Meeting:** [Name] | [Date] | [Duration]
**Participants:** [List]

## Speaking Ratios
[Table]

## Communication Patterns
[Sections per pattern with quotes and timestamps]

## Key Decisions and Action Items
[Table with owner]

## Strengths
[Numbered list with evidence]

## Growth Opportunities
[Numbered list with timestamps and better alternatives]

## Summary
[2–3 sentence overall assessment]
```

## Saving Outputs

Save analysis files to ~/Downloads/ with the naming convention:
  meeting-analysis-[firstname]-YYYY-MM-DD.txt

## Known Gotchas

- Recurring meetings share a single online meeting ID. Use `meetings_list_transcripts`
  and filter by `createdDateTime` to find the right occurrence.
- Teams only generates a transcript if "Start transcription" was active during the
  meeting. If the transcripts list is empty, no transcript was recorded.
- `meetings_get_by_join_url` only works for meetings organised by the signed-in account.
  If the user was an attendee (not the organiser), the lookup will return no results —
  ask them to try with the organiser's account.
- If `meetings_get_transcript` returns a 403, the Azure app registration is missing
  `OnlineMeetingTranscript.Read.All` admin consent — contact your administrator.

## Example Prompts

"Analyse last week's Weekly Medmin meeting for Sarah"
"Pull the transcript from the board show and tell and tell me what was decided"
"How did James communicate in Thursday's team meeting?"
"Analyse everyone's contribution to this week's Monday standup"
