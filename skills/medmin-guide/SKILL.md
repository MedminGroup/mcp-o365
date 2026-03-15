---
name: medmin-guide
description: Shows the user a friendly guide to all Medmin tools connected to Claude — what's available, how to use it, and example prompts. Use when the user asks what Claude can do, how the meeting analyser works, what tools are connected, or asks for help getting started.
---

# Medmin Guide

When this skill is invoked, present the following guide to the user exactly as formatted below. Do not summarise it — show it in full. Substitute the user's actual signed-in account name where indicated if you can determine it from `accounts_list`; otherwise leave the placeholders.

---

Present this to the user:

---

## What Claude can do for you

Claude is connected to your **Microsoft 365** account and can read your calendar, emails, files, and Teams meeting transcripts — all processed privately on your device.

---

### ✅ Teams Meeting Analyser

Fetches live transcripts from your recorded Teams meetings and produces a detailed communication analysis for any participant.

**Before you start — check you're signed in:**
> Ask Claude: *"am I signed in to Microsoft 365?"*

If not signed in yet, run these three steps:
1. Ask Claude: **`accounts_add`**
2. Open the link shown and sign in with your **medmin.co.uk** account
3. Ask Claude: **`accounts_complete`**

**How to start a meeting analysis:**

| What you want | What to ask Claude |
|---|---|
| Analyse one person's communication | *"Analyse last week's [meeting name] for [name]"* |
| Find out what was decided | *"What was decided in Thursday's [meeting name]?"* |
| Analyse everyone's contribution | *"Analyse everyone in this week's [meeting name]"* |
| Check someone's communication style | *"How did [name] communicate in yesterday's meeting?"* |

**What you get back:**
- Speaking ratios and word counts per participant
- Communication patterns: directness, conflict avoidance, active listening, facilitation
- Verbatim quotes with timestamps and coaching suggestions
- Key decisions and action items with owners
- Strengths and growth opportunities

**Requirements:**
- The meeting must have had **Start transcription** active during the call
  *(In Teams: click ••• → Start transcription before your meeting begins)*
- You must have **organised** the meeting — attendee-only access doesn't work for transcript retrieval

**Example prompts to try:**
```
"Analyse last Monday's weekly standup for Sarah"
"What action items came out of Thursday's board meeting?"
"How did James communicate in the product review — was he direct?"
"Analyse everyone's contribution to this week's all-hands"
```

**Where results are saved:**
Analysis files are saved to your Downloads folder as:
`meeting-analysis-[firstname]-YYYY-MM-DD.txt`

---

### 🔜 HubSpot Integration *(coming soon)*

Once you connect your HubSpot account, Claude will be able to:

- **Save meeting analysis directly to a contact record** — so every client interaction is logged automatically
- **Pull up a contact's history** before a meeting — who they are, what's been discussed, open deals
- **Update contact notes** from a conversation — just describe what happened and Claude will log it

When HubSpot is connected, prompts like these will work:
```
"Analyse Tuesday's meeting with John Smith and save it to his HubSpot record"
"What do we know about Sarah Jones before my call with her tomorrow?"
"Log a note on the Acme account — we agreed to push the demo to next week"
```

You'll be notified when this is ready to connect.

---

### 💡 Tips

- **Recurring meetings** — Claude can find any specific occurrence. Just say *"last Tuesday's"* or *"the 10th March"* version.
- **Multiple accounts** — if you have more than one Microsoft 365 account signed in, tell Claude which one: *"use my medmin.co.uk account"*
- **No transcript found?** — Check that transcription was started during the meeting. If it wasn't, there's no transcript to fetch.
- **Re-run this guide any time** — just ask: *"show me the Medmin guide"* or *"what can Claude do?"*

---
