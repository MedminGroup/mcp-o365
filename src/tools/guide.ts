import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';

export function registerGuideTools(server: McpServer): void {
  server.tool(
    'medmin_guide',
    [
      'Show a formatted guide to all Medmin tools connected to Claude.',
      'Call this when the user asks what Claude can do, how to use the meeting analyser,',
      'what tools are available, how to sign in, or needs help getting started.',
    ].join(' '),
    {},
    async () => {
      const guide = `
## What Claude can do for you

Claude is connected to your **Microsoft 365** account and can read your calendar, emails, files, and Teams meeting transcripts — all processed privately on your device.

---

### ✅ Teams Meeting Analyser

Fetches live transcripts from your recorded Teams meetings and produces a detailed communication analysis for any participant.

**Before you start — check you're signed in:**
Ask Claude: *"am I signed in to Microsoft 365?"*

If not signed in yet, follow these steps:
1. Ask Claude to run: **accounts_add**
2. Open the link shown and sign in with your **medmin.co.uk** account
3. Ask Claude to run: **accounts_complete**

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
\`\`\`
"Analyse last Monday's weekly standup for Sarah"
"What action items came out of Thursday's board meeting?"
"How did James communicate in the product review — was he direct?"
"Analyse everyone's contribution to this week's all-hands"
\`\`\`

Analysis files are saved to your Downloads folder as: \`meeting-analysis-[firstname]-YYYY-MM-DD.txt\`

---

### 🔜 HubSpot Integration *(coming soon)*

Once connected, Claude will be able to:
- **Save meeting analysis directly to a contact record** — every client interaction logged automatically
- **Pull up a contact's history** before a meeting — who they are, what's been discussed, open deals
- **Update contact notes** — just describe what happened and Claude will log it

\`\`\`
"Analyse Tuesday's meeting with John Smith and save it to his HubSpot record"
"What do we know about Sarah Jones before my call with her tomorrow?"
"Log a note on the Acme account — we agreed to push the demo to next week"
\`\`\`

---

### 💡 Tips

- **Recurring meetings** — just say *"last Tuesday's"* or *"the 10th March"* version
- **Multiple accounts** — tell Claude which one: *"use my medmin.co.uk account"*
- **No transcript found?** — transcription must have been started during the meeting
- **Show this guide again** — just ask: *"show me the Medmin guide"* or *"what can Claude do?"*
`.trim();

      return {
        content: [{ type: 'text' as const, text: guide }],
      };
    },
  );
}
