# Campaign Auditor v1

**Automated HubSpot marketing campaign analysis delivered to Slack.**

Campaign Auditor pulls live data from your HubSpot portal, sends it to Claude (Anthropic) for strategic analysis, and delivers a formatted summary plus a downloadable HTML report directly to your Slack channel — all triggered by a single `/audit` slash command.

---

## What It Does

When you type `/audit` in Slack, the system:

1. Pulls contacts, email activity, and deal data from HubSpot (Portal 243444428)
2. Assembles the data into a structured analysis payload
3. Sends it to Claude for AI-powered marketing analysis
4. Generates a professional Slack summary with scores, benchmarks, and recommendations
5. Creates a styled HTML report with detailed tables, charts, and strategic guidance
6. Uploads the report file to your Slack channel for download

**Total execution time:** ~35-40 seconds

---

## Architecture

```
Slack /audit Command
      │
      ▼
┌─────────────────────────┐
│  Webhook: Receive Audit │ ◄── Slack slash command POST
│  Request                │
└─────────┬───────────────┘
          │
          ▼
┌─────────────────────────┐
│  Parse CSV Data         │ ◄── Extracts any attached CSV data
└─────────┬───────────────┘
          │
    ┌─────┴─────┐
    ▼           ▼
┌────────┐ ┌────────┐
│HubSpot:│ │HubSpot:│
│  Get   │ │  Get   │
│Contacts│ │ Deals  │
└───┬────┘ └───┬────┘
    └─────┬────┘
          ▼
┌─────────────────────────┐
│  Assemble Analysis      │ ◄── Combines contacts + deals + CSV
│  Payload                │     into structured JSON
└─────────┬───────────────┘
          ▼
┌─────────────────────────┐
│  Build Claude Request   │ ◄── Formats the API request with
│                         │     analysis prompt + payload
└─────────┬───────────────┘
          ▼
┌─────────────────────────┐
│  Claude: Analyze        │ ◄── POST to api.anthropic.com
│  Campaign               │     Returns structured JSON analysis
└─────────┬───────────────┘
          ▼
┌─────────────────────────┐
│  Format Slack + Report  │ ◄── Parses Claude response, builds
│  Data                   │     mrkdwn message + report data
└─────────┬───────────────┘
          │
    ┌─────┴──────────────┐
    ▼                    ▼
┌────────────┐   ┌──────────────┐
│  Send      │   │  Generate    │
│  Audit     │   │  Report      │ ◄── HTML report (Code node)
│  Summary   │   │              │
│  (Slack)   │   └──────┬───────┘
└────────────┘          ▼
                 ┌──────────────┐
                 │  Get Upload  │ ◄── Slack files.getUploadURLExternal
                 │  URL         │
                 └──────┬───────┘
                        ▼
                 ┌──────────────┐
                 │  Prepare     │ ◄── Merges binary + upload URL
                 │  Upload      │
                 └──────┬───────┘
                        ▼
                 ┌──────────────┐
                 │  Upload File │ ◄── POST to pre-signed URL
                 └──────┬───────┘
                        ▼
                 ┌──────────────┐
                 │  Complete    │ ◄── files.completeUploadExternal
                 │  Upload      │     Posts file to #marketing
                 └──────────────┘
```

---

## What Gets Delivered

### Slack Summary Message

A formatted mrkdwn message posted to #marketing including:

- **Overall Score** (A+ through F) with emoji indicator
- **Executive Summary** — one-paragraph assessment
- **Benchmark Comparison** — open rate and click rate vs industry averages
- **Email Performance** — top and bottom performing emails with ratings
- **Audience Health** — contact count, active rate, disengaged rate, risk flags
- **Top 3 Recommendations** — prioritized actions with expected impact

### HTML Report File

A downloadable file (typically 14-15 KB) that opens in any browser, containing:

- **Title page** with portal ID and date
- **Score card** with color-coded overall grade and executive summary
- **Benchmark comparison table** with yours vs industry, difference, and verdict
- **Email performance table** with per-email scores, open/click rates, and insights
- **Key patterns** identified across email campaigns
- **Audience health dashboard** with stat cards (total contacts, active %, disengaged %, opted out %)
- **Findings and risks** sections
- **Lifecycle funnel** with distribution table and visual bar chart
- **Strategic recommendations** with priority levels, effort estimates, and expected impact
- **Print-ready** — uses CSS print styles for clean printing

---

## Infrastructure

| Component | Service | Details |
|-----------|---------|---------|
| Workflow Engine | n8n | Self-hosted on Render (Web Service) |
| Hosting | Render | Hobby tier, n8n-jobbot instance |
| AI Analysis | Anthropic Claude | API via HTTP Request node |
| CRM Data | HubSpot | Portal 243444428, API + CSV hybrid |
| Delivery | Slack | Bot in #marketing channel |
| Report Format | HTML | Zero external dependencies |

### Key Environment Variables (Render)

| Variable | Purpose |
|----------|---------|
| `ANTHROPIC_API_KEY` | Claude API authentication |
| `NODES_EXCLUDE` | Set to `[]` to enable Execute Command node |
| `N8N_RUNNERS_ENABLED` | Task runner setting |
| `N8N_HOST` | n8n instance hostname |
| `N8N_PORT` | n8n instance port |

### Slack Bot Configuration

- **App Name:** Campaign Generator
- **Required Scopes:** `chat:write`, `files:write`, `files:read`
- **Channel:** #marketing (Channel ID: C0AEH1BS9DH)
- **Slash Command:** `/audit` pointed at the n8n webhook URL
- **Bot must be invited** to #marketing channel (`/invite @Campaign Generator`)

### Slack File Upload Flow (3-Step API)

Slack deprecated `files.upload`. The current flow uses three API calls:

1. `files.getUploadURLExternal` — gets a pre-signed upload URL + file_id
2. POST to the pre-signed URL — uploads the binary file (no auth needed)
3. `files.completeUploadExternal` — finalizes and posts to channel

---

## n8n Node Reference

| # | Node Name | Type | Purpose |
|---|-----------|------|---------|
| 1 | Webhook: Receive Audit Request | Webhook | Entry point from /audit command |
| 2 | Parse CSV Data | Code | Extracts CSV data from webhook payload |
| 3 | HubSpot: Get Contacts | HTTP Request | Pulls contacts from HubSpot API |
| 4 | HubSpot: Get Deals | HTTP Request | Pulls deals from HubSpot API |
| 5 | Assemble Analysis Payload | Code | Combines all data sources |
| 6 | Build Claude Request | Code | Formats Anthropic API request |
| 7 | Claude: Analyze Campaign | HTTP Request | POST to api.anthropic.com |
| 8 | Format Slack + Report Data | Code | Parses response, builds messages |
| 9 | Generate Report | Code | Builds HTML report (zero dependencies) |
| 10 | Send Audit Summary | Slack | Posts mrkdwn summary to #marketing |
| 11 | Get Upload URL | HTTP Request | Slack files.getUploadURLExternal |
| 12 | Prepare Upload | Code | Merges binary data with upload URL |
| 13 | Upload File | HTTP Request | POSTs file to pre-signed URL |
| 14 | Complete Upload | HTTP Request | Finalizes file in Slack channel |

---

## Claude Analysis Schema

The Claude API returns a structured JSON object with these sections:

```
{
  "overallScore": "B-",
  "executiveSummary": "...",
  "benchmarkComparison": {
    "openRate": { "yours": 23.5, "industry": 21.5, "verdict": "above" },
    "clickRate": { "yours": 4.4, "industry": 2.3, "verdict": "above" },
    "bounceRate": { ... },
    "unsubRate": { ... }
  },
  "emailAnalysis": {
    "overallRating": "A-",
    "topPerformer": "Product Update",
    "bottomPerformer": "Monthly Newsletter - Feb",
    "emails": [
      { "name": "...", "score": "A", "openRate": 23.5, "clickRate": 4.4, "insight": "..." }
    ],
    "patterns": ["..."]
  },
  "audienceHealth": {
    "score": "D+",
    "totalContacts": 33,
    "activeRate": 0,
    "disengagedRate": 0,
    "optOutRate": 0,
    "findings": ["..."],
    "risks": ["..."]
  },
  "funnelAnalysis": {
    "score": "C",
    "distribution": { "subscriber": 10, "lead": 15, "mql": 4, "customer": 4 },
    "gaps": ["..."]
  },
  "recommendations": [
    {
      "title": "...",
      "description": "...",
      "priority": "high",
      "effort": "quick_win",
      "expectedImpact": "..."
    }
  ]
}
```

---

## JSON Parsing

Claude's response sometimes contains unescaped newlines or markdown code fences. The Format node uses a character-by-character state machine parser that:

1. Strips ` ```json ` and ` ``` ` markers
2. Walks character by character tracking string context
3. Escapes literal newlines, carriage returns, and tabs inside strings
4. Produces valid JSON for `JSON.parse()`

This is more resilient than simple regex replacement and handles edge cases in Claude's output formatting.

---

## Troubleshooting

| Problem | Cause | Fix |
|---------|-------|-----|
| `/audit` returns nothing | Webhook URL changed after redeploy | Update slash command URL in Slack app settings |
| Claude response truncation | Default max_tokens too low | Set `max_tokens: 4096` in Build Claude Request |
| Slack message not formatting | Using Block Kit instead of mrkdwn | Use Simple Text Message type with mrkdwn |
| `not_in_channel` error | Bot not in #marketing | Run `/invite @Campaign Generator` in channel |
| `method_deprecated` on upload | Using old files.upload API | Use 3-step upload flow (getURL → upload → complete) |
| Binary data missing in upload | HTTP Request nodes don't pass binary | Use Prepare Upload code node to carry binary forward |
| HTML shows as code in Slack | Normal Slack behavior for HTML files | Expected — users download and open in browser |
| `Module 'fs' is disallowed` | n8n Code node sandbox restriction | Don't use require('fs') — use Buffer globals instead |

---

## Development History

### Session 1: Foundation
- Built hybrid data architecture (HubSpot API + CSV upload)
- Implemented Claude integration with structured analysis prompt
- Created JSON parsing state machine for Claude's response format
- Resolved Claude truncation with explicit max_tokens setting

### Session 2: Slack Delivery
- Attempted Block Kit formatting — abandoned due to rendering issues
- Pivoted to simple mrkdwn text — immediate success
- Configured native Slack node with proper message formatting
- Added `/audit` slash command trigger

### Session 3: Report Generation
- Attempted DOCX via Code node with `require('docx')` — blocked by sandbox
- Attempted Execute Command with standalone script — node unavailable
- Enabled Execute Command via `NODES_EXCLUDE=[]`
- Attempted file system operations — `require('fs')` blocked by sandbox
- Attempted base64 piping through shell — data corruption
- **Solution: Pure HTML report in single Code node — zero dependencies**
- Implemented 3-step Slack file upload (new API)
- Full end-to-end flow operational

---

## Key Technical Decisions

**Why HTML instead of DOCX?**
The n8n Code node sandbox blocks `require('fs')` and `require('child_process')`. The docx npm package was allowed via `NODE_FUNCTION_ALLOW_EXTERNAL=docx` but the task runner process couldn't find it due to pnpm workspace isolation. HTML generation uses only string concatenation and `Buffer.from()` — both are Node.js globals that work in any sandbox. Zero dependencies, zero failure points.

**Why mrkdwn instead of Block Kit?**
Slack Block Kit rendered inconsistently in n8n's native Slack node. Simple mrkdwn text with manual formatting produced reliable, readable results with no rendering surprises.

**Why state machine JSON parsing?**
Claude occasionally wraps responses in markdown code fences and includes unescaped newlines in string values. A character-by-character parser handles these edge cases where `JSON.parse()` alone would fail.

**Why 3-step file upload?**
Slack deprecated `files.upload` in favor of `getUploadURLExternal` → upload → `completeUploadExternal`. The new API separates file storage from channel posting, which actually gives more control.
