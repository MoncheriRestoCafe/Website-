# The Lunch Club — AI Sales Agent

An AI assistant, built on Claude, that works through `LunchClub_Master_Sales_Funnel.xlsx` with you:

- **Finds who to contact.** It filters prospects by priority (A/B/C), zone (Indiranagar / EGL / Bagmane) and funnel stage.
- **Researches each company** on the web: Bengaluru headcount, office moves, hiring, employee perks.
- **Drafts personalised outreach** (email, WhatsApp or LinkedIn), following the "Outreach Approach" notes in the sheet.
- **Keeps the funnel up to date.** It sets the funnel stage, sets follow-up dates, adds dated notes and logs each contact in the *Weekly Outreach Tracker*.
- **Lists the follow-ups that are due.**

It **never sends anything**. Drafts are saved in `sales_agent/drafts/` for you to review, edit and send yourself.

## Setup (one time)

```bash
pip install -r sales_agent/requirements.txt
export ANTHROPIC_API_KEY=sk-ant-...      # from console.anthropic.com
```

## Use it

Interactive chat:

```bash
python sales_agent/agent.py
```

One-off task:

```bash
python sales_agent/agent.py "Draft emails for the top 5 Tier A prospects in Indiranagar"
```

Example requests:

- "Draft emails for all Tier A coworking spaces in EGL."
- "Research Goldman Sachs in EGL and write a WhatsApp message for their facilities team."
- "I just emailed BHIVE CMH Road. Mark it contacted, follow up next Thursday."
- "Which follow-ups are due this week?"
- "Give me today's 5 calls with a one-line opener for each."

Close the spreadsheet in Excel before the agent updates it, or your changes and the agent's can overwrite each other.

## How it works

`agent.py` uses the Anthropic Python SDK's tool runner with Claude Opus 5.5. Claude has six local tools, which read and write the spreadsheet and save drafts, plus Anthropic's hosted web search for research. Edit `SYSTEM_PROMPT` in `agent.py` to change the tone, the pitch, the pricing it quotes or the sign-off.
