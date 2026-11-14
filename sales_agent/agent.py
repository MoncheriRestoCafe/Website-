"""The Lunch Club — AI sales agent.

Works through LunchClub_Master_Sales_Funnel.xlsx: picks prospects, researches
them on the web, drafts personalised outreach (email / WhatsApp / LinkedIn),
and keeps the funnel and the weekly outreach tracker up to date.

The agent never sends anything. Drafts are written to sales_agent/drafts/
for a person to review and send.

Usage:
    python sales_agent/agent.py                      # interactive chat
    python sales_agent/agent.py "Draft emails for Tier A Indiranagar prospects"
"""

from __future__ import annotations

import datetime as dt
import re
import sys
from pathlib import Path

import anthropic
import openpyxl
from anthropic import beta_tool

ROOT = Path(__file__).resolve().parent.parent
FUNNEL_XLSX = ROOT / "LunchClub_Master_Sales_Funnel.xlsx"
DRAFTS_DIR = Path(__file__).resolve().parent / "drafts"

FUNNEL_SHEET = "Master Sales Funnel"
TRACKER_SHEET = "Weekly Outreach Tracker"
HEADER_ROW = 4  # both sheets have a title block above the header row

# Matches the Funnel Stage dropdown in the spreadsheet.
FUNNEL_STAGES = ["Cold", "Contacted", "Meeting Booked", "Proposal Sent", "Won", "Not Interested"]

MODEL = "claude-opus-5-5"

SYSTEM_PROMPT = """\
You are the sales assistant for The Lunch Club, a curated lunch membership \
for professionals in Indiranagar, Bengaluru, founded by Sanatha. You help \
Sanatha sell corporate and coworking lunch partnerships to the companies in \
the master sales funnel.

## What The Lunch Club offers
- Fresh lunches delivered to the desk every weekday, noon delivery windows.
- Three menu families: Power Bowls (high-protein, calorie-counted), \
Traditional Meals (South Indian, slow-cooked), and a rotating Daily Special.
- Compostable, temperature-stable packaging.
- Membership with a prepaid meal wallet; unused balance rolls over.
- Tiers: Starter ₹3,500/month (20 meals), Club ₹6,000/month (22 meals, \
priority 11:45am window, guest meal), Executive ₹10,000/month (unlimited, \
citywide, dedicated account manager).
- Founding-member pricing for companies that join before public launch.
- Contact: hello@thelunchclub.in, +91 98765 43210.

## How to work
- Start from the funnel: use list_prospects / get_prospect to find who to \
work on. Prioritise Tier A, then B, then C, unless told otherwise.
- Before drafting, research the company with web search when it would make \
the message more specific (headcount in Bengaluru, recent office moves, \
hiring, wellness or employee-benefit programs, coworking amenities). Keep \
research brief: one or two searches per company.
- Never invent facts about a company, a person's name, or a phone number. \
If you do not know a contact's name, address them by role.
- Draft with save_draft. Emails: subject line plus 90–150 words, warm and \
specific, one clear ask (a 15-minute call or a free tasting lunch for their \
team). WhatsApp: under 60 words. LinkedIn: under 300 characters. Sign off \
as Sanatha, Founder, The Lunch Club.
- Follow each prospect's "Outreach Approach" column; it holds Sanatha's own \
notes on the angle to take.
- After drafting, call update_prospect to set the stage to "Contacted" only \
if the user says the message was sent; otherwise leave the stage and add a \
note that a draft is ready. Log sent messages with log_outreach.
- You cannot send messages. Never claim that you did.
- Finish with a short summary: which prospects you worked on, where the \
drafts are, and any follow-ups due.
"""


# --------------------------------------------------------------------------- #
# Spreadsheet helpers
# --------------------------------------------------------------------------- #

def _load():
    return openpyxl.load_workbook(FUNNEL_XLSX)


def _headers(ws) -> list[str]:
    return [str(c.value).strip() if c.value is not None else "" for c in ws[HEADER_ROW]]


def _funnel_rows(ws):
    """Yield (excel_row_number, dict) for every prospect in the funnel."""
    headers = _headers(ws)
    for r in range(HEADER_ROW + 1, ws.max_row + 1):
        values = [ws.cell(row=r, column=c + 1).value for c in range(len(headers))]
        if values[0] is None and values[1] is None:
            continue
        yield r, dict(zip(headers, values))


def _find_prospect(ws, prospect_id: int):
    for r, row in _funnel_rows(ws):
        if row.get("#") == prospect_id:
            return r, row
    return None, None


def _fmt(row: dict) -> str:
    return "\n".join(f"{k}: {v}" for k, v in row.items() if v not in (None, ""))


def _parse_date(value: str) -> dt.date:
    return dt.date.fromisoformat(value.strip())


# --------------------------------------------------------------------------- #
# Tools
# --------------------------------------------------------------------------- #

@beta_tool
def list_prospects(priority: str = "", zone: str = "", stage: str = "", limit: int = 25) -> str:
    """List prospects from the master sales funnel, one line each.

    Args:
        priority: Filter by priority tier: "A", "B" or "C". Empty for all.
        zone: Filter by zone, e.g. "Indiranagar", "EGL", "Bagmane". Empty for all.
        stage: Filter by funnel stage, e.g. "Cold", "Contacted". Empty for all.
        limit: Maximum number of prospects to return.
    """
    ws = _load()[FUNNEL_SHEET]
    out = []
    for _, row in _funnel_rows(ws):
        if priority and str(row.get("Priority", "")).upper() != priority.upper():
            continue
        if zone and zone.lower() not in str(row.get("Zone", "")).lower():
            continue
        if stage and str(row.get("Funnel Stage", "")).lower() != stage.lower():
            continue
        out.append(
            f"#{row['#']} | {row['Company']} | {row.get('Zone')} | {row.get('Type')} | "
            f"Priority {row.get('Priority')} | {row.get('Est. Lunch Prospects')} | "
            f"Stage: {row.get('Funnel Stage')} | Follow-up: {row.get('Follow-up Date') or '-'}"
        )
        if len(out) >= limit:
            break
    return "\n".join(out) if out else "No prospects match those filters."


@beta_tool
def get_prospect(prospect_id: int) -> str:
    """Get every field for one prospect: contact role, phone, email, outreach approach, notes.

    Args:
        prospect_id: The prospect's number in the "#" column of the funnel.
    """
    ws = _load()[FUNNEL_SHEET]
    _, row = _find_prospect(ws, prospect_id)
    return _fmt(row) if row else f"Error: no prospect with # {prospect_id}."


@beta_tool
def list_due_followups(as_of: str = "") -> str:
    """List prospects whose follow-up date is on or before a given date.

    Args:
        as_of: ISO date (YYYY-MM-DD). Empty means today.
    """
    cutoff = _parse_date(as_of) if as_of else dt.date.today()
    ws = _load()[FUNNEL_SHEET]
    due = []
    for _, row in _funnel_rows(ws):
        fu = row.get("Follow-up Date")
        if fu is None or fu == "":
            continue
        if isinstance(fu, dt.datetime):
            fu = fu.date()
        elif isinstance(fu, str):
            try:
                fu = _parse_date(fu)
            except ValueError:
                continue
        if fu <= cutoff:
            due.append(f"#{row['#']} | {row['Company']} | Stage: {row.get('Funnel Stage')} | Due: {fu}")
    return "\n".join(due) if due else f"No follow-ups due on or before {cutoff}."


@beta_tool
def save_draft(prospect_id: int, channel: str, body: str, subject: str = "") -> str:
    """Save an outreach draft for a person to review and send. Does not send anything.

    Args:
        prospect_id: The prospect's number in the "#" column.
        channel: One of "email", "whatsapp", "linkedin".
        body: The full message text.
        subject: Email subject line (email only).
    """
    channel = channel.lower().strip()
    if channel not in {"email", "whatsapp", "linkedin"}:
        return "Error: channel must be email, whatsapp or linkedin."
    ws = _load()[FUNNEL_SHEET]
    _, row = _find_prospect(ws, prospect_id)
    if not row:
        return f"Error: no prospect with # {prospect_id}."

    DRAFTS_DIR.mkdir(exist_ok=True)
    slug = re.sub(r"[^a-z0-9]+", "-", str(row["Company"]).lower()).strip("-")[:50]
    path = DRAFTS_DIR / f"{prospect_id:03d}-{slug}-{channel}.md"
    to = row.get("Email") if channel == "email" else row.get("Phone") if channel == "whatsapp" else row.get("LinkedIn Search Query")
    lines = [
        f"# {row['Company']} — {channel}",
        "",
        f"- To: {to or 'unknown'} ({row.get('Contact Role') or 'contact'})",
        f"- Priority: {row.get('Priority')} · Zone: {row.get('Zone')}",
        f"- Drafted: {dt.date.today().isoformat()}",
        "",
    ]
    if subject:
        lines += [f"**Subject:** {subject}", ""]
    lines.append(body.strip())
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return f"Draft saved to {path.relative_to(ROOT)}"


@beta_tool
def update_prospect(prospect_id: int, funnel_stage: str = "", follow_up_date: str = "", note: str = "") -> str:
    """Update a prospect's funnel stage, follow-up date, and/or append a dated note.

    Args:
        prospect_id: The prospect's number in the "#" column.
        funnel_stage: New stage: Cold, Contacted, Meeting Booked, Proposal Sent, Won or Not Interested. Empty to leave unchanged.
        follow_up_date: ISO date (YYYY-MM-DD) for the next follow-up. Empty to leave unchanged.
        note: Text appended to the Notes column with today's date. Empty for none.
    """
    if funnel_stage and funnel_stage not in FUNNEL_STAGES:
        return f"Error: funnel_stage must be one of {', '.join(FUNNEL_STAGES)}."
    if follow_up_date:
        try:
            _parse_date(follow_up_date)
        except ValueError:
            return "Error: follow_up_date must be YYYY-MM-DD."

    wb = _load()
    ws = wb[FUNNEL_SHEET]
    r, row = _find_prospect(ws, prospect_id)
    if not row:
        return f"Error: no prospect with # {prospect_id}."
    col = {h: i + 1 for i, h in enumerate(_headers(ws))}

    changes = []
    if funnel_stage:
        ws.cell(row=r, column=col["Funnel Stage"], value=funnel_stage)
        changes.append(f"stage → {funnel_stage}")
    if follow_up_date:
        ws.cell(row=r, column=col["Follow-up Date"], value=follow_up_date)
        changes.append(f"follow-up → {follow_up_date}")
    if note:
        old = ws.cell(row=r, column=col["Notes"]).value or ""
        stamped = f"[{dt.date.today().isoformat()}] {note.strip()}"
        ws.cell(row=r, column=col["Notes"], value=f"{old}\n{stamped}".strip())
        changes.append("note added")
    if not changes:
        return "Nothing to update."
    wb.save(FUNNEL_XLSX)
    return f"Updated #{prospect_id} {row['Company']}: " + ", ".join(changes)


@beta_tool
def log_outreach(prospect_id: int, channel: str, notes: str, contact_name: str = "",
                 follow_up_date: str = "", outcome: str = "") -> str:
    """Add a row to the Weekly Outreach Tracker for an outreach attempt that was actually made.

    Args:
        prospect_id: The prospect's number in the "#" column.
        channel: How they were contacted, e.g. "Email", "WhatsApp", "Call", "LinkedIn", "Visit".
        notes: What was said or the response received.
        contact_name: Name of the person reached, if known.
        follow_up_date: ISO date (YYYY-MM-DD) of the next follow-up, if any.
        outcome: Short outcome, e.g. "No reply", "Interested", "Meeting booked".
    """
    wb = _load()
    _, row = _find_prospect(wb[FUNNEL_SHEET], prospect_id)
    if not row:
        return f"Error: no prospect with # {prospect_id}."
    ws = wb[TRACKER_SHEET]
    target = HEADER_ROW + 1
    while ws.cell(row=target, column=1).value or ws.cell(row=target, column=2).value:
        target += 1
    values = [dt.date.today().isoformat(), row["Company"], row.get("Zone"), contact_name,
              channel, notes, follow_up_date, outcome, row.get("Priority")]
    for c, v in enumerate(values, start=1):
        ws.cell(row=target, column=c, value=v)
    wb.save(FUNNEL_XLSX)
    return f"Logged outreach to {row['Company']} in row {target} of the tracker."


TOOLS = [
    list_prospects,
    get_prospect,
    list_due_followups,
    save_draft,
    update_prospect,
    log_outreach,
    # Anthropic-hosted web search, for researching prospects. Runs server-side.
    {"type": "web_search_20260209", "name": "web_search", "max_uses": 10,
     "user_location": {"type": "approximate", "city": "Bengaluru", "country": "IN",
                       "timezone": "Asia/Kolkata"}},
]


# --------------------------------------------------------------------------- #
# Agent loop
# --------------------------------------------------------------------------- #

def run_turn(client: anthropic.Anthropic, messages: list) -> list:
    """Run the agent until it finishes its turn; return the updated history."""
    runner = client.beta.messages.tool_runner(
        model=MODEL,
        max_tokens=16000,
        system=SYSTEM_PROMPT,
        tools=TOOLS,
        messages=messages,
        output_config={"effort": "medium"},
        # If a request is declined by a safety classifier, retry on a fallback model.
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",
        max_iterations=40,
    )
    history = list(messages)
    final = None
    for message in runner:
        final = message
        for block in message.content:
            if block.type == "tool_use":
                print(f"  · {block.name}({_short(block.input)})", flush=True)
            elif block.type == "server_tool_use":
                print(f"  · web_search({_short(block.input)})", flush=True)
        # Keep the whole exchange (assistant turns + tool results) so the next
        # question in interactive mode has full context. Content is passed back
        # unchanged, thinking blocks included.
        history.append({"role": "assistant", "content": message.content})
        tool_results = runner.generate_tool_call_response()
        if tool_results is not None:
            history.append(tool_results)

    if final is None:
        return messages
    if final.stop_reason == "refusal":
        print("\n[The request was declined. Try rephrasing it.]\n")
    elif final.stop_reason == "tool_use":
        print("\n[Stopped after the maximum number of steps. Ask me to continue.]\n")
    for block in final.content:
        if block.type == "text":
            print("\n" + block.text.strip() + "\n")
    return history


def _short(d: dict) -> str:
    s = ", ".join(f"{k}={v!r}" for k, v in d.items())
    return s if len(s) < 90 else s[:87] + "..."


def main() -> None:
    if not FUNNEL_XLSX.exists():
        sys.exit(f"Can't find {FUNNEL_XLSX}")
    client = anthropic.Anthropic()
    messages: list = []

    if len(sys.argv) > 1:
        messages.append({"role": "user", "content": " ".join(sys.argv[1:])})
        run_turn(client, messages)
        return

    print("The Lunch Club sales agent. Ask me to draft outreach, check follow-ups, "
          "or update the funnel. Type 'quit' to exit.\n")
    while True:
        try:
            text = input("you › ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            break
        if text.lower() in {"quit", "exit"}:
            break
        if not text:
            continue
        messages.append({"role": "user", "content": text})
        messages = run_turn(client, messages)


if __name__ == "__main__":
    main()
