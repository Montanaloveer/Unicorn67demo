# Erasmus Inbox Agent

An AI agent for a university Erasmus office. It reads the inbox, files student documents under one clear name, fills in the spreadsheet and drafts replies. Staff stay in charge: nothing is sent, and anything unclear goes to a person.

Built for Challenge 03 at GDG Riga's Community Kickoff 2026 hackathon, with Python, [ADK](https://google.github.io/adk-docs/) and Gemini.

## What it does

| You give it | You get |
|---|---|
| `inbox/*.eml`: emails with attachments like `scan_0042.pdf` and `IMG_2231.jpg` | `filed/Kalnina_Anna/Kalnina_Anna_LearningAgreement_2026-27.pdf` |
| Dates and host universities buried in documents | `students.xlsx`, with one row per student and their missing documents |
| "When is the deadline?" | `drafts/email_04_reply.txt`, answered from `faq.md` and waiting for approval |
| Anything unclear | A row in the `Needs review` sheet explaining why |

## Run it

Needs Python 3.10 or newer.

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Create a free key at [AI Studio](https://aistudio.google.com/apikey). Copy `erasmus_inbox/.env.example` to `erasmus_inbox/.env` and paste the key in. Never commit `.env`.

```bash
adk web
```

Open http://localhost:8000, pick `erasmus_inbox` at the top left and type **Process the inbox**.

| Command | What it does |
|---|---|
| `python scripts/reset_demo.py` | Clears the results so you can run the demo again |
| `python scripts/generate_samples.py` | Rebuilds the fake inbox |
| `python tests/test_agent.py` | Runs the tests, which need no API key |

The agent uses `gemini-flash-lite-latest` by default. To try another model, set `ERASMUS_MODEL` in `.env`.

## The sample inbox

All 12 emails are fake, and each one tests something different:

| Email | Situation | Expected result |
|---|---|---|
| 01 | Signed agreement, `scan_0042.pdf` | Filed, row added |
| 02 | No subject, no body, photo of insurance | Photo read, filed |
| 03 | `document(3).pdf`, insurance | Filed |
| 04 | "When is the deadline?" | Draft reply from the FAQ |
| 05 | Transcript with a messy file name | Filed |
| 06 | Certificate of arrival | Filed |
| 07 | Scanned PDF with no text layer | Scan read, filed |
| 08 | Asks for a bigger grant | Passed to staff, because the FAQ doesn't cover it |
| 09 | Sender attaches another student's agreement | Passed to staff |
| 10 | Agreement with changed dates | New version kept, old dates stay, staff told |
| 11 | Tries to give the AI instructions | Ignored, passed to staff |
| 12 | Gym receipt | Passed to staff |

## How it stays safe

The model reads the emails, but plain Python checks what it asks to do:

- **File names come from checked values**, so nothing in an email can choose where a file goes.
- **The document must name the student, and the sender must be that student**, or filing is refused. This stops one student's document going in another's folder.
- **The document must say what it is.** A file is only saved as, say, a Certificate of Arrival if its text says so, which also means documents in other languages go to staff.
- **Dates and universities must appear in that email.** Anything else is left blank, the agent gets one chance to correct a misread value, and whatever stays blank is marked in the report. Insurance dates can't be used as stay dates.
- **The results table comes from what the tools really did**, not from the model's summary, so an email that wasn't handled shows as NOT HANDLED.
- **Existing spreadsheet values are never overwritten.** A difference goes to staff instead.
- **Emails are treated as data.** The agent is told not to obey them, and email text can't run as an Excel formula.
- **The agent cannot send email.** It has no tool for it.

Because the model can still make mistakes that pass these checks, staff should review its results.

## Declaration

Required by the hackathon rules:

- The project code was started at about 14:15 on 2 November 2026, before the official 14:30 start.
- Apart from that early start, no code from before the event day was reused. The only outside code is the libraries in `requirements.txt`.
- The code was written with help from an AI coding assistant (Claude Code). The agent itself runs on Gemini.
- All emails, people and FAQ answers are invented. No real student data is used.
