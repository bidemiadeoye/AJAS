# AJAS — AI-Powered Job Application System

A free, private, local-first job search and application assistant powered by Claude AI and JSearch but built with Python and Flask.

---

## What AJAS does

- Searches live job openings worldwide from LinkedIn, Indeed, Glassdoor, Google Jobs, and more
- Reviews and scores your CV, then tailors it to each specific role
- Writes a fresh cover letter for every application
- Guides you through the application form field by field
- Tracks all your applications and their status in one place

Everything runs on your own computer. No data is sent to any server we control.

---

## Requirements

- Python 3.9 or later — https://python.org/downloads
- A free Claude account — https://claude.ai
- A free RapidAPI account — https://rapidapi.com (for job search)

---

## How to install and launch

### Mac
1. Download and unzip the AJAS package
2. Right-click `AJAS Setup.command` → Open → Open
3. Follow the on-screen setup (takes about 5 minutes)
4. After setup, double-click `Open AJAS` on your Desktop every time

### Windows
1. Download and unzip the AJAS package
2. Double-click `Open AJAS Windows.bat`
3. If Python is not installed, AJAS installs it automatically
4. Follow the on-screen setup in your browser

---

## First-time setup steps

1. Choose where to save your AJAS files
2. Upload your CV (.docx or .pdf)
3. Enter your job preferences (role, location, salary, work type)
4. Create a Claude Project at claude.ai and paste the generated instructions
5. Get a free RapidAPI key and subscribe to JSearch (free plan)
6. Start searching for jobs

---

## File structure

```
AJAS/
├── AJAS Setup.command        ← Mac: run once to set up
├── Open AJAS.command         ← Mac: launch daily
├── Open AJAS Windows.bat     ← Windows: launch
├── Open AJAS Windows.ps1     ← Windows: PowerShell launcher
├── app/
│   ├── ajas_launcher.py      ← Main application
│   ├── templates/            ← UI pages
│   └── static/               ← Icons and assets
└── README.md
```

---

## Privacy and security

- All your data stays on your computer
- CV, preferences, and API keys are stored locally only
- No usage tracking, no analytics, no accounts required from us
- CV text is sent to Anthropic via Claude when tailoring (Anthropic's privacy policy applies)
- RapidAPI key stored in plain text locally — use on personal computers only

---

## License

MIT License — free to use, share, and modify.
