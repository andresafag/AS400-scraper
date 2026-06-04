**AS400 Hourly Scraper 🚀**

Automates hourly scraping of ASP usage from an IBM i (AS/400) `PROD1` console and writes the results into a turnover document. This tool saves time, reduces human error, and frees up the team to focus on higher-value tasks — perfect to showcase automation skills to recruiters. ✨

**Key Features**
- **Purpose**: Collects ASP usage every hour on the dot and populates a Microsoft Word turnover document. ⏱️📄
- **Reliable**: Waits for the exact sharp hour (e.g., 14:00) before scraping to keep reports aligned with business requirements. ✅
- **Hands-free**: Runs continuously in the background, so you can be on calls or multitask without missing reports. 🤝📞
- **Readable output**: Updates a Word (.docx) file used by operations to deliver daily turnover information. 📝

**How It Works 🔎**
- **Source**: Executes a remote command against the AS/400 host `PROD1` using SSH to run `dspaspbrm` and parse the console output. 🔐
- **Parser**: Extracts names and ASP usage via regular expressions and writes values into specific tables inside the turnover Word document. 🧩
- **Schedule**: Sleeps until the next exact hour, runs the scrape, writes data, then waits for the next hour. 🛌➡️⏰

**Files**
- **Script**: [as400_scraper.py](as400_scraper.py) — main scraper and updater. 🧰

**Quick Start ▶️**
- **Python**: Use Python 3.9+ (recommended) because the script uses the `zoneinfo` module.
- **Install dependencies**: `pip install -r requirements.txt`
- **Environment**:
  - **USER**: SSH username for the AS/400 host. Example: `export USER=mysshuser` 🔐
  - **PASSWORD**: SSH password. Example: `export PASSWORD='s3cr3t'` 🔒

**Configuration ⚙️**
- **Time zone**: The script uses `America/Bogota` by default. Adjust `bogota_tz` inside the script if a different timezone is required. 🌎
- **Output document path**: Edit the `file_path` variable in [as400_scraper.py](as400_scraper.py) to point to your turnover document location. The example path in the script targets a Windows Downloads folder — update this to match your environment (Linux path, network share, or a synced folder). 📂

**Run Locally ▶️**
- Start the scraper in a terminal to run continuously:

```
python as400_scraper.py
```

**Run as a Background Service 🛠️**
- To keep the scraper running across restarts consider one of these options:
  - **systemd service** (recommended on Linux)
  - **tmux / screen** session
  - **Docker container** (package and run with a lightweight image) 🐳

Example `crontab` alternative (runs at startup and keeps process managed elsewhere):

```
@reboot /usr/bin/python /path/to/as400_scraper.py &
```

**Security & Best Practices 🔒**
- Do not hard-code credentials. Use environment variables (as in the script) or a secrets manager. 🗝️
- Limit SSH access to the service account used by the scraper and use key-based auth where possible. The current script uses password auth; consider migrating to SSH keys. 🔑
- Ensure the machine running the script has access to the timezone database (`tzdata`) on Linux. 🕰️

**Dependencies 📦**
- See `requirements.txt` for Python packages required to run the scraper. ✅

**Troubleshooting 🛟**
- If SSH fails: verify `PROD1` hostname, network connectivity, and credentials. 🔍
- If parsing fails: confirm the `dspaspbrm` output format hasn't changed. The script uses a regex to parse usage values; small changes in console output can break parsing. 🧪
- If writing to the Word file fails: ensure the `file_path` is correct, writable, and that any Windows-style paths are adapted for the host running the script. 💾

**Why this matters (for recruiters) ✨**
- **Impact**: Automates a manual hourly operation, reducing turnaround time and human error. 🎯
- **Skills demonstrated**: Remote systems automation, SSH integration, text parsing with regular expressions, programmatic Word file editing, timezone-aware scheduling, and production-ready thinking (background services). 🧠

**Next Steps / Improvements 🔧**
- Add SSH key authentication and remove password text handling.
- Add a configurable JSON/YAML config file to avoid editing the script directly.
- Add structured logging and alerting (email or Slack) on failures.
- Add unit tests for parsing logic.

**Contact / Attribution**
- Script author: Automations & Scripting Team 🙌

Thank you for reviewing — this small automation is a great example of practical impact through scripting. 🌟
# AS400-scraper
