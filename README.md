# SNOC IP Monitoring Engine

A Python automation tool I built from scratch to replace manual network monitoring at our NOC. It reads our entire device inventory from Excel, pings every IP across all network segments, calculates uptime statistics, generates a fully formatted multi-sheet Excel report, and emails it to the management team — all without anyone having to do a thing.

This started as a personal project to solve a real problem at work. It's now the primary tool our team uses for shift-end reporting and SLA tracking.

---

## The problem it solves

In a Network Operations Center, keeping track of hundreds of devices across multiple infrastructure segments is genuinely difficult. Before this tool, our process looked like this: an engineer would open a spreadsheet, manually ping a list of devices one by one, note down which ones responded, fill in the status columns by hand, calculate availability percentages manually, and then format everything before sending it to the manager. This happened every shift.

The problems were obvious. It took a long time. People made mistakes. The reports looked different depending on who made them. There was no historical record, so if a manager asked "how much downtime did that server have this week?" nobody had a clean answer. High latency devices were only caught if someone happened to notice.

This tool eliminates all of that. It runs on a schedule, handles everything automatically, and produces consistent, professional output every single time.

---

## What it actually does

### Reading the inventory

The script opens the Excel inventory file we already maintain for our network. This file has multiple sheets — one for networking devices, one for servers, one for VMs, one for storage, and so on. Each sheet has rows with device names, hostnames, IP addresses, VLANs, and locations, but the columns aren't always in the same order across sheets.

I wrote a smart parser that scans each row looking for anything that matches an IP address pattern. Once it finds the IP column, it figures out the surrounding fields by their position relative to the IP rather than by column index. This means the script works with our existing spreadsheet without us needing to reformat it.

Sheets we want to exclude (like the SNMP configuration sheet, which has a completely different format) are listed in a skip list and are passed over entirely.

### Pinging every device

For each IP address found, the script runs a native ping command — using Windows syntax on Windows and Linux syntax on Linux, so it works in both environments without any changes. It sends a configurable number of packets and parses the response to determine three things: whether the device responded at all, the average round-trip latency in milliseconds, and the last line of diagnostic output from the ping command (useful for troubleshooting).

Devices are classified into one of three states:
- **UP** — responded to at least one packet
- **DOWN** — no response, request timed out
- **UNASSIGNED** — the host was unreachable, which usually means the IP exists in our records but hasn't been assigned to a live device yet

If a device responds but its latency is above the configured threshold (100ms by default), it gets flagged separately as a high latency warning rather than being treated as fully healthy.

### Tracking history

After every monitoring run, the results for every device are appended to a persistent `uptime_history.xlsx` file. Each row stores the timestamp, IP address, and status from that check. Over time this file builds up a complete picture of every device's availability history.

This history is what powers the daily uptime and downtime columns in the report. When building the report, the script looks back through the history file, filters to just today's records for each IP, and calculates what percentage of checks came back UP. It also estimates total downtime hours by multiplying the number of DOWN checks by the polling interval.

### Building the Excel report

The report is the main deliverable. It's built using OpenPyXL and uses a custom styling system I wrote to apply consistent formatting, colors, fonts, borders, and cell merges across every sheet. Everything is branded in a dark navy and slate color scheme with color-coded status indicators throughout.

The report contains the following sheets:

**Executive Summary** — This is designed to be the first thing a manager opens. At the top there's an overall health grade for the entire network, displayed as a large letter grade (A through D) alongside the label that goes with it (EXCELLENT STATUS, MOSTLY OPERATIONAL, ATTENTION REQUIRED, or CRITICAL INCIDENT). Below that are five KPI tiles showing total IPs monitored, devices online, active alerts, unassigned blocks, and overall availability percentage.

Under the KPIs is a classification matrix — a table listing every network segment with its totals and availability percentage. Segments with any DOWN devices are highlighted in red so they stand out immediately.

At the bottom is the outage section. If any devices are down, they're listed here grouped into three tiers: Tier 1 for core infrastructure (switches, routers, storage), Tier 2 for virtualisation and data layers (VMs), and Tier 3 for end-user and peripheral access. Each entry shows the source sheet, IP address, device name, hostname, VLAN, diagnostic output, and ownership. If everything is up, this section shows a single green line confirming no outages were detected.

**System Dashboard** — A cleaner version of the segment breakdown, formatted as a proper data table with alternating row colors. Includes a grand totals row at the bottom showing aggregated figures across all segments.

**SLA Report** — A sheet for weekly and monthly uptime and downtime figures per device. The framework is in place and ready to be extended with the full calculation logic.

**Per-category sheets** — One sheet per network segment. Each row is a device with its full details: sequence number, IP address, device module, hostname, VLAN, physical location, status (color coded green/red/grey), latency (color coded by threshold), packet loss percentage, timestamp, daily uptime percentage, estimated daily downtime in hours, raw diagnostic output from the ping, and a remarks column that describes the device's condition in plain language.

### Sending the email

Once the report is saved, the script builds an HTML email and sends it via Gmail's SMTP server. The email is fully styled and contains:

- A header bar with the date, time, and server OS information
- A critical alert block in red (only shown if any devices are down) calling out the number of unreachable devices
- A row of four KPI cards: total IPs, reachable count, down count, and overall availability with grade
- A per-category summary table with UP/DOWN/UNASSIGNED counts, availability percentage, and health grade for each segment
- A short closing note pointing recipients to the attached Excel file for full details

The Excel report is always attached. The entire email is generated dynamically based on the results of that run, so it's always accurate.

---

## Project structure

```
SNOC-IP-Monitoring-Engine/
│
├── main.py                  # Everything lives here — parser, ping engine, report builder, email
├── input_excel.xlsx         # Your device inventory (you provide and maintain this)
├── uptime_history.xlsx      # Auto-created and appended to after every run
├── requirements.txt         # Python dependencies
├── README.md
└── reports/
    └── SNOC_Ping_Report_2025-06-10_14-30.xlsx   # Timestamped output files
```

---

## Setup

You need Python 3.8 or higher. Everything else installs via pip.

```bash
git clone https://github.com/your-username/SNOC-IP-Monitoring-Engine.git
cd SNOC-IP-Monitoring-Engine
pip install openpyxl pandas schedule
```

---

## Configuration

At the top of `main.py` there's a configuration block. Fill these in before running:

```python
INPUT_EXCEL          = "input_excel.xlsx"      # Path to your network inventory Excel file
GMAIL_SENDER         = "your@gmail.com"        # Gmail address the report is sent from
GMAIL_PASSWORD       = "xxxx xxxx xxxx xxxx"   # Gmail App Password — see note below
MANAGER_EMAIL        = ["noc@company.com"]     # List of recipient addresses
REPORT_FOLDER        = "reports"               # Folder where Excel reports are saved
PING_COUNT           = 3                       # Number of ICMP packets sent per device
HIGH_LAT_MS          = 100                     # Latency in ms above which a device is flagged
SKIP_SHEETS          = {"SNMP"}                # Sheet names to skip in the inventory file
CHECK_INTERVAL_HOURS = 3                       # Polling interval, used for downtime hour calculation
```

**About the Gmail password** — Google doesn't allow regular account passwords for SMTP connections. You need to generate an App Password instead. Go to your Google Account → Security → 2-Step Verification → App Passwords, create one for Mail, and paste the 16-character code into `GMAIL_PASSWORD`. Your actual Gmail login password won't work here.

---

## Running it

```bash
python main.py
```

The terminal shows progress as it goes — each network segment is announced, then each IP is printed with its result. When it finishes it saves the report and sends the email. A run across a few hundred IPs typically completes in two to five minutes depending on how many devices time out (each timeout waits up to one second before moving on).

To run on an automatic schedule instead of once manually, uncomment the block at the bottom of the script:

```python
schedule.every(CHECK_INTERVAL_HOURS).hours.do(run_job)
while True:
    schedule.run_pending()
    time.sleep(1)
```

This will run one cycle immediately on startup, then repeat at the configured interval indefinitely.

---

## Health grading

Both the report and the email use a letter grade system to give an at-a-glance indication of network health. Grades are calculated per segment and also for the network overall:

| Availability | Grade | Label |
|---|---|---|
| 95% and above | A | Excellent Status |
| 85% to 94% | B | Mostly Operational |
| 75% to 84% | C | Attention Required |
| Below 75% | D | Critical Incident |

Latency is also graded visually in the per-device sheets:

| Response Time | Indicator |
|---|---|
| 50ms and below | Green — nominal |
| 51ms to 100ms | Amber — elevated |
| Above 100ms | Red — degraded |

---

## Skills demonstrated in this project

For anyone reviewing this as part of a job application, here's a breakdown of what went into building it:

**Python** — the entire tool is written in Python with no external frameworks. Data parsing, regex-based IP extraction, subprocess management for cross-platform ping, file I/O, scheduling, and SMTP email sending are all handled in pure Python.

**Network fundamentals** — understanding of ICMP, ping mechanics, latency, packet loss, VLANs, and how to interpret ping diagnostic output across Windows and Linux.

**Excel automation** — deep use of OpenPyXL to build multi-sheet workbooks from scratch with custom styling, cell merging, conditional formatting logic, freeze panes, column widths, and row heights. No templates — every cell is built programmatically.

**Data handling with Pandas** — reading, filtering, and appending to a persistent Excel-based data store. Date-based filtering for daily statistics.

**Email automation** — constructing multipart MIME emails with HTML bodies and binary attachments, sent via Gmail SMTP with SSL.

**Real-world problem solving** — this wasn't a tutorial project. It was built to replace a manual process at an actual NOC and is actively used by the team.

---

## What I plan to add

- SNMP integration for CPU, memory, and interface-level stats beyond just ping
- A lightweight web dashboard so status is visible in a browser without opening Excel every time
- Microsoft Teams and Slack webhook notifications alongside email
- PostgreSQL or SQLite backend to replace the Excel history file for better query performance
- REST API endpoints so other monitoring tools can pull device status programmatically
- Grafana dashboard fed from the database for visual trend analysis

---

## Dependencies

- `openpyxl` — building and styling Excel workbooks
- `pandas` — managing the uptime history data store
- `schedule` — optional recurring execution

---

## License

MIT. Free to use, modify, and adapt for your own environment.

---

Built by **Prashanth Kumar Sake**  
Network Engineer  · Network Automation  
Bangalore, India
