# Event Suite
**Control a big screen from your phone. Perfect for events, competitions, and presentations.**

An open source Taskmaster style event toolkit. One device controls what appears on a big screen. Run timers, show scoreboards, pick random names, display announcements, run through slideshows and more.

## Quick Start
Only Python 3 required (no dependencies).

```bash
python3 server.py          # Default port 8000
python3 server.py 3000     # Custom port
```

This opens the homepage in your browser and prints the local network URLs for both the controller and the display. That's it.

## Taskmaster Mode (Controller + Display)

The main feature: one device (your phone) controls what's shown on another screen (projector/TV).

1. **Run the server** on any machine on your network
2. **Open Display** (`/display.html`) on the big screen — go fullscreen (F11)
3. **Open Controller** (`/controller.html`) on your phone

The controller lets you switch between scenes on the display:

| Scene | What it shows |
|-------|--------------|
| **Welcome** | Event name with animated background |
| **Scoreboard** | Ranked leaderboard with scores |
| **Timer** | Giant stopwatch or countdown |
| **Picker** | Spin-the-wheel random name picker |
| **Schedule** | Event timeline |
| **Message** | Custom announcement in big text |
| **Slides** | Images, videos, and PDFs from `assets/` |
| **Waiting** | "Time Wasted" counter with snarky messages |
| **Blank** | Black screen for transitions |

The green dot on the controller confirms the display is connected.

## Standalone Tools

Each tool also works on its own from the homepage:

- **Team Generator** — create balanced teams with no-go constraints
- **Scoring** — track scores per team across rounds
- **Timer** — stopwatch and countdown with alarm
- **Schedule** — event timeline with current-event tracking
- **Random Picker** — pick names with a spinning animation
- **Notes** — auto-saving scratchpad

## Slides / Presentations

Drop files into the `assets/` folder and they'll appear in the Slides scene on the controller:

- **Images** (PNG, JPG, GIF, SVG, WebP) — shown fullscreen
- **Videos** (MP4, WebM) — shown with play/pause controls from phone
- **PDFs** — embedded viewer

For PowerPoint: export your `.pptx` as images (File → Export → Images) and drop them into `assets/`. They'll sort alphabetically, so naming them `01.png`, `02.png`, etc. works well.

Use the prev/next buttons on your phone, arrow keys on the display, or tap thumbnails. The controller also shows a preview of the next slide.

## Custom Port

```bash
python3 server.py 3000
```

## How It Works

The Python server relays commands between devices using Server-Sent Events (SSE). The controller sends commands via POST, and the display listens for them via an SSE stream. No WebSockets, no database, no accounts, no external dependencies. All event data lives in localStorage.
