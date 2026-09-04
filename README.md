# Auction Tracker Pro

A modern, offline desktop auction tracking application for football, cricket, and other player-based team auctions.

Auction Tracker Pro is designed for team owners and managers who need to track player sales, team spending, remaining balances, squad limits, target players, and auction activity in real time — without requiring a database, account, server, or internet connection.

---

## Features

### Auction Setup
- Create auctions with any number of teams
- Set budget per team
- Set maximum players per team
- Select your own team
- Optional common base price for all players
- Custom currency symbol
- Full or compact price display (`100,000` / `100K` / `2.5M`)

### Player Management
- Add, edit, and delete players
- Search by player name, role, category, or serial
- Filter by available, sold, unsold, My Team, wishlist, role, or category
- Bulk player actions
- Personal maximum bid for target players
- Wishlist / target-player tracking

### Flexible Excel Import
Auction Tracker Pro does not require a fixed Excel structure.

Only **Player Name** is mandatory. Other fields are optional:
- Position / Role
- Category
- Base Price
- Personal Max Bid
- Wishlist / Target
- Notes

The import process includes:
1. Excel file selection
2. Sheet selection
3. Column mapping
4. Import preview
5. Duplicate handling
6. Final import

Example source file:

| Player | Role | Grade |
|---|---|---|
| Rahim Ahmed | Forward | A |
| Karim Hasan | Goalkeeper | B |

Map it as:

```text
Player -> Player Name
Role   -> Position
Grade  -> Category
```

Unknown columns can be ignored.

If a base-price column is missing:
- the configured common base price is used, or
- the base price defaults to `0`

### Live Auction Tracker
- Searchable player selection designed for 100+ players
- Current-player details
- Team selection with live budget information
- Sold-price entry
- Mark player as sold or unsold
- Personal max-bid warning
- Safe max-bid estimate
- Quick price increment presets

Available quick-price profiles:

```text
Fine       +10     +50      +100      +500
Standard   +500    +1K      +2K       +5K
Thousands  +1K     +5K      +10K      +50K
Lakhs      +100K   +200K    +500K     +1M
Millions   +1M     +5M      +10M      +50M
Huge       +100M   +200M    +500M     +1B
```

### Live Team Tracking
The Live Tracker includes a resizable team-status section.

Available columns:
- Team
- Players Bought
- Player Count
- Latest Player
- Latest Price
- Spent
- Remaining Balance
- Slots Left
- Safe Max Bid
- Average Purchase Price

Columns can be shown or hidden based on preference.

### My Team
- Starting budget
- Total spent
- Remaining balance
- Players bought
- Slots remaining
- Safe max bid
- Squad list
- Personal target players

### History and Recovery
- Complete auction history
- Edit previous sales
- Undo auction entries
- Restore players to available status
- Automatic JSON backups
- Configurable backup retention

### Import / Export
- Flexible Excel import
- Import preview
- Duplicate detection
- Excel results export
- Portable JSON project files

### Interface
- Modern light and dark themes
- Collapsible sidebar
- Resizable Live Tracker sections
- Application icon
- Keyboard shortcuts
- No database required

---

## Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| `Ctrl/Cmd + S` | Save auction |
| `Ctrl/Cmd + O` | Open auction |
| `Ctrl/Cmd + Shift + S` | Save As |
| `Ctrl/Cmd + Z` | Undo last auction entry |
| `Ctrl/Cmd + F` | Focus Live Tracker search |
| `Ctrl/Cmd + Enter` | Mark selected player sold |
| `Ctrl/Cmd + U` | Mark selected player unsold |
| `Ctrl/Cmd + B` | Hide / show sidebar |

---

## Data Storage

Auction Tracker Pro does **not** use MySQL, SQLite, or any external database.

Each auction is stored as a portable JSON file:

```text
My_Auction.json
```

Automatic backups are stored in:

```text
AuctionTrackerBackups/
└── <project_name>/
```

---

## Requirements

- Python 3.10 or newer
- PySide6
- openpyxl

Install dependencies:

```bash
python -m pip install -r requirements.txt
```

---

## Run from Source

### macOS

```bash
python3 --version
python3 -m venv .venv
source .venv/bin/activate

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

python main.py
```

If the project path contains spaces:

```bash
cd "/path/to/Auction Tracker Pro"
```

### Windows

```bat
py --version
py -m venv .venv
.venv\Scripts\activate

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

python main.py
```

---

## Download Ready-to-Use Builds

Prebuilt versions should be published from the repository's **Releases** page.

Recommended release assets:

```text
AuctionTrackerPro-v1.0.0-macOS.dmg
AuctionTrackerPro-v1.0.0-Windows-x64.exe
```

Users who download the DMG or EXE do not need Python installed.

---

## Release Versioning

Use semantic versioning:

```text
v1.0.0  Initial public release
v1.0.1  Bug fixes
v1.1.0  New backward-compatible features
v2.0.0  Major redesign or breaking changes
```

---

## Project Status

Auction Tracker Pro is intended as a lightweight offline utility for:

- football player auctions
- cricket player auctions
- university or club tournaments
- local leagues
- custom team-based auctions


---

### Developer

Made by **Mahfuz Rahman**

[Facebook](https://www.facebook.com/mahfuzrahmannn) ·
[Instagram](https://www.instagram.com/mahfuz_rahmannn/) ·
[GitHub](https://github.com/dev-Mahfuz)
