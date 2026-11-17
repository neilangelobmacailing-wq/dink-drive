# Dink & Drive – Reservation & Operations System

Node.js + Express backend. Every screen (front desk, lobby TV, phones) shares the same bookings,
stored in `data/db.json` on the computer that runs the server.

## 1. Install (one time)
1. Install **Node.js 18 or newer** from https://nodejs.org (LTS version).
2. Open a terminal in this folder and run:
   ```
   npm install
   ```

## 2. Run
```
npm start
```
The terminal prints two addresses:
- `http://localhost:3000` – on the server computer
- `http://192.168.x.x:3000` – open this on the lobby TV / other devices on the same Wi-Fi

Want sample bookings to look around with? Run `npm run demo` once on a fresh install
(Windows PowerShell: `$env:SEED_DEMO=1; npm start`).

## 3. Set real passwords (important)
Default logins are `staff / 1234` and `admin / admin123`. Change them before real use:

**Mac / Linux**
```
STAFF_PASSWORD=your-pin ADMIN_PASSWORD=your-password npm start
```
**Windows PowerShell**
```
$env:STAFF_PASSWORD="your-pin"; $env:ADMIN_PASSWORD="your-password"; npm start
```

Other optional settings: `PORT` (default 3000), `CLUB_TZ` (default `Asia/Manila`), `DATA_DIR`, `BOOKING_DAYS_AHEAD` (default 365), `CLUB_NAME`.

## 4. Lobby TV
Open the server address on the TV's browser, click **Customer TV Mode**. It refreshes itself every
few seconds and rolls over to the new day at midnight (club timezone). It never receives customer
names or phone numbers – only which slots are taken.

## 5. Daily Excel revenue reports (automatic)
Right after midnight (club timezone) the system:
1. Saves the finished day to an Excel file, then
2. Clears that day from the live data so the new day starts clean (future-dated bookings are kept).

Files are saved on the server computer and are never deleted by the system:
```
data/reports/2026-11/Daily-Report-2026-11-15.xlsx   one file per day
data/reports/Revenue-History.xlsx                   one row per day + monthly totals
```
Each daily file has 4 sheets: **Summary** (revenue, unpaid money, payment methods, by court/channel,
items sold), **Bookings** (every transaction), **Counter Sales** and **Paddle Rentals**.

- Staff can download any report from the **Revenue Reports (Excel)** tab, including **TODAY (so far)**.
- Revenue = paid bookings (court + add-ons) + counter sales. Unpaid and cancelled bookings are listed
  but not counted. A booking that was paid and then cancelled is flagged "check refund".
- If the Excel file can't be written (for example it is open and locked), nothing is lost: the day waits in
  `data/pending-reports/` and is saved automatically within a minute of the problem being fixed.
- Keep `data/reports/` backed up – it is your permanent record.

## Backups
Everything lives in the `data/` folder. Copy it to back up; copy it back to restore.
(Stop the server before restoring.)

## Project layout
```
server.js          API, validation, login, storage, end-of-day archive
reports.js         builds the Excel report files
public/index.html  page markup
public/styles.css  custom styles
public/app.js      front-end logic
public/tailwind-config.js
data/              created automatically (db.json, logo.bin, reports/, revenue-history.json)
```
