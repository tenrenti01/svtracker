# SV Tracker — DPWH Bureau of Equipment

A mobile web app for recording where DPWH service vehicles are.

**Live app:** https://tenrenti01.github.io/svtracker/

## What it does
- **Pick a vehicle** by property no., DPWH property code or plate no.
- **One-time update**: use the phone's GPS, or tap a spot on the map and drag the pin to adjust it.
- **Track a trip**: tap *Start trip* and the app records the vehicle's location **every 5 minutes** until you tap *Stop*. It also shows elapsed time, the number of points and the distance in km, and draws the route on the map.
- **Works with weak signal**: records are saved on the phone first and upload when the connection comes back.
- Everything goes to the **svtracker Google Sheet** (Equipment, Location_Log, Trips, Latest_Location).

## Files
| File | What it is |
|---|---|
| `index.html` | The whole app (GitHub Pages serves this) |
| `Code.gs` | Google Apps Script that connects the app to the Google Sheet |

## Connect it to the Google Sheet (one time)
1. Open the **svtracker** Google Sheet → **Extensions → Apps Script**.
2. Paste all of `Code.gs` → **Save**.
3. **Deploy → New deployment → Web app**. Set *Execute as*: **Me** and *Who has access*: **Anyone**.
4. Copy the Web app URL (it ends in `/exec`).
5. In the app, tap the **settings icon** → paste the URL → **Save & test connection**.

Until you do this, the app runs in **demo mode**: it uses sample vehicles and keeps records on the phone.

## Good to know
- Web pages on phones pause when the screen turns off or you switch apps. During a trip, switch on **Keep screen on**. If the page was paused, it records a point as soon as you come back to it.
- Location needs permission. Allow it when the browser asks.
